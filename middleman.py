import sys
import os
import json
import urllib.request
import urllib.error
import socket
from http.server import HTTPServer, BaseHTTPRequestHandler, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# Middleman server/sync

# To handle login issues, we allow the user to provide a session cookie via a config file.
CONFIG_FILE = "sartopo_config.json"
BUNDLE_FILE = "shared_bundle.json"
ACTIVE_USERS_FILE = "active_users.json"

def get_session_cookie():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                return data.get('session_cookie', '')
        except:
            pass
    return ''

def get_bundle():
    if os.path.exists(BUNDLE_FILE):
        try:
            with open(BUNDLE_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return None

def save_bundle(data):
    with open(BUNDLE_FILE, 'w') as f:
        json.dump(data, f)

def get_active_users():
    if os.path.exists(ACTIVE_USERS_FILE):
        try:
            with open(ACTIVE_USERS_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}

def save_active_users(users):
    with open(ACTIVE_USERS_FILE, 'w') as f:
        json.dump(users, f)

class ProxyHandler(SimpleHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200, "ok")
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header("Access-Control-Allow-Headers", "X-Requested-With, Content-type")
        self.end_headers()

    def do_GET(self):
        parsed_path = urlparse(self.path)
        if parsed_path.path == '/api/proxy':
            qs = parse_qs(parsed_path.query)
            map_id = qs.get('mapId', [''])[0]
            domain = qs.get('domain', ['sartopo.com'])[0]

            if not map_id:
                self.send_error(400, "Missing mapId")
                return

            # Fetch from SarTopo
            url = f"https://{domain}/api/v1/map/{map_id}/features"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 SARWebTheory/1.0'})
            
            cookie = get_session_cookie()
            if cookie:
                req.add_header('Cookie', f'id={cookie}')

            try:
                with urllib.request.urlopen(req) as response:
                    data = json.loads(response.read().decode('utf-8'))
            except urllib.error.HTTPError as e:
                self.send_response(e.code)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                error_msg = {"error": f"HTTP Error {e.code}: {e.reason}", "status": e.code}
                if e.code in [401, 403, 404]:
                    error_msg["details"] = "Login required or map not found. Map is private and no valid session cookie was provided."
                self.wfile.write(json.dumps(error_msg).encode('utf-8'))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e), "status": 500}).encode('utf-8'))
                return

            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(data).encode('utf-8'))
        
        elif parsed_path.path == '/api/bundle':
            bundle = get_bundle()
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(bundle).encode('utf-8'))

        elif parsed_path.path == '/api/active-user':
            qs = parse_qs(parsed_path.query)
            pin = qs.get('pin', [''])[0]
            users = get_active_users()
            device_id = users.get(pin, None)
            
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"deviceId": device_id}).encode('utf-8'))
        else:
            return super().do_GET()

    def do_POST(self):
        parsed_path = urlparse(self.path)
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        
        if parsed_path.path == '/api/bundle':
            data = json.loads(post_data.decode('utf-8'))
            save_bundle(data)
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok"}).encode('utf-8'))
            
        elif parsed_path.path == '/api/active-user':
            data = json.loads(post_data.decode('utf-8'))
            pin = data.get('pin')
            device_id = data.get('deviceId')
            
            if pin and device_id:
                users = get_active_users()
                users[pin] = device_id
                save_active_users(users)
                
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok"}).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # doesn't even have to be reachable
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    return IP

def run(port=5050):
    server_address = ('', port)
    httpd = HTTPServer(server_address, ProxyHandler)
    local_ip = get_ip()
    print(f"Middleman server/sync running on port {port}...")
    print(f"Local Access: http://localhost:{port}")
    print(f"Network Access (from phone): http://{local_ip}:{port}")
    print(f"To sync your phone, use the Network Access URL in the app settings.")
    httpd.serve_forever()

if __name__ == '__main__':
    run()
