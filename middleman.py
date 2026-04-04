import sys
import os
import json
import urllib.request
import urllib.error
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

try:
    import geopandas as gpd
    from shapely.geometry import shape
    import pyproj
except ImportError:
    print("Missing dependencies. Please install: pip install geopandas shapely pyproj pyogrio")
    sys.exit(1)

# To handle login issues, we allow the user to provide a session cookie via a config file.
CONFIG_FILE = "sartopo_config.json"

def get_session_cookie():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                return data.get('session_cookie', '')
        except:
            pass
    return ''

class ProxyHandler(BaseHTTPRequestHandler):
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
                if e.code in [401, 403]:
                    error_msg["details"] = "Login required. Map is private and no valid session cookie was provided."
                self.wfile.write(json.dumps(error_msg).encode('utf-8'))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e), "status": 500}).encode('utf-8'))
                return

            # For the frontend to use its existing showSarTopoShapesPopup, 
            # we need to return the raw data!
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(data).encode('utf-8'))
        else:
            self.send_response(404)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()

def run(port=5050):
    server_address = ('', port)
    httpd = HTTPServer(server_address, ProxyHandler)
    print(f"Middleman proxy server running on port {port}...")
    httpd.serve_forever()

if __name__ == '__main__':
    run()
