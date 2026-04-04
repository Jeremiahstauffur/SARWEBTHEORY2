import urllib.request
import json

map_id = "B61TTKN"
url = f"https://sartopo.com/api/v1/map/{map_id}/features"

try:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as response:
        data = json.loads(response.read().decode('utf-8'))
        print("Success! Features found:", len(data.get('features', [])))
except Exception as e:
    print("Error:", e)
