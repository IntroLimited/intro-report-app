import json
import os
import urllib.request
import urllib.error
from http.server import BaseHTTPRequestHandler

NOTION_API_KEY = os.environ.get("NOTION_API_KEY", "")
TRELLO_DATABASE_ID = os.environ.get("TRELLO_DATABASE_ID", "b3e6ab7470994b3690c110b728f6593b")

def notion_request(method, path, payload=None):
    url = f"https://api.notion.com/v1{path}"
    headers = {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28"
    }
    data = json.dumps(payload).encode() if payload else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())

def get_live_roles():
    """Fetch all Live searches from the Trello/pipeline database."""
    roles = []
    cursor = None

    while True:
        payload = {
            "filter": {
                "property": "Status",
                "status": {"equals": "Live"}
            },
            "sorts": [{"property": "Client", "direction": "ascending"}],
            "page_size": 100
        }
        if cursor:
            payload["start_cursor"] = cursor

        data = notion_request("POST", f"/databases/{TRELLO_DATABASE_ID}/query", payload)

        for page in data.get("results", []):
            props = page.get("properties", {})

            # Get role name from title
            name_prop = props.get("Name", {})
            role_name = "".join(
                t.get("plain_text", "") for t in name_prop.get("title", [])
            ).strip()

            # Get client from select field
            client_prop = props.get("Client", {})
            client_select = client_prop.get("select")
            client_name = client_select.get("name", "").strip() if client_select else ""

            if role_name and client_name:
                roles.append({
                    "client": client_name,
                    "role": role_name
                })

        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")

    # Sort by client then role
    roles.sort(key=lambda x: (x["client"].lower(), x["role"].lower()))
    return roles

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        path = self.path.split('?')[0]
        if path in ('/', '/index.html'):
            self._serve_file(
                os.path.join(os.path.dirname(__file__), '..', 'public', 'index.html'),
                'text/html; charset=utf-8'
            )
        else:
            self.send_error(404)

    def _serve_file(self, filepath, content_type):
        try:
            with open(filepath, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except FileNotFoundError:
            self.send_error(404)

    def _json(self, status, data):
        body = json.dumps(data).encode()
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args): pass
