import http.server
import socketserver
import socket

PORT = 8000

# Get local IP address
hostname = socket.gethostname()
local_ip = socket.gethostbyname(hostname)

Handler = http.server.SimpleHTTPRequestHandler

with socketserver.TCPServer(("", PORT), Handler) as httpd:
    print("====================================")
    print("📂 File sharing server started")
    print(f"🌐 Access from other device:")
    print(f"👉 http://{local_ip}:{PORT}")
    print("❌ Press CTRL + C to stop")
    print("====================================")
    httpd.serve_forever()
