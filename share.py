import requests
import sys
import os

SERVER_IP = "192.168.43.55"   # personal laptop IP
PORT = 8000

file_path = sys.argv[1]
filename = os.path.basename(file_path)

with open(file_path, "rb") as f:
    r = requests.post(
        f"http://{SERVER_IP}:{PORT}",
        data=f,
        headers={"X-Filename": filename}
    )

print("Upload status:", r.status_code)
