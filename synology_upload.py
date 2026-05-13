"""
synology_upload.py — Varsany NAS Integration
Upload PSD files to Synology NAS via FileStation API.
Credentials verified working: varsany_api / Varsany2026
"""
import os, requests, urllib3
urllib3.disable_warnings()

NAS_HOST   = "192.168.0.113"
NAS_PORT   = 5001
NAS_USER   = "varsany_api"
NAS_PASS   = "Varsany2026"
NAS_FOLDER = "/Automated"

class SynologyUploader:
    def __init__(self):
        self.sid  = None
        self.base = f"https://{NAS_HOST}:{NAS_PORT}/webapi"
        self._login()

    def _login(self):
        try:
            r = requests.get(f"{self.base}/auth.cgi", params={
                "api":"SYNO.API.Auth","version":"3","method":"login",
                "account":NAS_USER,"passwd":NAS_PASS,
                "session":"FileStation","format":"sid"
            }, timeout=10, verify=False)
            d = r.json()
            if d.get("success"):
                self.sid = d["data"]["sid"]
                print("[Synology] Connected OK")
            else:
                print(f"[Synology] Login failed: {d}")
        except Exception as e:
            print(f"[Synology] Error: {e}")

    def upload(self, local_path, sub_folder=""):
        if not self.sid:
            print("[Synology] Not connected - skipping upload")
            return False
        if not os.path.exists(local_path):
            print(f"[Synology] File not found: {local_path}")
            return False
        nas_path = f"{NAS_FOLDER}/{sub_folder}" if sub_folder else NAS_FOLDER
        filename = os.path.basename(local_path)
        try:
            with open(local_path, "rb") as f:
                r = requests.post(f"{self.base}/entry.cgi", params={
                    "api":"SYNO.FileStation.Upload","version":"2","method":"upload",
                    "_sid":self.sid,"path":nas_path,
                    "create_parents":"true","overwrite":"true",
                }, files={"file":(filename, f, "application/octet-stream")},
                timeout=120, verify=False)
            d = r.json()
            if d.get("success"):
                size_mb = os.path.getsize(local_path) / (1024*1024)
                print(f"[Synology] Uploaded: {nas_path}/{filename} ({size_mb:.1f} MB)")
                return True
            else:
                print(f"[Synology] Upload failed: {d}")
                return False
        except Exception as e:
            print(f"[Synology] Upload error: {e}")
            return False

    def logout(self):
        if self.sid:
            requests.get(f"{self.base}/auth.cgi", params={
                "api":"SYNO.API.Auth","version":"3","method":"logout",
                "session":"FileStation","_sid":self.sid
            }, timeout=5, verify=False)
            self.sid = None
            print("[Synology] Logged out")
