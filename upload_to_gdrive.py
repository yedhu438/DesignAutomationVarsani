"""
Upload today's PSD exports to Google Drive.

First-time setup:
  1. Go to https://console.cloud.google.com/
  2. Create a project (or select existing)
  3. Enable "Google Drive API"
  4. Go to APIs & Services > Credentials > Create Credentials > OAuth client ID
  5. Application type: Desktop app
  6. Download the JSON and save it as:
       W:\VarsaniAutomation\credentials.json
  7. Run this script — a browser window will open to authorise once.
     After that, token.json is saved and no browser is needed again.
"""

import os, sys, time
from pathlib import Path
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────────────────
LOCAL_OUTPUT   = r"C:\Varsany\Output\DTF_Excel"
GDRIVE_FOLDER  = "1ZObOngMUAQo519ThI0vEckR4waKp7bsj"   # target Google Drive folder ID
CREDENTIALS    = r"C:\Varsany\credentials.json"
TOKEN_FILE     = r"C:\Varsany\token.json"
SCOPES         = ["https://www.googleapis.com/auth/drive.file"]

# Upload only today's output by default
UPLOAD_DATE    = datetime.now().strftime("%Y-%m-%d")   # e.g. 2026-04-14

# ── Auth ──────────────────────────────────────────────────────────────────────
def get_service():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS):
                print("ERROR: credentials.json not found.")
                print("See the setup instructions at the top of this file.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
        print("Authenticated and token saved.")

    return build("drive", "v3", credentials=creds)

# ── Drive helpers ─────────────────────────────────────────────────────────────
_folder_cache = {}   # local_path -> drive_folder_id

def get_or_create_folder(service, name, parent_id):
    """Return existing or newly-created folder ID under parent_id."""
    cache_key = (parent_id, name)
    if cache_key in _folder_cache:
        return _folder_cache[cache_key]

    # Search for existing folder
    q = (f"name='{name}' and mimeType='application/vnd.google-apps.folder'"
         f" and '{parent_id}' in parents and trashed=false")
    results = service.files().list(q=q, fields="files(id,name)").execute()
    files   = results.get("files", [])
    if files:
        fid = files[0]["id"]
    else:
        meta = {"name": name, "mimeType": "application/vnd.google-apps.folder",
                "parents": [parent_id]}
        f   = service.files().create(body=meta, fields="id").execute()
        fid = f["id"]
        print(f"  Created folder: {name}")

    _folder_cache[cache_key] = fid
    return fid

def file_exists(service, name, parent_id):
    """Check if a file with this name already exists in the folder."""
    q = f"name='{name}' and '{parent_id}' in parents and trashed=false"
    r = service.files().list(q=q, fields="files(id)").execute()
    return len(r.get("files", [])) > 0

def upload_file(service, local_path, parent_id):
    from googleapiclient.http import MediaFileUpload
    name = os.path.basename(local_path)
    if file_exists(service, name, parent_id):
        print(f"  SKIP (already exists): {name}")
        return
    size_mb = os.path.getsize(local_path) / 1024 / 1024
    print(f"  Uploading {name}  ({size_mb:.1f} MB) ...", end=" ", flush=True)
    t0 = time.time()
    media = MediaFileUpload(local_path, resumable=True)
    service.files().create(
        body={"name": name, "parents": [parent_id]},
        media_body=media,
        fields="id"
    ).execute()
    print(f"done in {time.time()-t0:.1f}s")

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    upload_dir = os.path.join(LOCAL_OUTPUT, UPLOAD_DATE)
    if not os.path.isdir(upload_dir):
        print(f"No output folder found for {UPLOAD_DATE}: {upload_dir}")
        sys.exit(1)

    # Collect all PSD files with their relative subfolder path
    files_to_upload = []
    for root, dirs, files in os.walk(upload_dir):
        for fname in files:
            if fname.lower().endswith((".psd", ".psb")):
                files_to_upload.append(os.path.join(root, fname))

    if not files_to_upload:
        print(f"No PSD/PSB files found in {upload_dir}")
        sys.exit(0)

    print(f"Found {len(files_to_upload)} files to upload from {upload_dir}")
    print("Authenticating with Google Drive...")
    service = get_service()
    print("Authenticated.")
    print()

    # Create date folder inside the target Google Drive folder
    date_folder_id = get_or_create_folder(service, UPLOAD_DATE, GDRIVE_FOLDER)

    total    = len(files_to_upload)
    uploaded = 0
    skipped  = 0

    for i, local_path in enumerate(sorted(files_to_upload), 1):
        # Reconstruct relative subfolder path (relative to upload_dir)
        rel       = os.path.relpath(local_path, upload_dir)
        parts     = Path(rel).parts   # e.g. ('DTF Front', 'black', 'order.psd')
        subfolder = parts[:-1]        # directory parts only
        fname     = parts[-1]

        # Build (or reuse) the matching folder chain in Drive
        parent_id = date_folder_id
        for part in subfolder:
            parent_id = get_or_create_folder(service, part, parent_id)

        print(f"[{i}/{total}] {os.path.join(*subfolder, fname) if subfolder else fname}")
        try:
            if file_exists(service, fname, parent_id):
                print(f"  SKIP (already exists)")
                skipped += 1
            else:
                upload_file(service, local_path, parent_id)
                uploaded += 1
        except Exception as e:
            print(f"  ERROR: {e}")

    print()
    print(f"Done.  Uploaded: {uploaded}  |  Skipped: {skipped}  |  Total: {total}")
    print(f"View at: https://drive.google.com/drive/folders/{GDRIVE_FOLDER}")

if __name__ == "__main__":
    main()
