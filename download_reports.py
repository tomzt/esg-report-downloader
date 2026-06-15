"""
download_reports.py  requests-only version (no Playwright/browser)
Calls SET API directly with session cookies from requests.Session()
"""
import csv
import io
import json
import os
import time
import zipfile
from pathlib import Path

import openpyxl
import requests


#  Google Drive 
def get_drive_service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    creds_json = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        raise RuntimeError("GDRIVE_SERVICE_ACCOUNT_JSON not set")
    info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"])
    return build("drive", "v3", credentials=creds)


def get_or_create_folder(service, name, parent_id):
    q = (f"name='{name}' and mimeType='application/vnd.google-apps.folder'"
         f" and '{parent_id}' in parents and trashed=false")
    res = service.files().list(q=q, fields="files(id)").execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id]}
    return service.files().create(body=meta, fields="id").execute()["id"]


def upload_to_drive(service, local_path, filename, folder_id):
    from googleapiclient.http import MediaFileUpload
    media = MediaFileUpload(local_path, resumable=False)
    meta = {"name": filename, "parents": [folder_id]}
    service.files().create(body=meta, media_body=media, fields="id").execute()
    print(f"   Uploaded {filename}")


#  Checkpoint 
LOG_FILE = "download_log.csv"
LOG_FIELDS = ["symbol", "doc_type", "language", "status", "filename", "note"]


def load_done():
    done = set()
    if Path(LOG_FILE).exists():
        with open(LOG_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row["status"] == "success":
                    done.add((row["symbol"], row["doc_type"], row["language"]))
    return done


def log_result(symbol, doc_type, language, status, filename="", note=""):
    write_header = not Path(LOG_FILE).exists()
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        if write_header:
            w.writeheader()
        w.writerow({"symbol": symbol, "doc_type": doc_type, "language": language,
                    "status": status, "filename": filename, "note": note})


#  Read Excel 
def read_companies(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    companies = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[1]:
            continue
        companies.append({
            "order": row[0],
            "symbol": str(row[1]).strip(),
            "name_en": str(row[2]).strip() if row[2] else "",
        })
    return companies


#  SET API 
def make_session():
    """Create requests session with SET cookies."""
    s = requests.Session()
    s.headers.update({
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"),
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "th,en-US;q=0.9",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://www.set.or.th/th/home",
    })
    s.get("https://www.set.or.th/th/home", timeout=20)
    return s


def get_report_url(session, symbol, endpoint, lang, target_year=2024):
    """Return (url, note) for the target fiscal year from SET API."""
    api = f"https://www.set.or.th/api/set/company/{symbol}/report/{endpoint}?lang={lang}"
    try:
        r = session.get(api, timeout=15)
        if not r.ok:
            return None, f"API {r.status_code}"
        items = r.json()
        if not items:
            return None, "No items from API"
        for item in items:
            if item.get("year") == target_year:
                return item["url"], ""
        # Fallback: most recent
        return items[0]["url"], f"year={items[0].get('year')} (fallback)"
    except Exception as e:
        return None, str(e)[:80]


def download_file(session, file_url, out_path):
    """Download ZIP or PDF; extract PDF from ZIP if needed. Returns (ok, filename)."""
    try:
        r = session.get(file_url, timeout=60)
        if not r.ok:
            return False, f"HTTP {r.status_code}"
        content = r.content
        if len(content) < 500:
            return False, "File too small"

        # If ZIP, extract PDF
        if file_url.lower().endswith(".zip") or content[:2] == b"PK":
            try:
                z = zipfile.ZipFile(io.BytesIO(content))
                pdf_names = [n for n in z.namelist() if n.lower().endswith(".pdf")]
                if pdf_names:
                    pdf_data = z.read(pdf_names[0])
                    out_path.write_bytes(pdf_data)
                    return True, out_path.name
                else:
                    # No PDF inside, save as ZIP
                    zip_path = out_path.with_suffix(".zip")
                    zip_path.write_bytes(content)
                    return True, zip_path.name
            except Exception as e:
                return False, f"ZIP error: {e}"
        else:
            out_path.write_bytes(content)
            return True, out_path.name
    except Exception as e:
        return False, str(e)[:80]


#  Process one company 
def process_company(session, service, company, done, root_folder_id, tmp_dir):
    symbol = company["symbol"]
    symbol_dir = tmp_dir / symbol
    symbol_dir.mkdir(exist_ok=True)
    company_folder_id = get_or_create_folder(service, symbol, root_folder_id)

    tasks = [
        ("56-1 One Report", "th", "one"),
        ("56-1 One Report", "en", "one"),
        ("Annual Report",   "th", "annual"),
        ("Annual Report",   "en", "annual"),
    ]

    for doc_type, lang, endpoint in tasks:
        key = (symbol, doc_type, lang)
        if key in done:
            print(f"  [{symbol}] {doc_type} ({lang}) - skipped (done)")
            continue

        print(f"  [{symbol}] {doc_type} ({lang})...")

        doc_url, note = get_report_url(session, symbol, endpoint, lang, target_year=2024)
        if not doc_url:
            print(f"     {note}")
            log_result(symbol, doc_type, lang, "failed", note=note)
            continue

        safe_type = doc_type.replace(" ", "").replace("-", "")
        out_path = symbol_dir / f"{symbol}_{safe_type}_{lang.upper()}.pdf"
        ok, result = download_file(session, doc_url, out_path)

        if ok:
            actual = symbol_dir / result
            try:
                upload_to_drive(service, str(actual), result, company_folder_id)
                log_result(symbol, doc_type, lang, "success", result)
            except Exception as e:
                log_result(symbol, doc_type, lang, "upload_error", result, str(e)[:80])
        else:
            print(f"     {result}")
            log_result(symbol, doc_type, lang, "failed", note=result)

        time.sleep(0.5)

    import shutil
    shutil.rmtree(symbol_dir, ignore_errors=True)


#  Main 
def main():
    xlsx_path = os.environ.get("EXCEL_PATH", "companies.xlsx")
    root_folder_id = os.environ.get("GDRIVE_FOLDER_ID", "")
    batch_start = int(os.environ.get("BATCH_START", "0"))
    batch_end = int(os.environ.get("BATCH_END", "10"))

    if not root_folder_id:
        raise RuntimeError("GDRIVE_FOLDER_ID not set")

    all_companies = read_companies(xlsx_path)
    companies = all_companies[batch_start:batch_end]
    done = load_done()

    print(f"Processing {len(companies)} companies (batch {batch_start}-{batch_end})")
    print(f"Already done: {len(done)} tasks")

    service = get_drive_service()
    tmp_dir = Path("tmp_downloads")
    tmp_dir.mkdir(exist_ok=True)
    session = make_session()

    for i, company in enumerate(companies):
        process_company(session, service, company, done, root_folder_id, tmp_dir)
        # Refresh session every 30 companies to avoid cookie expiry
        if (i + 1) % 30 == 0:
            print("  [session] Refreshing...")
            session = make_session()

    if Path(LOG_FILE).exists():
        try:
            upload_to_drive(service, LOG_FILE, LOG_FILE, root_folder_id)
        except Exception:
            pass

    print("\nDone!")


if __name__ == "__main__":
    main()
