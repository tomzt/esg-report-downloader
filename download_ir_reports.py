#!/usr/bin/env python3
"""
download_ir_reports.py
──────────────────────
Supplementary downloader for 56-1 One Reports that are missing from the
SET API (ไม่มีในระบบ SET / ไม่พบในระบบ).

Sources:
  • Direct PDF URLs from company IR websites
  • SEC Thailand ZIP archives (extracts the matching TH/EN PDF inside)

Output mirrors download_reports.py exactly:
  • Filename  : {SYMBOL}_56-1OneReport_{TH|EN}.pdf
  • Drive path: root_folder / {SYMBOL} / {filename}
  • Log       : download_log.csv  (appended)
  • Status    : esg_report_status.csv  (updated in-place)

Required environment variables (same as download_reports.py):
  GDRIVE_OAUTH_CLIENT_ID
  GDRIVE_OAUTH_CLIENT_SECRET
  GDRIVE_OAUTH_REFRESH_TOKEN
"""

import io
import os
import csv
import sys
import zipfile
import logging
from datetime import datetime
from pathlib import Path

import requests
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ──────────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────────

ROOT_FOLDER_ID = "1khVs15lhm71uL5f8HQHmveo9gpXsoci3"
DOC_TYPE_LABEL = "56-1OneReport"          # matches safe_type in download_reports.py
LOG_FILE       = "download_log.csv"
STATUS_FILE    = "esg_report_status.csv"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/pdf,application/octet-stream,*/*",
}

# ──────────────────────────────────────────────────────────────────────────────
# URL inventory
# ──────────────────────────────────────────────────────────────────────────────
#
# Format: (symbol, lang, url, source_type)
#   source_type "direct"  → download PDF directly from URL
#   source_type "sec_zip" → download SEC ZIP and extract the TH or EN PDF
#
# Companies confirmed as having NO FY2024 filing — omitted intentionally:
#   PRIN EN  : no 2024 English One Report published anywhere
#   TGE  EN  : SEC shows no FY2024 filing (gap between 2023 and 2025)
#   THANA TH : no FY2024 One Report filed (SEC confirms only FY2025)
#   THANA EN : same as above
#
# SC TH note: scasset.com is fully JS-rendered; the URL below is the best
# candidate found via Google index. Verify the document covers FY2024 before
# treating the upload as authoritative.
#
REPORTS = [
    # ── Direct PDF downloads ──────────────────────────────────────────────────

    # S (Singha Estate) — both languages missing from SET API
    ("S",      "TH", "https://hub.optiwise.io/th/documents/155838/s-or2024-th.pdf",    "direct"),
    ("S",      "EN", "https://hub.optiwise.io/en/documents/155838/s-or2024-en.pdf",    "direct"),

    # SAT (Somboon Advance Technology) — TH missing from SET API
    ("SAT",    "TH",
     "https://sat.listedcompany.com/misc/one-report/20250325-sat-one-report-2024-th.pdf",
     "direct"),

    # BKIH (BKI Holdings) — EN missing from SET API
    ("BKIH",   "EN", "https://www.bkiholdings.com/files-fm/pdf/onereport2024_EN.pdf",  "direct"),

    # SEAFCO — EN missing from SET API
    ("SEAFCO", "EN",
     "https://seafco.listedcompany.com/misc/or/20250428-seafco-or2024-en.pdf",
     "direct"),

    # CPL Group — EN missing from SET API
    # (apostrophe in path is intentional; requests handles it correctly)
    ("CPL",    "EN",
     "https://www.cpl.co.th/Shareholder's-Meeting/Invitation-EN/2025/"
     "Attachment%202%20Form%2056-1%20One%20Report%20and%20Financial%20Statements.pdf",
     "direct"),

    # CHAO (Chaosua Foods) — EN missing from SET API
    ("CHAO",   "EN",
     "https://chao.listedcompany.com/misc/ar/20250324-chao-or2024-en.pdf",
     "direct"),

    # PCC (Precise Corporation) — EN missing from SET API (TH already downloaded)
    ("PCC",    "EN",
     "https://hub.optiwise.io/en/documents/157792/pcc-56-1-one-report-2024-en.pdf",
     "direct"),

    # RBF (R&B Food Supply) — EN missing from SET API
    ("RBF",    "EN",
     "https://hub.optiwise.io/en/documents/155190/rbf-or2024-en.pdf",
     "direct"),

    # SC Asset — TH missing from SET API
    # ⚠️  VERIFY: page title found via Google says "2568"; confirm this is FY2024 data.
    ("SC",     "TH",
     "https://www.scasset.com/stocks/sustainability_list/o0x0/jv/sa/"
     "hxbijvsav7z/E-One_Report_TH.pdf",
     "direct"),

    # XO (Exotic Food) — EN missing from SET API
    # irplus.in.th publishes a single bilingual PDF for TH and EN.
    ("XO",     "EN",
     "https://www.irplus.in.th/Listed/XO/annual/an_xo_2024.pdf",
     "direct"),

    # ── SEC ZIP downloads ─────────────────────────────────────────────────────
    # The ZIP contains both TH and EN PDFs; we extract only the one we need.

    # TEAMG — TH missing from SET API; ZIP has the Thai PDF inside
    ("TEAMG",  "TH",
     "https://market.sec.or.th/public/idisc/Download"
     "?FILEID=dat/f56/1439ONE310320251935530114E.zip",
     "sec_zip"),

    # SO (Siamrajathanee) — EN missing from SET API
    ("SO",     "EN",
     "https://market.sec.or.th/public/idisc/Download"
     "?FILEID=dat/f56/1602ONE240320251554580833E.zip",
     "sec_zip"),

    # TBN Corporation — EN missing from SET API
    ("TBN",    "EN",
     "https://market.sec.or.th/public/idisc/Download"
     "?FILEID=dat/f56/1759ONE260320252005430254E.zip",
     "sec_zip"),
]

# ──────────────────────────────────────────────────────────────────────────────
# Google Drive helpers  (same logic as download_reports.py)
# ──────────────────────────────────────────────────────────────────────────────

def _get_drive_service():
    creds = Credentials(
        token=None,
        refresh_token=os.environ["GDRIVE_OAUTH_REFRESH_TOKEN"],
        client_id=os.environ["GDRIVE_OAUTH_CLIENT_ID"],
        client_secret=os.environ["GDRIVE_OAUTH_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    creds.refresh(Request())
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def get_or_create_folder(service, name: str, parent_id: str) -> str:
    """Return Drive folder id, creating it under parent_id if needed."""
    q = (
        f"name='{name}' and mimeType='application/vnd.google-apps.folder'"
        f" and '{parent_id}' in parents and trashed=false"
    )
    res = service.files().list(q=q, fields="files(id)").execute()
    hits = res.get("files", [])
    if hits:
        return hits[0]["id"]
    meta = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    folder = service.files().create(body=meta, fields="id").execute()
    return folder["id"]


def file_exists_in_drive(service, filename: str, folder_id: str) -> bool:
    q = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
    res = service.files().list(q=q, fields="files(id)").execute()
    return bool(res.get("files"))


def upload_to_drive(service, data: bytes, filename: str, folder_id: str) -> str:
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype="application/pdf", resumable=False)
    meta  = {"name": filename, "parents": [folder_id]}
    f     = service.files().create(body=meta, media_body=media, fields="id").execute()
    return f["id"]

# ──────────────────────────────────────────────────────────────────────────────
# Download helpers
# ──────────────────────────────────────────────────────────────────────────────

def download_direct(url: str) -> bytes:
    """Download a PDF directly. Raises on HTTP error or suspiciously small body."""
    resp = requests.get(url, headers=HEADERS, timeout=60, allow_redirects=True)
    resp.raise_for_status()
    if len(resp.content) < 1_024:
        raise ValueError(
            f"Response body too small ({len(resp.content)} B) — likely a 404 error page"
        )
    return resp.content


def _pick_pdf_from_zip(zf: zipfile.ZipFile, lang: str) -> tuple[bytes, str]:
    """
    Extract the best-matching PDF from a SEC ZIP.
    Preference order:
      1. Filename contains the target language marker (TH / EN)
      2. Only one PDF in the archive → use it regardless
      3. First PDF alphabetically
    Returns (pdf_bytes, original_filename).
    """
    pdf_names = [n for n in zf.namelist() if n.lower().endswith(".pdf")]
    if not pdf_names:
        raise ValueError("No PDF files found inside ZIP")

    lang_markers = {
        "TH": ["_th", "-th", " th", "(th)", "ไทย", "thai"],
        "EN": ["_en", "-en", " en", "(en)", "english", "eng"],
    }.get(lang.upper(), [])

    for name in pdf_names:
        if any(m in name.lower() for m in lang_markers):
            log.info(f"    ZIP match for {lang}: {name}")
            return zf.read(name), name

    # Fallback
    chosen = pdf_names[0]
    log.warning(
        f"    No {lang} marker found in ZIP PDFs {pdf_names}; using '{chosen}'"
    )
    return zf.read(chosen), chosen


def download_sec_zip(url: str, lang: str) -> tuple[bytes, str]:
    """Download SEC ZIP and extract the PDF matching *lang*."""
    log.info(f"    Downloading SEC ZIP …")
    resp = requests.get(url, headers=HEADERS, timeout=120, allow_redirects=True)
    resp.raise_for_status()
    with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
        return _pick_pdf_from_zip(zf, lang)

# ──────────────────────────────────────────────────────────────────────────────
# CSV helpers
# ──────────────────────────────────────────────────────────────────────────────

_LOG_HEADER = ["timestamp", "symbol", "doc_type", "language", "status", "filename", "note"]


def _append_log(symbol: str, lang: str, status: str, filename: str, note: str = ""):
    write_header = not Path(LOG_FILE).exists()
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if write_header:
            w.writerow(_LOG_HEADER)
        w.writerow([
            datetime.now().isoformat(),
            symbol, DOC_TYPE_LABEL, lang, status, filename, note,
        ])


def _update_status(symbol: str, lang: str, new_status: str):
    """Update the relevant column in esg_report_status.csv."""
    col = "56-1 One Report (ไทย)" if lang.upper() == "TH" else "56-1 One Report (English)"
    path = Path(STATUS_FILE)
    if not path.exists():
        log.warning(f"{STATUS_FILE} not found — skipping status update for {symbol}")
        return

    rows, fieldnames, updated = [], None, False
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        for row in reader:
            if row.get("Symbol") == symbol:
                row[col] = new_status
                updated = True
            rows.append(row)

    if updated:
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(rows)
        log.info(f"    Status updated → {new_status}")
    else:
        log.warning(f"    Symbol {symbol} not found in {STATUS_FILE}")

# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────

def main():
    log.info("Connecting to Google Drive …")
    service = _get_drive_service()
    log.info("Connected.\n")

    ok_count = err_count = skip_count = 0

    for symbol, lang, url, source_type in REPORTS:
        filename  = f"{symbol}_{DOC_TYPE_LABEL}_{lang.upper()}.pdf"
        log.info(f"{'─'*60}")
        log.info(f"  {filename}")
        log.info(f"  source : {source_type}")
        log.info(f"  url    : {url}")

        # Ensure per-symbol subfolder exists
        folder_id = get_or_create_folder(service, symbol, ROOT_FOLDER_ID)

        # Skip if already uploaded
        if file_exists_in_drive(service, filename, folder_id):
            log.info("  → already in Drive, skipping")
            _append_log(symbol, lang, "skipped", filename, "already in Drive")
            skip_count += 1
            continue

        try:
            if source_type == "direct":
                pdf_bytes    = download_direct(url)
                original     = url.rsplit("/", 1)[-1]
            elif source_type == "sec_zip":
                pdf_bytes, original = download_sec_zip(url, lang)
            else:
                raise ValueError(f"Unknown source_type: {source_type!r}")

            file_id = upload_to_drive(service, pdf_bytes, filename, folder_id)
            log.info(f"  ✓ uploaded  ({len(pdf_bytes):,} B)  Drive id: {file_id}")
            _append_log(symbol, lang, "success", filename, f"from: {original}")
            _update_status(symbol, lang, "มีไฟล์")
            ok_count += 1

        except Exception as exc:
            log.error(f"  ✗ FAILED: {exc}")
            _append_log(symbol, lang, "error", filename, str(exc))
            err_count += 1

    log.info(f"\n{'═'*60}")
    log.info(f"  Done — success: {ok_count}  errors: {err_count}  skipped: {skip_count}")
    log.info(f"{'═'*60}")

    if err_count:
        sys.exit(1)


if __name__ == "__main__":
    main()
