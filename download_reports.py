"""
ESG Report Downloader v2 â Fixed version
Fixes:
  1. SEC IDISC maintenance detection â skip immediately instead of 60s timeout
  2. SET Annual Report: try /news-and-publications/annual-report URL first,
     use domcontentloaded + JavaScript evaluation for PDF links
  3. Explicit upload error logging (was silent before)
  4. Minimum file size check before upload (reject 0-byte / HTML responses)
  5. Use asyncio.Semaphore properly with gather
"""

import asyncio
import csv
import os
import shutil
import sys
from pathlib import Path

import openpyxl
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import json

# ââ Config âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
XLSX_PATH = "companies.xlsx"
LOG_FILE  = "download_log.csv"
TMP_DIR   = Path("tmp_downloads")
GDRIVE_ROOT_FOLDER_ID = os.environ.get("GDRIVE_FOLDER_ID", "")
BATCH_START = int(os.environ.get("BATCH_START", 0))
BATCH_END   = int(os.environ.get("BATCH_END", 264))
CONCURRENCY = int(os.environ.get("CONCURRENCY", 3))

# SEC maintenance indicators
SEC_MAINTENANCE_HOSTS = ["secdocumentstorage", "maintenance", "under construction"]
SEC_MAINTENANCE_TEXT  = [
    "à¸à¸£à¸±à¸à¸à¸£à¸¸à¸à¸£à¸°à¸à¸", "à¸à¸´à¸à¸à¸£à¸±à¸à¸à¸£à¸¸à¸", "under maintenance",
    "temporarily unavailable", "secdocumentstorage"
]

# ââ Google Drive helpers ââââââââââââââââââââââââââââââââââââââââââââââââââââ

def get_drive_service():
    cred_json = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON", "")
    if not cred_json:
        raise RuntimeError("GDRIVE_SERVICE_ACCOUNT_JSON not set")
    info = json.loads(cred_json)
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def get_or_create_folder(service, name: str, parent_id: str) -> str:
    q = (
        f"name='{name}' and mimeType='application/vnd.google-apps.folder'"
        f" and '{parent_id}' in parents and trashed=false"
    )
    results = service.files().list(q=q, fields="files(id)").execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]
    meta = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    folder = service.files().create(body=meta, fields="id").execute()
    return folder["id"]


def upload_to_drive(service, local_path: str, filename: str, folder_id: str):
    media = MediaFileUpload(local_path, resumable=True)
    meta  = {"name": filename, "parents": [folder_id]}
    service.files().create(body=meta, media_body=media, fields="id").execute()
    print(f"  â Uploaded {filename} to Drive")


# ââ Checkpoint helpers ââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def load_done() -> set:
    done = set()
    if not Path(LOG_FILE).exists():
        return done
    with open(LOG_FILE, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row.get("status") == "success":
                done.add((row["symbol"], row["doc_type"], row["language"]))
    return done


def log_result(symbol: str, doc_type: str, language: str,
               status: str, filename: str = "", note: str = ""):
    file_exists = Path(LOG_FILE).exists()
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["symbol", "doc_type", "language", "status", "filename", "note"]
        )
        if not file_exists:
            writer.writeheader()
        writer.writerow({
            "symbol": symbol, "doc_type": doc_type, "language": language,
            "status": status, "filename": filename, "note": note,
        })


# ââ Excel reader ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def read_companies(xlsx_path: str) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).strip() if c else "" for c in rows[0]]
    companies = []
    for i, row in enumerate(rows[1:], start=1):
        d = dict(zip(header, row))
        symbol = str(d.get("Symbol", d.get("symbol", ""))).strip()
        if symbol and symbol != "None":
            companies.append({
                "order":     i,
                "symbol":    symbol,
                "name_en":   str(d.get("Name (EN)", d.get("name_en", ""))).strip(),
                "name_th":   str(d.get("Name (TH)", d.get("name_th", ""))).strip(),
                "industry":  str(d.get("Industry", d.get("industry", ""))).strip(),
                "rating":    str(d.get("ESG Rating", d.get("rating", ""))).strip(),
            })
    wb.close()
    return companies


# ââ SEC 56-1 One Report downloader âââââââââââââââââââââââââââââââââââââââââ

async def download_sec_one_report(
    page, symbol: str, out_dir: Path, lang: str = "th"
) -> tuple[bool, str, str]:
    lang_path = "th" if lang == "th" else "en"
    url = f"https://market.sec.or.th/public/idisc/{lang_path}/Idiscreport/56-1oneReport"

    try:
        # Use domcontentloaded so we can detect maintenance redirect fast
        await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    except PWTimeout:
        return False, "", "Timeout on goto SEC"
    except Exception as e:
        return False, "", f"goto error: {e}"

    # ââ Maintenance detection ââ
    current_url = page.url
    if any(s in current_url for s in SEC_MAINTENANCE_HOSTS):
        print(f"  â  SEC MAINTENANCE detected for {symbol} ({lang}) â skipping")
        return False, "", "SEC_MAINTENANCE"

    try:
        content = await page.content()
    except Exception:
        content = ""
    if any(s.lower() in content.lower() for s in SEC_MAINTENANCE_TEXT):
        print(f"  â  SEC MAINTENANCE (page text) for {symbol} ({lang}) â skipping")
        return False, "", "SEC_MAINTENANCE"

    # ââ Search for company ââ
    try:
        await page.wait_for_selector('input[placeholder*="à¸à¹à¸à¸«à¸²"], input[type="search"]',
                                     timeout=20000)
    except PWTimeout:
        return False, "", "Search input not found"

    search_box = page.locator(
        'input[placeholder*="à¸à¹à¸à¸«à¸²"], input[type="search"]'
    ).first
    await search_box.fill(symbol)
    await page.wait_for_timeout(2000)

    # Click first result
    try:
        result = page.locator(
            f'tr:has-text("{symbol}"), li:has-text("{symbol}"), '
            f'div[class*="result"]:has-text("{symbol}")'
        ).first
        await result.wait_for(timeout=10000)
        await result.click()
        await page.wait_for_timeout(2000)
    except Exception:
        return False, "", f"Company result not found: {symbol}"

    # Find PDF download link
    try:
        pdf_link = await page.evaluate("""
            () => {
                const links = Array.from(document.querySelectorAll('a'));
                const pdf = links.find(a =>
                    a.href && (a.href.endsWith('.pdf') ||
                               a.href.includes('/download') ||
                               a.href.includes('/pdf'))
                );
                return pdf ? pdf.href : null;
            }
        """)
    except Exception:
        pdf_link = None

    if not pdf_link:
        return False, "", "No PDF link found on SEC page"

    # Download the PDF
    try:
        response = await page.request.get(pdf_link)
        if not response.ok:
            return False, "", f"HTTP {response.status} for PDF"
        content_bytes = await response.body()
        if len(content_bytes) < 1000:
            return False, "", f"File too small ({len(content_bytes)} bytes) â not a PDF"
        filename = f"{symbol}_56-1OneReport_2567_{lang.upper()}.pdf"
        (out_dir / filename).write_bytes(content_bytes)
        return True, filename, ""
    except Exception as e:
        return False, "", f"Download error: {e}"


# ââ SET Annual Report downloader ââââââââââââââââââââââââââââââââââââââââââââ

async def download_set_annual_report(
    page, symbol: str, out_dir: Path, lang: str = "th"
) -> tuple[bool, str, str]:
    lang_code = "th" if lang == "th" else "en"

    # Priority: news-and-publications/annual-report page (dedicated PDF listing)
    # Fallback: company-snapshot/profile page
    candidate_urls = [
        f"https://www.set.or.th/{lang_code}/market/product/stock/quote/{symbol}/news-and-publications/annual-report",
        f"https://www.set.or.th/{lang_code}/market/product/stock/quote/{symbol}/company-snapshot/profile",
        f"https://www.set.or.th/th/market/product/stock/quote/{symbol}/news-and-publications/annual-report",
    ]

    pdf_url  = None
    last_err = "No annual report link found"

    for url in candidate_urls:
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=45000)
            await page.wait_for_timeout(2500)   # let JS render
        except PWTimeout:
            last_err = f"Timeout: {url}"
            continue
        except Exception as e:
            last_err = f"goto error: {e}"
            continue

        # JavaScript: collect all anchor hrefs that look like Annual Report PDFs
        try:
            links = await page.evaluate("""
                () => {
                    return Array.from(document.querySelectorAll('a')).map(a => ({
                        href: a.href || '',
                        text: (a.textContent || '').trim().slice(0, 120)
                    })).filter(l =>
                        l.href.length > 0 && (
                            l.href.toLowerCase().includes('.pdf') ||
                            l.href.toLowerCase().includes('annual') ||
                            l.text.toLowerCase().includes('annual report') ||
                            l.text.includes('à¸£à¸²à¸¢à¸à¸²à¸à¸à¸£à¸°à¸à¸³à¸à¸µ')
                        )
                    );
                }
            """)
        except Exception:
            links = []

        if not links:
            continue

        # Prefer links that look like direct PDF downloads for recent years
        year_hints = ["2024", "2567", "2025", "2568", "annual"]

        def score_link(lnk):
            h = lnk["href"].lower()
            t = lnk["text"].lower()
            score = 0
            if ".pdf" in h:
                score += 10
            for y in year_hints:
                if y in h or y in t:
                    score += 5
            if "annual" in h or "annual" in t:
                score += 3
            return score

        links_sorted = sorted(links, key=score_link, reverse=True)
        pdf_url = links_sorted[0]["href"]
        break

    if not pdf_url:
        return False, "", last_err

    # Download the PDF
    try:
        response = await page.request.get(pdf_url)
        if not response.ok:
            return False, "", f"HTTP {response.status} fetching Annual Report PDF"
        content_bytes = await response.body()
        if len(content_bytes) < 1000:
            return False, "", f"Annual Report too small ({len(content_bytes)}B) â likely not a PDF"
        filename = f"{symbol}_AnnualReport_2567_{lang.upper()}.pdf"
        (out_dir / filename).write_bytes(content_bytes)
        return True, filename, ""
    except Exception as e:
        return False, "", f"Download error: {e}"


# ââ Process one company âââââââââââââââââââââââââââââââââââââââââââââââââââââ

TASKS = [
    ("56-1 One Report", "th", download_sec_one_report),
    ("56-1 One Report", "en", download_sec_one_report),
    ("Annual Report",   "th", download_set_annual_report),
    ("Annual Report",   "en", download_set_annual_report),
]


async def process_company(
    browser, company: dict, done: set, service, root_folder_id: str, tmp_dir: Path
):
    symbol     = company["symbol"]
    symbol_dir = tmp_dir / symbol
    symbol_dir.mkdir(parents=True, exist_ok=True)

    # Create (or find) Drive subfolder for this company
    try:
        company_folder_id = get_or_create_folder(service, symbol, root_folder_id)
    except Exception as e:
        print(f"[{symbol}] â Could not create Drive folder: {e}")
        for doc_type, lang, _ in TASKS:
            log_result(symbol, doc_type, lang, "folder_error", note=str(e)[:100])
        return

    context = await browser.new_context()
    page    = await context.new_page()

    try:
        for doc_type, lang, download_fn in TASKS:
            key = (symbol, doc_type, lang)
            if key in done:
                print(f"[{symbol}] â­ Skip {doc_type} ({lang}) â already done")
                continue

            print(f"[{symbol}] â {doc_type} ({lang})â¦")
            try:
                success, filename, note = await download_fn(page, symbol, symbol_dir, lang)
            except Exception as e:
                success, filename, note = False, "", f"Unexpected: {e}"

            if success:
                local_path = symbol_dir / filename
                # Verify file exists and is not empty
                if not local_path.exists() or local_path.stat().st_size < 1000:
                    msg = f"File missing or too small after download"
                    print(f"[{symbol}] â {doc_type} ({lang}): {msg}")
                    log_result(symbol, doc_type, lang, "failed", note=msg)
                    continue
                try:
                    upload_to_drive(service, str(local_path), filename, company_folder_id)
                    log_result(symbol, doc_type, lang, "success", filename)
                except Exception as e:
                    err_msg = str(e)[:150]
                    print(f"[{symbol}] â Upload FAILED for {filename}: {err_msg}")
                    log_result(symbol, doc_type, lang, "upload_error", filename, err_msg)
            else:
                if note != "SEC_MAINTENANCE":
                    print(f"[{symbol}] â {doc_type} ({lang}): {note}")
                log_result(symbol, doc_type, lang, "failed", note=note)
    finally:
        await context.close()
        if symbol_dir.exists():
            shutil.rmtree(symbol_dir, ignore_errors=True)


# ââ Main ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

async def main():
    companies = read_companies(XLSX_PATH)
    batch     = companies[BATCH_START:BATCH_END]
    done      = load_done()

    print(f"Companies in batch: {len(batch)} ({BATCH_START}â{BATCH_END})")
    print(f"Already done (success): {len(done)}")

    service = get_drive_service()

    TMP_DIR.mkdir(exist_ok=True)
    sem = asyncio.Semaphore(CONCURRENCY)

    async def run_with_sem(browser, company):
        async with sem:
            await process_company(browser, company, done, service, GDRIVE_ROOT_FOLDER_ID, TMP_DIR)

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        try:
            await asyncio.gather(*[run_with_sem(browser, c) for c in batch])
        finally:
            await browser.close()

    # Upload checkpoint log to Drive
    if Path(LOG_FILE).exists():
        try:
            upload_to_drive(service, LOG_FILE, LOG_FILE, GDRIVE_ROOT_FOLDER_ID)
            print("â Checkpoint log uploaded to Drive")
        except Exception as e:
            print(f"â Could not upload log to Drive: {e}")

    print("Done.")


if __name__ == "__main__":
    asyncio.run(main())
