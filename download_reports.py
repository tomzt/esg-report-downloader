"""
download_reports.py
-------------------
Download 56-1 One Report (SET documents page) and Annual Report (SET) for 264 ESG companies.
Runs on GitHub Actions → uploads PDFs to Google Drive.

Requirements: playwright, openpyxl, google-api-python-client, google-auth
"""

import asyncio
import csv
import json
import os
import sys
import time
from pathlib import Path

import openpyxl
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ── Google Drive upload ──────────────────────────────────────────────────────
def get_drive_service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    creds_json = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        raise RuntimeError("GDRIVE_SERVICE_ACCOUNT_JSON secret not set")

    info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


def get_or_create_folder(service, name, parent_id):
    """Return folder ID, creating it if needed."""
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


def upload_to_drive(service, local_path, filename, folder_id):
    from googleapiclient.http import MediaFileUpload

    # resumable=False avoids GCS backend permission issues with service accounts
    media = MediaFileUpload(local_path, resumable=False)
    meta = {"name": filename, "parents": [folder_id]}
    service.files().create(body=meta, media_body=media, fields="id").execute()
    print(f"  ✓ Uploaded {filename} to Drive")


# ── Checkpoint helpers ────────────────────────────────────────────────────────
LOG_FILE = "download_log.csv"
LOG_FIELDS = ["symbol", "doc_type", "language", "status", "filename", "note"]


def load_done() -> set:
    """Return set of (symbol, doc_type, language) already done."""
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
        w.writerow(
            {
                "symbol": symbol,
                "doc_type": doc_type,
                "language": language,
                "status": status,
                "filename": filename,
                "note": note,
            }
        )


# ── Read Excel ────────────────────────────────────────────────────────────────
def read_companies(xlsx_path: str) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    companies = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[1]:
            continue
        companies.append(
            {
                "order": row[0],
                "symbol": str(row[1]).strip(),
                "name_en": str(row[2]).strip() if row[2] else "",
                "name_th": str(row[3]).strip() if row[3] else "",
                "industry": str(row[4]).strip() if row[4] else "",
                "rating": str(row[5]).strip() if row[5] else "",
            }
        )
    return companies


# ── SET: Download 56-1 One Report from SET documents page ────────────────────
async def download_sec_one_report(page, symbol: str, out_dir: Path, lang: str = "th") -> tuple[bool, str, str]:
    """
    Download 56-1 One Report from SET /documents page.
    More reliable than SEC IDISC (frequent maintenance).
    lang: 'th' or 'en'
    Returns (success, filename, note)
    """
    lang_code = "th" if lang == "th" else "en"
    url = f"https://www.set.or.th/{lang_code}/market/product/stock/quote/{symbol}/company-profile/information"

    captured_urls = []
    page.on("response", lambda r: captured_urls.append(r.url) if "json" in r.headers.get("content-type","") else None)
    try:
        await page.goto(url, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(6000)

        pdf_link = None

        # Strategy 1: href or text contains 56-1 / one-report keywords
        selectors = [
            'a[href*="56-1"]',
            'a[href*="56_1"]',
            'a[href*="one-report" i]',
            'a[href*="onereport" i]',
            'a[href*="OneReport"]',
        ]
        for sel in selectors:
            links = page.locator(sel)
            c = await links.count()
            if c > 0:
                for i in range(c):
                    href = await links.nth(i).get_attribute("href") or ""
                    text = await links.nth(i).inner_text()
                    if "2024" in href or "2567" in href or "2024" in text or "2567" in text:
                        pdf_link = href
                        break
                if not pdf_link:
                    pdf_link = await links.first.get_attribute("href")
                if pdf_link:
                    break

        # Strategy 2: any PDF link near "56" text
        if not pdf_link:
            all_pdf = page.locator('a[href$=".pdf"], a[href*="/pdf/"]')
            c = await all_pdf.count()
            for i in range(c):
                href = await all_pdf.nth(i).get_attribute("href") or ""
                text = await all_pdf.nth(i).inner_text()
                if "56" in href or "56" in text or "one" in href.lower():
                    pdf_link = href
                    break

        if not pdf_link:
            total_links = await page.locator("a").count()
            pdf_links = await page.locator("a[href$=\".pdf\"]").count()
            body_text = await page.inner_text("body")
            has56 = "56-1" in body_text or "56" in body_text
            url_now = page.url
            return False, "", f"No 56-1 link (apis={captured_urls})"

        if not pdf_link.startswith("http"):
            pdf_link = "https://www.set.or.th" + pdf_link

        filename = f"{symbol}_56-1_OneReport_2567_{lang.upper()}.pdf"
        out_path = out_dir / filename

        response = await page.request.get(pdf_link)
        if response.ok:
            content = await response.body()
            if len(content) < 1000:
                return False, "", "Downloaded file too small — likely not a real PDF"
            out_path.write_bytes(content)
            return True, filename, ""
        else:
            return False, "", f"HTTP {response.status}"

    except PWTimeout:
        return False, "", "Timeout"
    except Exception as e:
        return False, "", str(e)[:100]


# ── SET: Download Annual Report ───────────────────────────────────────────────
async def download_set_annual_report(page, symbol: str, out_dir: Path, lang: str = "th") -> tuple[bool, str, str]:
    """
    Navigate SET company profile and download Annual Report 2024/2567.
    lang: 'th' or 'en'
    Returns (success, filename, note)
    """
    lang_code = "th" if lang == "th" else "en"
    url = f"https://www.set.or.th/{lang_code}/market/product/stock/quote/{symbol}/company-profile/information"

    captured_urls = []
    page.on("response", lambda r: captured_urls.append(r.url) if "json" in r.headers.get("content-type","") else None)
    try:
        await page.goto(url, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(6000)

        pdf_link = None

        selectors = [
            'a[href*="annual-report" i]',
            'a[href*="annualreport" i]',
            'a:has-text("Annual Report")',
            'a:has-text("รายงานประจำปี")',
            'a[href*=".pdf"][href*="annual" i]',
        ]

        for sel in selectors:
            links = page.locator(sel)
            c = await links.count()
            if c > 0:
                for i in range(c):
                    href = await links.nth(i).get_attribute("href")
                    text = await links.nth(i).inner_text()
                    if href and ("2024" in href or "2567" in href or "2024" in text or "2567" in text):
                        pdf_link = href
                        break
                if not pdf_link and c > 0:
                    pdf_link = await links.first.get_attribute("href")
                if pdf_link:
                    break

        if not pdf_link:
            total_links = await page.locator("a").count()
            pdf_links = await page.locator("a[href$=\".pdf\"]").count()
            body_text = await page.inner_text("body")
            hasAnn = "Annual Report" in body_text or "" in body_text
            url_now = page.url
            return False, "", f"No annual link (apis={captured_urls})"

        if not pdf_link.startswith("http"):
            pdf_link = "https://www.set.or.th" + pdf_link

        filename = f"{symbol}_AnnualReport_2567_{lang.upper()}.pdf"
        out_path = out_dir / filename

        response = await page.request.get(pdf_link)
        if response.ok:
            content = await response.body()
            if len(content) < 1000:
                return False, "", "Downloaded file too small"
            out_path.write_bytes(content)
            return True, filename, ""
        else:
            return False, "", f"HTTP {response.status}"

    except PWTimeout:
        return False, "", "Timeout"
    except Exception as e:
        return False, "", str(e)[:100]


# ── Main ──────────────────────────────────────────────────────────────────────
async def process_company(browser, company: dict, done: set, service, root_folder_id: str, tmp_dir: Path):
    symbol = company["symbol"]
    symbol_dir = tmp_dir / symbol
    symbol_dir.mkdir(exist_ok=True)

    tasks = [
        ("56-1 One Report", "th", download_sec_one_report),
        ("56-1 One Report", "en", download_sec_one_report),
        ("Annual Report",   "th", download_set_annual_report),
        ("Annual Report",   "en", download_set_annual_report),
    ]

    company_folder_id = get_or_create_folder(service, symbol, root_folder_id)

    context = await browser.new_context(accept_downloads=True)
    page = await context.new_page()

    for doc_type, lang, fn in tasks:
        key = (symbol, doc_type, lang)
        if key in done:
            print(f"  [{symbol}] {doc_type} ({lang}) - skipped (already done)")
            continue

        print(f"  [{symbol}] Downloading {doc_type} ({lang})...")
        success, filename, note = await fn(page, symbol, symbol_dir, lang)

        if success:
            try:
                upload_to_drive(service, str(symbol_dir / filename), filename, company_folder_id)
                log_result(symbol, doc_type, lang, "success", filename)
            except Exception as e:
                log_result(symbol, doc_type, lang, "upload_error", filename, str(e)[:100])
        else:
            print(f"    ✗ Failed: {note}")
            log_result(symbol, doc_type, lang, "failed", note=note)

        await asyncio.sleep(2)

    await context.close()

    import shutil
    shutil.rmtree(symbol_dir, ignore_errors=True)


async def main():
    xlsx_path = os.environ.get("EXCEL_PATH", "companies.xlsx")
    root_folder_id = os.environ.get("GDRIVE_FOLDER_ID", "")
    batch_start = int(os.environ.get("BATCH_START", "0"))
    batch_end = int(os.environ.get("BATCH_END", "264"))
    concurrency = int(os.environ.get("CONCURRENCY", "3"))

    if not root_folder_id:
        raise RuntimeError("GDRIVE_FOLDER_ID env var not set")

    companies = read_companies(xlsx_path)
    companies = companies[batch_start:batch_end]
    done = load_done()

    service = get_drive_service()
    tmp_dir = Path("tmp_downloads")
    tmp_dir.mkdir(exist_ok=True)

    print(f"Processing {len(companies)} companies (batch {batch_start}-{batch_end})")
    print(f"Already done: {len(done)} tasks")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )

        sem = asyncio.Semaphore(concurrency)

        async def process_with_sem(company):
            async with sem:
                await process_company(browser, company, done, service, root_folder_id, tmp_dir)

        await asyncio.gather(*[process_with_sem(c) for c in companies])
        await browser.close()

    if Path(LOG_FILE).exists():
        try:
            upload_to_drive(service, LOG_FILE, LOG_FILE, root_folder_id)
        except Exception:
            pass

    print("\nDone!")


if __name__ == "__main__":
    asyncio.run(main())
