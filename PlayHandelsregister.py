#!/usr/bin/env python3
"""
Playwright-based CLI for handelsregister.de

- Performs "Advanced search"
- Lists result rows (court, name, state, status)
- Optional: downloads "AD" (Current hard copy printout) PDFs for each result
  to the given --outdir using the filename: <Firmenname>_<YYYY-MM-DD>.pdf

Note: Respects site behavior; do not exceed rate limits.
"""

import argparse
import asyncio
import os
import re
import time
import sys
import pandas as pd
from string import ascii_uppercase
from datetime import date
import exel

from playwright.async_api import async_playwright, TimeoutError as PwTimeoutError


# Map to existing CLI semantics
SCHLAGWORT_OPTIONEN = {
    "all": 1,    # contain all keywords
    "min": 2,    # contain at least one keyword
    "exact": 3,  # exact company name
}


# TODO: Debug helper to print element value and outerHTML
async def _debug_dump_element(page, selector: str, label: str, clip: int = 1500):
    """Print value + outerHTML of an element (trimmed), only for debug."""
    try:
        el = page.locator(selector).first
        # Works for <input> and <textarea>
        try:
            current_value = await el.input_value()
        except Exception:
            current_value = "<no input_value()>"
        outer = await el.evaluate("n => n.outerHTML")
        short = outer if len(outer) <= clip else (outer[:clip] + "…[truncated]")
        print(f"[debug] {label} value:", repr(current_value))
        print(f"[debug] {label} outerHTML:", short)
    except Exception as e:
        print(f"[debug] Could not dump {label}: {e}")

async def _debug_dump_results(page, clip: int = 5000):
    """
    Dump outerHTML of the results area after search.
    Tries ergebnisForm first, then its result table, then the grid.
    """
    selectors = [
        "form#ergebnisForm",
        "form[id^='ergebnisForm']",
        "#ergebnisForm\\:selectedSuchErgebnisFormTable_data",
        "[id$='selectedSuchErgebnisFormTable_data']",
        "table[role='grid']",
    ]
    for sel in selectors:
        try:
            await page.wait_for_selector(sel, timeout=6000)
            el = page.locator(sel).first
            html = await el.evaluate("n => n.outerHTML")
            short = html if len(html) <= clip else (html[:clip] + "…[truncated]")
            print(f"[debug] results HTML from {sel}:\n{short}")
            return
        except Exception:
            continue

    # Last resort: dump body
    try:
        body_html = await page.locator("body").evaluate("n => n.outerHTML")
        short = body_html if len(body_html) <= clip else (body_html[:clip] + "…[truncated]")
        print("[debug] results section not found; dumping <body> instead:\n", short)
    except Exception as e:
        print(f"[debug] could not dump body: {e}")



#TODO: Excel helper functions

def col_letter_to_idx(letter: str) -> int:
    """
    Convert an Excel-style column letter into a zero-based index.
    Examples:
        'A'  -> 0
        'B'  -> 1
        ...
        'Z'  -> 25
        'AA' -> 26
    """
    letter = letter.strip().upper()  # Remove whitespace and normalize to uppercase
    idx = 0
    for ch in letter:
        # Shift previous value by 26 and add current letter position
        # 'A' is 1, 'B' is 2, ..., 'Z' is 26
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1  # Convert from 1-based to 0-based index


def read_jobs_from_excel(path: str, sheet: str | None, name_col: str, regno_col: str | None,
                         sap_supplier_col: str | None, sap_customer_col: str | None, start: int | None,
                         end: int | None):
    """
    Read company data from an Excel file and return a list of jobs to process.

    Parameters:
        path               - Path to the Excel file (e.g., "TestBP.xlsx")
        sheet              - Optional sheet name (None = first sheet)
        name_col           - Column letter for company name (Name1)
        regno_col          - Column letter for register number (AF) or None
        sap_supplier_col   - Column letter for supplier SAP number (A) or None
        sap_customer_col   - Column letter for customer SAP number (other col) or None
        limit              - Maximum number of rows to process (None = all)

    Returns:
        List of dicts with keys: "name", "register_no", "sap"
    """
    # Read the Excel file with no header row (so we can use absolute column positions)
    df = pd.read_excel(path, sheet_name=sheet, header=None)  # absolute column positions
    total_rows = len(df)

    # Normalize 1-based bounds → clamp to [1, total_rows]
    s = start if (start and start > 0) else 1
    e = end if (end and end > 0) else total_rows
    if s > total_rows:  # nothing to do
        return []
    if e < s:
        return []

    # Slice rows (convert to 0-based iloc)
    df_slice = df.iloc[s - 1:e]

    def safe_get(row, col_letter: str | None):
        if not col_letter:
            return None
        idx = col_letter_to_idx(col_letter)
        if idx < 0 or idx >= len(row):
            return None
        val = row.iloc[idx]
        return None if (pd.isna(val) or (isinstance(val, str) and val.strip() == "")) else val

    jobs = []
    for _, row in df_slice.iterrows():
        name = safe_get(row, name_col)
        if not name:
            continue  # skip rows without a company name

        register_no = safe_get(row, regno_col)
        sap_supplier = safe_get(row, sap_supplier_col)
        sap_customer = safe_get(row, sap_customer_col) if sap_customer_col else None

        # Prefer supplier SAP, else customer SAP, else None
        sap_raw = sap_supplier if sap_supplier not in (None, "") else sap_customer

        # Normalize SAP (Excel often gives numerics as floats)
        if isinstance(sap_raw, float) and sap_raw.is_integer():
            sap = str(int(sap_raw))
        else:
            sap = str(sap_raw).strip() if sap_raw not in (None, "") else None

        jobs.append({
            "name": str(name).strip(),
            "register_no": str(register_no).strip() if register_no not in (None, "") else None,
            "sap": sap,
        })

    return jobs


#TODO: Perform search

async def open_startpage(page, debug=False):
    # Landing page → ensure we land on welcome.xhtml
    await page.goto("https://www.handelsregister.de/rp_web/welcome.xhtml", wait_until="domcontentloaded")
    if debug:
        print("[debug] opened welcome page:", page.url)

async def perform_search(page, keyword: str, mode: str, register_number: str = None, debug=False):
    """
    Click 'Advanced search', fill the form, submit.
    """
    # The site has a link/button with text like "Advanced search" (English UI).
    # On German UI it may be "Erweiterte Suche". We try both.
    # Prefer a robust selector by partial text.
    # Open Advanced Search (prefer fixed ID; keep fallbacks)
    clicked = False
    try:
        await page.click("#naviForm\\:erweiterteSucheLink", timeout=3000)
        clicked = True
    except Exception:
        # fallbacks by text (en/de)
        advanced_candidates = [
            page.get_by_role("link", name="Advanced search"),
            page.get_by_role("link", name="Erweiterte Suche"),
            page.locator("a:has-text('Advanced search')"),
            page.locator("a:has-text('Erweiterte Suche')"),
        ]
        for loc in advanced_candidates:
            try:
                await loc.first.click(timeout=2000)
                clicked = True
                break
            except Exception:
                pass

    if not clicked:
        raise RuntimeError("Could not open Advanced search. UI may have changed.")

    # Wait for the form to be present (use a field we know)
    await page.wait_for_selector("#form\\:schlagwoerter", timeout=10000)

    # Form fields (JSF IDs usually 'form:schlagwoerter' and 'form:schlagwortOptionen')
    # We'll try robust selectors by id and by name.
    # Keyword input:
    try:
        await page.fill("#form\\:schlagwoerter", keyword)
    except Exception:
        # fallback: look for input with name form:schlagwoerter
        await page.fill("input[name='form:schlagwoerter']", keyword)

    # Print outerHTML
    #if debug:
        #await _debug_dump_element(page, "#form\\:schlagwoerter", "schlagwoerter")

    #If a register number is provided, fill it in:
    if register_number:
        try:
            await page.fill("#form\\:registerNummer", register_number)
        except Exception:
            await page.fill("input[name='form:registerNummer']", register_number)
        #Print outerHTML
        #if debug:
            #await _debug_dump_element(page, "#form\\:registerNummer", "registerNummer")

    # Radio/select for schlagwortOptionen:
    so_value = SCHLAGWORT_OPTIONEN[mode]
    # This field is rendered as radio buttons (values 1,2,3). Try robust selection:
    # Common pattern: input[name='form:schlagwortOptionen'][value='1']
    radio_selector = f"input[name='form:schlagwortOptionen'][value='{so_value}']"
    await page.check(radio_selector)

    # Submit the form by clicking the search button (works even if only registerNummer is filled)
    try:
        await page.click("#form\\:btnSuche", timeout=2000) # ID for the search button
    except Exception:
        # Fallback by button label
        await page.get_by_role("button", name="Suchen").click()

    if debug:
        print("[debug] clicked search; waiting for results…")

    # Wait for results table, check for specific section in HTML body
    try:
        await page.locator(
            "#ergebnissForm\\:selectedSuchErgebnisFormTable_data"
        ).first.wait_for(timeout=5000)
    except PwTimeoutError:
        # If the results table is not found, it may be a different page or no results.
        # We can check for a message or the URL change.
        if "keine Ergebnisse" in await page.content():
            print("[warn] No results found for the search criteria.")
            return
        else:
            raise RuntimeError("Results table not found; check if the search was successful.")
    #time.sleep(1) # Small pause to ensure results are loaded

    if debug:
        #await _debug_dump_results(page)
        print("[debug] results page:", page.url)

async def get_results(page, debug=False):
    """
    Scrape the visible rows of the results table (first page).
    Returns list of dicts with minimal fields, and row locators for clicking AD per row.
    """
    rows = page.locator("table[role='grid'] tr[data-ri]")  # JSF data row index attribute
    count = await rows.count()
    results = []

    for i in range(count):
        row = rows.nth(i)

        # Typical columns: [state, court, name/company, registered office, status, ...]
        # The exact order can vary slightly; we’ll fetch visible cell texts and assign minimally.
        cells = row.locator("td")
        ccount = await cells.count()
        texts = []
        for j in range(ccount):
            # inner_text keeps formatting; text_content is also fine.
            t = (await cells.nth(j).inner_text()).strip()
            texts.append(t)

        # Heuristic mapping (based on your mechanize parser)
        # 0: (ui panel filler)
        # 1: court
        # 2: company/name
        # 3: registered office (state sometimes elsewhere)
        # 4: status
        # We’ll just keep some fields robustly:
        court = texts[1] if len(texts) > 1 else ""
        name = texts[2] if len(texts) > 2 else ""
        registered_office = texts[3] if len(texts) > 3 else ""
        status = texts[4] if len(texts) > 4 else ""

        results.append(
            {
                "row_index": i,
                "court": court,
                "name": name,
                "registered_office": registered_office,
                "status": status,
                "row_locator": row,  # keep for AD click
            }
        )

    return results


#TODO: Download functions

def sanitize_filename(name: str) -> str:
    # Windows-safe filename
    name = name.strip()
    return re.sub(r'[\\/:*?"<>|]+', "_", name)

def create_human_check_file():
    """Create a HumanCheck text file in Downloads directory for searches with multiple results or none"""
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    check_file = os.path.join(downloads_path, "HumanCheck.txt")

    if os.path.exists(check_file):
        return

    try:
        with open(check_file, "w") as f:
            f.write(f"Script execution started at: {date.today().strftime('%d-%m-%Y %H:%M:%S')}")
        print(f"[info] Created human check file: {check_file}")
    except Exception as e:
        print(f"[error] Failed to create human check file: {e}")

def replace_umlauts(text):
    """Replace German umlauts after uppercase conversion"""
    replacements = {
        'Ö': 'OE',
        'Ä': 'AE',
        'Ü': 'UE',
        'ß': 'SS'
    }
    for umlaut, replacement in replacements.items():
        text = text.replace(umlaut, replacement)
    return text

async def download_ad_for_row(page, row_locator, company_name, outdir, sap_number=None, debug=False):
    """
    Waits for the single search result, clicks the AD link, and saves the PDF as
    '<Company>_dd.mm.yyyy.pdf'. Returns the saved path or None on failure.
    """
    #for now only one row is expected,
    #TODO: if more than one row is found, iterate over them, use row_locator
    #Uppercase and replace umlauts in company name
    company_name = replace_umlauts(company_name.upper())
    try:
        #Check if the outdir exists, if not create it
        os.makedirs(outdir, exist_ok=True)


        tbody = page.locator("#ergebnisForm\\:selectedSuchErgebnisFormTable_data")


        # Most specific & stable selector for AD (per your HTML)
        ad_link = page.locator("#ergebnissForm\\:selectedSuchErgebnisFormTable_data a[onclick*='Global.Dokumentart.AD']")
        try:
            await ad_link.first.wait_for(state="attached", timeout=2000)
            await ad_link.first.scroll_into_view_if_needed()
        except PwTimeoutError:
            print(f"[warn] Timeout waiting for visibility: '{company_name}'.")
            return None

        # Trigger and capture the download
        try:
            async with page.expect_download(timeout=5000) as dl_info:
                await ad_link.click()
            download = await dl_info.value
            if debug:
                print(f"[debug] Download started for '{company_name}': {download.suggested_filename}")
        except Exception:
            if debug:
                print(f"[warn] Failed to click AD link for '{company_name}'; download may not have started.")
            return None

        # sap_company_dd.mm.yyyy filename
        date_str = date.today().strftime("%d-%m-%Y")
        prefix = (sanitize_filename(sap_number) + "_") if sap_number else ""
        fname = f"{prefix}{sanitize_filename(company_name)}_{date_str}.pdf"
        save_path = os.path.join(outdir, fname)

        await download.save_as(save_path)
        if debug:
            print(f"[debug] Saved AD PDF: {save_path}")
        return save_path

    except Exception as e:
        if debug:
            print(f"[warn] Failed to download AD for '{company_name}': {e}")
        return None


#TODO: Main

async def main_async(args):
    """
        Main asynchronous entry point for the Handelsregister script.
        Depending on CLI args, runs either:
          - Excel batch mode: read multiple company names/register numbers from Excel
          - Single-shot mode: run for a single provided search term

        Handles:
          - Opening the Handelsregister start page
          - Performing search queries
          - Downloading 'AD' PDF documents if requested
          - Logging cases where results are ambiguous
    """
    # Ensure output dir
    if not os.path.exists(args.outdir):
        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "BP")
        print(f"[warn] Path not found: {args.outdir}")
        print(f"[info] Creating default directory: {default_path}")
        os.makedirs(default_path, exist_ok=True)
        args.outdir = default_path
    else:
        os.makedirs(args.outdir, exist_ok=True)

    async with async_playwright() as p:
        # Headless Chromium; accept downloads in the context
        browser = await p.chromium.launch(headless=not args.headful)
        context = await browser.new_context(accept_downloads=True, locale="en-GB")
        page = await context.new_page()

        await open_startpage(page, debug=args.debug)




        #TODO: Excel batch mode
        if args.excel:
            # Read company/job data from Excel using helper function
            jobs = read_jobs_from_excel(
                path=args.excel,
                sheet=args.sheet,
                name_col=args.name_col,
                regno_col=args.regno_col,
                sap_supplier_col=args.sap_supplier_col,
                sap_customer_col=args.sap_customer_col,
                start=args.start,
                end=args.end,
            )

            print(f"[info] loaded {len(jobs)} jobs from Excel (rows {args.start or 1}..{args.end or 'last'})")

            # Iterate through each job (company) from the Excel list
            for i, job in enumerate(jobs, 1):
                kw = job["name"]  # Company name to search for
                reg = job["register_no"]  # Register number (if available)
                sap = job["sap"]  # SAP number (if available)

                if args.debug:
                    print("")
                    print(f"[debug] ({i}/{len(jobs)}) {sap or 'NoSAP'} | {kw} | reg={reg or 'None'}")

                # Refresh the start page for each job (avoids leftover form state) TODO: check if this is needed
                #await open_startpage(page, debug=args.debug)

                # Perform the advanced search with the name + optional register number
                await perform_search(page, kw, args.schlagwortOptionen, register_number=reg, debug=args.debug)

                # Retrieve the search results (list of rows)
                results = await get_results(page, debug=args.debug)

                # If we don't have exactly one match, log it to HumanCheck.txt and skip
                if len(results) != 1:
                    check_file = os.path.join(os.path.expanduser("~"), "Downloads", "HumanCheck.txt")
                    with open(check_file, "a") as f:
                        f.write(f"\n\n[info] found {len(results)} result row(s) for '{kw}' (SAP={sap or 'None'})")
                    print(f"[warn] {kw}: expected 1 result, got {len(results)} → logged to HumanCheck.txt")
                    continue  # Skip to next company

                # If PDF download is enabled, download the AD (Current hard copy printout)
                if args.download_ad:
                    r = results[0]  # The single matching result TODO: Multiple if needed
                    await download_ad_for_row(
                        page=page,
                        row_locator=r["row_locator"],
                        company_name=r["name"] or kw,
                        outdir=args.outdir,
                        sap_number=sap,  # Prefix SAP number to the filename if available
                        debug=args.debug,
                    )

                # Small pause to avoid sending requests too quickly
                await page.wait_for_timeout(1000)


        #TODO: Single-shot mode
        else:
            if not args.schlagwoerter:
                print("In single-shot mode you must provide --schlagwoerter.")
                return
            await perform_search(page, args.schlagwoerter, args.schlagwortOptionen, register_number=args.register_number, debug=args.debug)
            results = await get_results(page, debug=args.debug)
            if len(results) != 1:
                # Write to HumanCheck.txt in Downloads
                check_file = os.path.join(os.path.expanduser("~"), "Downloads", "HumanCheck.txt")
                with open(check_file, "a") as f:
                    f.write(f"\n\n[info] found {len(results)} result row(s) for '{args.schlagwoerter}'")
                #    for r in results:
                #        f.write(f"\nname: {r['name']}")
                #        f.write(f"\ncourt: {r['court']}")
                #        f.write(f"\nstate/office: {r['registered_office']}")
                #        f.write(f"\nstatus: {r['status']}")
                #        f.write("\n")

                print("[warn] Found multiple results or none; check HumanCheck.txt in Downloads for details.")
                return
            #check if only one result is found, if more TODO: interation already exists

            if args.download_ad and len(results) == 1:
                # Process each row and click AD
                for r in results:
                    await download_ad_for_row(
                        page=page,
                        row_locator=r["row_locator"],
                        company_name=r["name"] or "company",
                        outdir=args.outdir,
                        sap_number=None, # No SAP number in this case TODO
                        debug=args.debug,
                    )
                    # Small pause to be gentle (adjust if needed)
                    #await page.wait_for_timeout(1200)

        await context.close()
        await browser.close()


def parse_args():
    parser = argparse.ArgumentParser(description="A handelsregister CLI (Playwright)")
    parser.add_argument(
        "-d", "--debug",
        help="Enable debug-style prints",
        action="store_true"
    )
    parser.add_argument(
        "-s", "--schlagwoerter",
        help="Search for the provided keywords",
        default=None,
    )
    parser.add_argument(
        "-so", "--schlagwortOptionen",
        help="Keyword options: all=contain all keywords; min=contain at least one; exact=exact company name.",
        choices=["all", "min", "exact"],
        default="all"
    )
    parser.add_argument(
        "--download-ad",
        help="Download the AD (Current hard copy printout) PDF for each result row.",
        action="store_true",
    )
    parser.add_argument(
        "--outdir",
        help="Output directory for downloaded PDFs",
        default=os.path.join(os.path.expanduser("~"), "Downloads", "BP"),
    )
    parser.add_argument(
        "--headful",
        help="Run with a visible browser window (useful for debugging).",
        action="store_true",
    )
    parser.add_argument(
        "-rn", "--register-number",
        help="Optional: Handelsregisternummer to search for",
        required=False
    )
    # for excel import
    parser.add_argument("--excel", help="Path to Excel file (e.g., TestBP.xlsx)")
    parser.add_argument("--sheet", default=None, help="Optional sheet name")
    parser.add_argument("--name-col", default="C", help="Excel column with Name1 (default C)")
    parser.add_argument("--regno-col", default="AF", help="Excel column with register number (default AF)")
    parser.add_argument("--sap-supplier-col", default="A", help="Excel column with supplier SAP no. (default A)")
    parser.add_argument("--sap-customer-col", default="B", help="Excel column with customer SAP no. (optional)")
    parser.add_argument(
        "--start",
        type=int,
        default=None,
        help="Start row (1-based) in the Excel sheet to process (default: first row)."
    )
    parser.add_argument(
        "--end",
        type=int,
        default=None,
        help="End row (1-based, inclusive) in the Excel sheet to process (default: last row)."
    )

    return parser.parse_args()


def main():
    create_human_check_file()
    args = parse_args()
    try:
        asyncio.run(main_async(args))
    except KeyboardInterrupt:
        print("Aborted by user.")
    except Exception as e:
        print(f"[error] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()



#TODO: Error website ->reload start from the element where it failed, exel empty rows to hold position, over 5 words, umlaute
# python PlayHandelsregister.py -d --download-ad --excel "C:\Users\Nguyen-Bang\Downloads\TestBP.xlsx" --sheet "Tabelle1" --start 25 --end 30
# python PlayHandelsregister.py -s "THYSSENKRUPP SCHULTE GMBH"  --register-number "26718" -d --download-ad
