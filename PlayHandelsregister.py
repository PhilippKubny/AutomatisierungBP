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
import PDFScanner

from playwright.async_api import async_playwright, TimeoutError as PwTimeoutError


page = None  # Global page object for async context
counter = 0 #for timer logic
reruns = 0 #for reruns logic

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

#TODO: No reach, errors

async def rerun_search():
    global page
    global reruns
    reruns += 1
    if reruns > 3:
        # sleep for 10 minutes if more than 3 reruns
        print("[warn] More than 3 reruns, sleeping for 10 minutes to avoid rate limiting...")
        time.sleep(600)
        reruns = 0
    page = await page.context.new_page()  # Reset page context
    await open_startpage(debug=False)


#TODO: Perform search

async def open_startpage(debug=False):
    # Landing page → ensure we land on welcome.xhtml
    global page
    await page.goto("https://www.handelsregister.de/rp_web/welcome.xhtml", wait_until="domcontentloaded")
    if debug:
        print("[debug] opened welcome page:", page.url)

async def perform_search(keyword: str, mode: str, register_number: str = None, postal_code: str = None, postal_code_option=False, debug=False):
    """
    Click 'Advanced search', fill the form, submit.
    """
    # The site has a link/button with text like "Advanced search" (English UI).
    # On German UI it may be "Erweiterte Suche". We try both.
    # Prefer a robust selector by partial text.
    # Open Advanced Search (prefer fixed ID; keep fallbacks)
    global page
    global counter
    clicked = False
    try:
        await page.click("#naviForm\\:erweiterteSucheLink", timeout=10000)
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
        print("[warn] Could not open Advanced search. UI may have changed.")
        print("[warn] Rerun")
        page = await page.context.new_page()  # Reset page context
        await open_startpage(debug=debug)
        await perform_search(keyword, mode, register_number=register_number, postal_code=postal_code, postal_code_option=postal_code_option, debug=debug)
        return

    # Wait for the form to be present (use a field we know)
    try:
        await page.wait_for_selector("#form\\:schlagwoerter", timeout=10000)
    except PwTimeoutError:
        print("[warn] Advanced search form not found; UI may have changed.")
        print("[warn] Rerun")
        page = await page.context.new_page()  # Reset page context
        await open_startpage(debug=debug)
        await perform_search(keyword, mode, register_number=register_number, postal_code=postal_code, postal_code_option=postal_code_option, debug=debug)
        return

    # Form fields (JSF IDs usually 'form:schlagwoerter' and 'form:schlagwortOptionen')
    # We'll try robust selectors by id and by name.
    # Keyword input:
    try:
        words = [w for w in re.split(r"[ -]+", keyword) if w]
        first_five = " ".join(words[:5])
        keyword = first_five  # Limit to first 5 words for search
        await page.fill("#form\\:schlagwoerter", keyword)
    except Exception:
        print("[warn] Could not fill 'schlagwoerter' by ID")
        print("[warn] Rerun")
        page = await page.context.new_page()  # Reset page context
        await open_startpage(debug=debug)
        await perform_search(keyword, mode, register_number=register_number, postal_code=postal_code, postal_code_option=postal_code_option, debug=debug)
        return

    # Print outerHTML
    #if debug:
        #await _debug_dump_element(page, "#form\\:schlagwoerter", "schlagwoerter")

    #If a register number is provided, fill it in:

    try:
        if register_number is None:
            register_number = ""
        await page.fill("#form\\:registerNummer", register_number)
    except Exception:
        print("[warn] Could not fill 'registerNummer' by ID")
        print("[warn] Rerun")
        page = await page.context.new_page()  # Reset page context
        await open_startpage(debug=debug)
        await perform_search(keyword, mode, register_number=register_number, postal_code=postal_code, postal_code_option=postal_code_option, debug=debug)
        return

        #Print outerHTML
        #if debug:
            #await _debug_dump_element(page, "#form\\:registerNummer", "registerNummer")

    if postal_code_option:
        # Fill postal code if provided
        try:
            if postal_code is None:
                postal_code = ""
            await page.fill("#form\\:postleitzahl", postal_code)
            if debug:
                print("[debug] postal code:", postal_code)
        except Exception:
            print("[warn] Could not fill 'postleitzahl' by ID")
            print("[warn] Rerun")
            page = await page.context.new_page()  # Reset page context
            await open_startpage(debug=debug)
            await perform_search(keyword, mode, register_number=register_number, postal_code=postal_code, postal_code_option=postal_code_option, debug=debug)
            return


    # Radio/select for schlagwortOptionen:
    so_value = SCHLAGWORT_OPTIONEN[mode]
    # This field is rendered as radio buttons (values 1,2,3). Try robust selection:
    # Common pattern: input[name='form:schlagwortOptionen'][value='1']
    radio_selector = f"input[name='form:schlagwortOptionen'][value='{so_value}']"
    await page.check(radio_selector)

    # Submit the form by clicking the search button (works even if only registerNummer is filled)
    try:
        await page.click("#form\\:btnSuche", timeout=15000) # ID for the search button
    except Exception:
        # Fallback by button label
        print("[warn] Could not find search button by ID; trying by role.")
        print("[warn] Rerun")
        page = await page.context.new_page()  # Reset page context
        await open_startpage(debug=debug)
        await perform_search(keyword, mode, register_number=register_number, postal_code=postal_code, postal_code_option=postal_code_option, debug=debug)
        return


    if debug:
        print("[debug] clicked search; waiting for results…")

    counter += 1 # Successfully clicked on search -> counter +1
    # Wait for results table, check for specific section in HTML body
    try:
        await page.locator(
            "#ergebnissForm\\:selectedSuchErgebnisFormTable_data"
        ).first.wait_for(timeout=20000)
    except PwTimeoutError:
        print("[warn] Results table not found; UI may have changed.")
        print("[warn] Rerun")
        page = await page.context.new_page()  # Reset page context
        await open_startpage(debug=debug)
        await perform_search(keyword, mode, register_number=register_number, postal_code=postal_code, postal_code_option=postal_code_option, debug=debug)
        return
    #time.sleep(1) # Small pause to ensure results are loaded

    if debug:
        #await _debug_dump_results(page)
        print("[debug] results page:", page.url)

async def get_results(debug=False):
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

async def download_ad_for_row(row_locator, company_name, outdir, sap_number=None, debug=False):
    """
    Waits for the single search result, clicks the AD link, and saves the PDF as
    '<Company>_dd.mm.yyyy.pdf'. Returns the saved path or None on failure.
    """
    #for now only one row is expected,
    #TODO: if more than one row is found, iterate over them, use row_locator
    #Uppercase and replace umlauts in company name
    global page
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
            async with page.expect_download(timeout=25000) as dl_info:
                await ad_link.click()
            download = await dl_info.value
            if debug:
                print(f"[debug] Download started for '{company_name}': {download.suggested_filename}")
        except Exception:

            check_file = os.path.join(os.path.expanduser("~"), "Downloads", "HumanCheck.txt")
            with open(check_file, "a") as f:
                f.write(f"\n\n[warn] Failed to click AD link for '{company_name}'; download may not have started. (SAP={sap_number or 'None'})")
            if debug:
                print(f"[warn] Failed to click AD link for '{company_name}'; download may not have started.")
            return None

        # sap_company_dd.mm.yyyy filename
        date_str = date.today().strftime("%d-%m-%Y")
        prefix = (sanitize_filename(sap_number) + "_") if sap_number else ""
        fname = f"{prefix}{sanitize_filename(company_name)}_{date_str}.pdf"
        save_path = os.path.join(outdir, fname)

        await download.save_as(save_path)
        # Verify the file exists and has size > 0 (not corrupted)
        if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
            if debug:
                print(f"[debug] Saved AD PDF: {save_path}")
            return save_path
        else:
            if debug:
                print(f"[warn] PDF download failed or file empty: {save_path}")
            return None

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
    global page
    global counter
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

        await open_startpage(debug=args.debug)




        #TODO: Excel batch mode
        if args.excel:
            # Read company/job data from Excel using helper function
            jobs = exel.read_jobs_from_excel(
                path=args.excel,
                sheet=args.sheet,
                name_col=args.name_col,
                regno_col=args.regno_col,
                sap_supplier_col=args.sap_supplier_col,
                sap_customer_col=args.sap_customer_col,
                postal_code_col=args.postal_code_check_col,
                start=args.start,
                end=args.end,
            )

            print(f"[info] loaded {len(jobs)} jobs from Excel (rows {args.start or 3}..{args.end or 'last'})")

            start_time = time.time() # 1 hour timer to not go over the 60 search limitation (VERY IMPORTANT)
            # Iterate through each job (company) from the Excel list
            for i, job in enumerate(jobs, 1):
                # Timer logic: every 60 jobs, wait for the remaining time to complete the hour
                if counter >= 60:
                    elapsed_time = time.time() - start_time
                    hour_in_seconds = 3600

                    # If less than an hour has passed, wait for remaining time
                    if elapsed_time < hour_in_seconds:
                        wait_time = hour_in_seconds - elapsed_time
                        print(f"[info] Waiting {wait_time:.0f} seconds to complete the hour...")
                        await page.wait_for_timeout(wait_time * 1000)  # convert to milliseconds

                    # Reset the timer
                    start_time = time.time()
                    print("[info] Hour timer reset")
                    page = await context.new_page()  # Reset page context
                    await open_startpage(debug=args.debug)


                if job["name"] is None:
                    print(f"[warn] Skipping job {i}: no company name provided.")

                    continue
                kw = job["name"]  # Company name to search for
                reg = str(job["register_no"]) if job["register_no"] is not None else ""  # Register number (if available), normalized if it was in float
                sap = job["sap"]  # SAP number (if available)
                postal_code = job["postal_code"]  # Postal code (if available)

                if args.debug:
                    print("")
                    print(f"[debug] ({i}/{len(jobs)}) {sap or 'NoSAP'} | {kw} | reg={reg or 'None'}")

                # Refresh the start page for each job (avoids leftover form state) TODO: check if this is needed
                #await open_startpage(page, debug=args.debug)

                # Perform the advanced search with the name + optional register number
                await perform_search(kw, args.schlagwortOptionen, register_number=reg, postal_code=postal_code, postal_code_option=args.postal, debug=args.debug)
                print(f"[debug] {kw} | reg={reg or 'None'}")

                # Retrieve the search results (list of rows)
                results = await get_results(debug=args.debug)

                # If we don't have exactly one match, log it to HumanCheck.txt and skip
                check_file = os.path.join(os.path.expanduser("~"), "Downloads", "HumanCheck.txt")
                f = open(check_file, "a")
                if len(results) != 1:
                    f.write(f"\n[info] found {len(results)} result row(s) for '{kw}' (SAP={sap or 'None'})")
                    print(f"[warn] {kw}: expected 1 result, got {len(results)} → logged to HumanCheck.txt")
                    exel.write_to_excel_error(
                        path=args.excel,
                        sheet=args.sheet,
                        row=i + args.start - 1,  # Adjust for 0-based index
                        changes_check_col=args.changes_check_col,
                    )
                    print(
                        f"[warn] Failed '{kw}' (SAP={sap or 'None'}); marked row {i + args.start - 1} in Excel as error.")
                    continue  # Skip to next company

                # If PDF download is enabled, download the AD (Current hard copy printout)
                path = None
                r = results[0]  # The single matching result TODO: Multiple if needed
                company_name = replace_umlauts(r["name"].upper()) # Uppercase and replace umlauts
                if args.download_ad:
                    path = await download_ad_for_row(
                        row_locator=r["row_locator"],
                        company_name=company_name,
                        outdir=args.outdir,
                        sap_number=sap,  # Prefix SAP number to the filename if available
                        debug=args.debug,
                    )

                if path is not None:
                    update_info = PDFScanner.extract_from_pdf(path) | {"company_name": company_name, "sap_number": sap, "download_path": path} # Extract fields from the downloaded PDF into a dict, override company_name with umlauts replaced
                    if update_info["register_type"] == "unexpected Format":
                        f.write(f"[warn] Error, unexpected PDF Format '{company_name}' (SAP={sap or 'None'}) in row {i + args.start - 1}")
                        print(f"[warn] Error, unexpected PDF Format '{company_name}' (SAP={sap or 'None'})")
                        exel.write_to_excel_error(
                            path=args.excel,
                            sheet=args.sheet,
                            row=i + args.start - 1,  # Adjust for 0-based index
                            changes_check_col=args.changes_check_col,
                            pdf_path = path,  # Save the path of PDF in excel
                            pdf_path_col=args.doc_path_col,
                        )
                        continue
                    exel.write_update_to_excel(
                        path=args.excel,
                        sheet=args.sheet,
                        row=i + args.start - 1,  # Adjust for 0-based index
                        update_info=update_info, # All info to update
                        name_col=args.name1_col,
                        regno_col=args.regno_col,
                        sap_supplier_col=args.sap_supplier_col,
                        sap_customer_col=args.sap_customer_col,
                        name2_col=args.name2_col,
                        name3_col=args.name3_col,
                        street_col=args.street_col,
                        house_number_col=args.house_number_col,
                        city_col=args.city_col,
                        postal_code_col=args.postal_code_col,
                        doc_path_col=args.doc_path_col,
                        changes_check_col=args.changes_check_col,
                        date_check_col=args.date_check_col,
                        register_type_col=args.register_type_col,
                        check_file=os.path.join(os.path.expanduser("~"), "Downloads", "HumanCheck.txt")
                    )
                    print(f"[info] Updated Excel row {i + args.start - 1} for '{company_name}' (SAP={sap or 'None'})")


                
                # Small pause to avoid sending requests too quickly
                await page.wait_for_timeout(1000)

        
        #TODO: Single-shot mode
        else:
            if args.sap_number == "-1":
                print("In single-shot mode you must provide --sap_number.")
                return
            if not args.schlagwoerter:
                print("In single-shot mode you must provide --schlagwoerter.")
                return
            if args.row_number == "-1":
                print("In single-shot mode you must provide --row_number.")
                return
            await perform_search(args.schlagwoerter, args.schlagwortOptionen, register_number=args.register_number, debug=args.debug)
            results = await get_results(debug=args.debug)
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

            path = None  # Initialize path to None
            if args.download_ad and len(results) == 1:
                # Process each row and click AD
                for r in results:
                    path = await download_ad_for_row(
                        row_locator=r["row_locator"],
                        company_name=r["name"] or "company",
                        outdir=args.outdir,
                        sap_number=args.sap_number, # No SAP number in this case TODO
                        debug=args.debug,
                    )
                    # Small pause to be gentle (adjust if needed)
                    #await page.wait_for_timeout(1200)
            
            if path is not None:
                    update_info = PDFScanner.extract_from_pdf(path) | {"company_name": results[0]["name"], "sap_number": args.sap_number, "download_path": path} # Extract fields from the downloaded PDF into a dict, override company_name with umlauts replaced
                    if update_info["register_type"] == "unexpected Format":
                        print(f"[warn] Error, unexpected PDF Format '{results[0]['name']}' (SAP={args.sap_number or 'None'})")
                        exel.write_to_excel_error(
                            path=args.excel,
                            sheet=args.sheet,
                            row=args.row_number,
                            changes_check_col=args.changes_check_col,
                            pdf_path = path,  # Save the path of PDF in excel
                            pdf_path_col=args.doc_path_col,
                        )
                        return
                    exel.write_update_to_excel(
                        path=args.excel,
                        sheet=args.sheet,
                        row=args.row_number,
                        update_info=update_info, # All info to update
                        name_col=args.name1_col, # Update column
                        regno_col=args.regno_col,
                        sap_supplier_col=args.sap_supplier_col,
                        sap_customer_col=args.sap_customer_col,
                        name2_col=args.name2_col,
                        name3_col=args.name3_col,
                        street_col=args.street_col,
                        house_number_col=args.house_number_col,
                        city_col=args.city_col,
                        postal_code_col=args.postal_code_col,
                        doc_path_col=args.doc_path_col,
                        changes_check_col=args.changes_check_col,
                        date_check_col=args.date_check_col,
                        register_type_col=args.register_type_col,
                        check_file=os.path.join(os.path.expanduser("~"), "Downloads", "HumanCheck.txt")
                    )
                    print(f"[info] Updated Excel row {args.row} for '{results[0]["name"]}' (SAP={args.sap_number or 'None'})")
        
                    

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
        "-postal", "--postal",
        help="Perform search with postal code",
        action="store_true"
    )
    parser.add_argument(
        "--postal-code-check-col",
        default="I",
        help="Excel column with Postal Code (default I)"
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
    parser.add_argument(
        "-sap", "--sap-number",
        default="-1",
        help="Optional: for singelshot only",
        required=False
    )
    parser.add_argument(
        "-row", "--row-number",
        default="-1",
        help="Optional: for singelshot only",
        required=False
    )
    # for excel import
    parser.add_argument("--excel", help="Path to Excel file (e.g., TestBP.xlsx)")
    parser.add_argument("--sheet", default=None, help="Optional sheet name")
    parser.add_argument("--name-col", default="C", help="Excel column with Name1 (default C)")
    parser.add_argument("--regno-col", default="AF", help="Excel column with register number (default AF)")
    parser.add_argument("--sap-supplier-col", default="A", help="Excel column with supplier SAP no. (default A)")
    parser.add_argument("--sap-customer-col", default="B", help="Excel column with customer SAP no. (default B)")

    # excel columns need to be changed, only for change not for extraction
    parser.add_argument("--name1-col", default="T", help="Excel column with Name1 (default T)")
    parser.add_argument("--name2-col", default="D", help="Excel column with Name2 (default D)")
    parser.add_argument("--name3-col", default="E", help="Excel column with Name3 (default E)")
    parser.add_argument("--street-col", default="X", help="Excel column with Street (default X)")
    parser.add_argument("--house-number-col", default="Y", help="Excel column with house number (default Y)")
    parser.add_argument("--city-col", default="Z", help="Excel column with City (default Z)")
    parser.add_argument("--postal-code-col", default="AA", help="Excel column with Postal Code (default AA)") #TODO: check if this is correct
    parser.add_argument("--doc-path-col", default="P", help="Excel column with document stored (default P)")
    parser.add_argument("--changes-check-col", default="Q", help="Excel column with Changes necessary (default Q)")
    parser.add_argument("--date-check-col", default="S", help="Excel column with Date of last check (default S)")
    parser.add_argument("--register-type-col", default="AG", help="Excel column with Register type (default AG)")
    #parser.add_argument("--country-col", default="AB", help="Excel column with Country (default AB)")


    parser.add_argument(
        "--start",
        type=int,
        default=3,
        help="Start row (comapnies start at row 3)."
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
