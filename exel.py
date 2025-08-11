import argparse
import asyncio
import os
import re
import time
import sys
import pandas as pd
from string import ascii_uppercase
from datetime import date

from playwright.async_api import async_playwright, TimeoutError as PwTimeoutError

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


def main():
    parser = argparse.ArgumentParser(description="Read company data from an Excel file.")

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

    args = parser.parse_args()

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

    print(f"Found {len(jobs)} jobs:")
    for job in jobs:
        print(job)




if __name__ == "__main__":
    main()
