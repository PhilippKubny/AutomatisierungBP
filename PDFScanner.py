# extract_register_info.py
# pip install pdfplumber pandas

from pathlib import Path
import re
import unicodedata
import pdfplumber
import pandas as pd

# --- helpers ---------------------------------------------------------------
def replace_umlauts(text):
    """Replace German umlauts after uppercase conversion"""
    replacements = {
        'STR.': 'STRASSE',
        'Ö': 'OE',
        'Ä': 'AE',
        'Ü': 'UE',
        'ß': 'SS'
    }
    for umlaut, replacement in replacements.items():
        text = text.replace(umlaut, replacement)
    return text

def split_german_address(addr: str) -> dict:
    """
    Zerlegt 'STRAßENNAME 12A, 12345 STADT' in {street, house_number, postal_code, city}.
    (Adressformat wie im Handelsregisterauszug unter 'Geschäftsanschrift')
    """

    # 1) Links/Rechts um das erste Komma trennen
    if "," in addr:
        left, right = addr.split(",", 1)
    else:
        # Fallback: kein Komma → wir versuchen trotzdem zu trennen
        left, right = addr, ""

    left = left.strip()
    right = right.strip()

    # 2) Straße + Hausnummer: erste Ziffer in 'left' suchen
    m = re.search(r"\d", left)
    if m:
        street = left[:m.start()].strip()
        house = left[m.start():].strip()
    else:
        # keine Ziffer gefunden → alles als Straße, Rest leer
        street, house = left, ""

    # Hausnummer etwas normalisieren (z.B. "12 A" -> "12A"), aber nur wenn es passt
    if re.fullmatch(r"\d+\s+[A-Za-z]", house):
        house = house.replace(" ", "")

    # 3) Rechts: PLZ + Stadt
    # Primär: 5-stellige deutsche PLZ; Fallback erlaubt 4–5
    m = re.match(r"\s*(\d{5})\s+(.+)", right)
    if not m:
        m = re.match(r"\s*(\d{4,5})\s+(.+)", right)

    if m:
        plz = m.group(1)
        city = m.group(2).strip()
    else:
        # Fallback: keine klare PLZ gefunden
        plz = ""
        city = right

    return {
        "street": street,
        "house_number": house,
        "postal_code": plz,
        "city": city,
    }

def normalize_text(s: str) -> str:
    """Normalize unicode & whitespace; fix common hyphenation at line breaks."""
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u00A0", " ")   # non-breaking space -> space
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"-\n(?=\w)", "", s)  # remove hyphenations at line ends
    s = re.sub(r"[ \t]+", " ", s)
    return s

FIRMA_LINE = re.compile(
    r"\b2\.\s*a\)\s*Firma:\s*(.+)",  # capture only the first line after "Firma:"
    re.IGNORECASE
)
GESCH_ADDR = re.compile(
    r"Gesch[aä]ftsanschrift:\s*(.+)",  # handle ä or ae
    re.IGNORECASE
)

def extract_from_text(text: str) -> dict:
    """
    Extract register type/number, company name (2.a) and address (2.b -> Geschäftsanschrift).
    Returns empty strings if not found.
    """
    text = normalize_text(text)

    # 1) Register type + number (2nd line last 2 words)
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    line2 = lines[1]
    words = line2.split()
    w_type_raw, w_num_raw = words[-2], words[-1]
    reg_type = re.sub(r"[^A-Za-zÄÖÜäöü]", "", w_type_raw).upper()
    reg_number  = re.sub(r"[^\d]", "", w_num_raw)


    # 2) Company name (section 2.a) Firma:)
    company = ""
    m = FIRMA_LINE.search(text)
    if m:
        # take this line; strip any trailing artifacts
        company = m.group(1).strip()
    else:
        # likely the second format
        return extract_from_text_2nd_format(text)

    # 3) Address (after "Geschäftsanschrift:")
    address = ""
    m = GESCH_ADDR.search(text)
    if m:
        # take rest of the line after the label
        address = m.group(1).strip()
        # Sometimes address line ends with a page artifact; trim trailing section markers
        address = re.sub(r"\s*(?:\n|$)", "", address)
    else:
        return extract_from_text_2nd_format(text)
    addr = replace_umlauts(address.upper())
    addr_parts = split_german_address(addr)
    return {
        "register_type": reg_type,       # e.g. HRB, HRA, VR, PR
        "register_number": reg_number,   # e.g. 12038
        "company_name": company,         # from 2.a) Firma:
        "address": addr,              # from 2.b) after Geschäftsanschrift:
    } | addr_parts  # merge address parts into the dict

def extract_from_text_2nd_format(text: str) -> dict:
    """
    Extract fields from Vereinsregister PDFs like:

    Ausdruck - Vereinsregister - VR 601 SE
    Aktueller Ausdruck VR 601 SE
    ...
    2.a) Name des Vereins
    <name lines...>
    b) Sitz des Vereins
    <city>
    <street house, plz city>
    """
    t = normalize_text(text)
    lines = [ln.strip() for ln in t.split("\n")]

    # 1) Register type & number (from 'Aktueller Ausdruck <TYPE> <NUMBER...>')
    line3 = lines[3]
    print(f"Line 4 Format 2: {line3}")  # Debugging output
    words = line3.split()
    if int(words[-2]):
        # sometimes register number is two words, e.g. "VR 601 SE" -> reg = "601 SE"
        w_type_raw = words[-3]
        w_num_raw = words[-2] + " " + words[-1]
    else:
        w_type_raw, w_num_raw = words[-2], words[-1]
    reg_type = re.sub(r"[^A-Za-zÄÖÜäöü]", "", w_type_raw).upper()
    reg_number = w_num_raw


    # 2) Company name: lines after "2.a)" up to "b)"
    company = ""
    idx_a = None
    for i, ln in enumerate(lines):
        if re.search(r"^\s*2\.\s*a\)", ln, flags=re.IGNORECASE):
            idx_a = i
            break
    if idx_a is not None:
        collected = []
        for ln in lines[idx_a+1:]:
            if re.search(r"^\s*(?:2\.\s*)?b\)", ln, flags=re.IGNORECASE):
                break
            if ln:
                collected.append(ln)
        company = " ".join(collected)
    else:
        return {reg_type: "unexpected Format"}

    # 3) Address: section 2.b) — take the “third row” pattern, or only the city (if no street, i.e Vereine often don't have one)
    address = ""
    idx_b = None
    for i, ln in enumerate(lines):
        if re.search(r"^\s*(?:2\.\s*)?b\)\s*", ln, flags=re.IGNORECASE):
            idx_b = i
            break
    if idx_b is not None:
        if re.search(r"^\s*3\.", lines[idx_b+2], flags=re.IGNORECASE):
            # only city
            city = lines[idx_b + 1].strip()
            return {
                "register_type": reg_type,
                "register_number": reg_number,
                "company_name": company,
                "city": city,
            }
        else:
            address = lines[idx_b+2].strip()
    else:
        return {reg_type: "unexpected Format"}


    parts = split_german_address(replace_umlauts(address.upper()))

    return {
        "register_type": reg_type,
        "register_number": reg_number,
        "company_name": company,
        "address": address,
        **parts,
    }


def extract_from_pdf(pdf_path: Path) -> dict:
    """
    Read all pages’ text (order-preserving) and run extractor once on full text.
    Most documents you showed keep everything on page 1, but this is safer.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages_text = []
            for page in pdf.pages:
                t = page.extract_text(x_tolerance=1.5, y_tolerance=1.5) or ""
                pages_text.append(t)
        full_text = "\n".join(pages_text)
        info = extract_from_text(full_text)
        info["file"] = str(pdf_path)
        return info
    except Exception as e:
        return {"file": str(pdf_path), "error": str(e),
                "register_type": "", "register_number": "", "company_name": "", "address": ""}

# --- batch runner ----------------------------------------------------------

def scan_directory(in_dir: str, out_csv: str | None = None) -> pd.DataFrame:
    """
    Scan all PDFs in a directory and return a DataFrame with the extracted fields.
    Optionally write a CSV for inspection.
    """
    in_dir = Path(in_dir)
    rows = []
    for p in sorted(in_dir.glob("*.pdf")):
        rows.append(extract_from_pdf(p))
    df = pd.DataFrame(rows, columns=["file", "register_type", "register_number", "company_name", "address", "error", "street", "house_number", "postal_code", "city"])
    if out_csv:
        Path(out_csv).parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(out_csv, index=False, encoding="utf-8")
    return df

# --- CLI -------------------------------------------------------------------

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Extract Handelsregister fields from PDFs.")
    ap.add_argument("--in", dest="in_dir", required=True, help="Folder with downloaded PDFs")
    ap.add_argument("--out", dest="out_csv", default=None, help="Optional CSV path to save results")
    args = ap.parse_args()

    df = scan_directory(args.in_dir, args.out_csv)
    # Print a compact preview to console
    for _, r in df.iterrows():
        print(f"\nFile: {r['file']}")
        if isinstance(r.get("error"), str) and r["error"]:
            print(f"  ERROR: {r['error']}")
            continue
        print(f"  Register: {r['register_type']} {r['register_number']}")
        print(f"  Company : {r['company_name']}")
        print(f"  Address : {r['address']}")
        print(f"  Street  : {r['street']}")
        print(f"  House # : {r['house_number']}")   
        print(f"  Postal Code: {r['postal_code']}")
        print(f"  City    : {r['city']}")

