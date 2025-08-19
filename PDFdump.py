from pathlib import Path
import pdfplumber
import unicodedata
import os
import re

def dump_pdf_bytes(pdf_path: str, out_txt: str | None = None, width: int = 16) -> str:
    """
    Hex-dump a PDF's raw bytes.
    Each line shows: offset  hex bytes  ASCII (printables; others as '.').
    Example codes: space=0x20, tab=0x09, LF=0x0A, CR=0x0D.

    Args:
        pdf_path: path to the PDF.
        out_txt:  optional .txt file to also save the dump.
        width:    bytes per line (default 16).

    Returns:
        The dump as a single string.
    """
    with open(pdf_path, "rb") as f:
        data = f.read()

    lines = []
    for off in range(0, len(data), width):
        chunk = data[off:off+width]
        hex_part = " ".join(f"{b:02X}" for b in chunk)
        ascii_part = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        lines.append(f"{off:08X}  {hex_part:<{width*3}}  {ascii_part}")
    dump = "\n".join(lines)

    if out_txt:
        with open(out_txt, "w", encoding="utf-8") as fh:
            fh.write(dump)
    return dump


def dump_pdf_structure(pdf_path: str, out_txt: str | None = None, max_words: int | None = None, max_chars: int | None = None) -> str:
    """
    Dump the entire PDF structure (per page) to a string and optionally write to a .txt file.
    Shows: page size, extract_text(), words with boxes, chars with boxes/font/size, and object counts.

    Args:
        pdf_path: path to a PDF
        out_txt:  optional path to save the dump (recommended for large files)
        max_words: cap the number of word lines per page (None = no cap)
        max_chars: cap the number of char lines per page (None = no cap)

    Returns:
        The full dump as a single string.
    """
    def _norm(s: str) -> str:
        s = unicodedata.normalize("NFKC", s or "")
        s = s.replace("\u00A0", " ")
        s = s.replace("\r\n", "\n").replace("\r", "\n")
        s = re.sub(r"-\n(?=\w)", "", s)  # de-hyphenate across line breaks
        return s

    pdf_path = str(pdf_path)
    lines: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages, 1):
            lines.append(f"\n===== PAGE {i}/{total_pages} =====")
            lines.append(f"size: {page.width:.2f} x {page.height:.2f} pt")

            # 1) Full page text (what regexes usually see)
            text = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
            lines.append("\n-- extract_text() --")
            lines.append(_norm(text))

            # 2) Words with bounding boxes
            words = page.extract_words(x_tolerance=2, y_tolerance=2) or []
            lines.append(f"\n-- words (count={len(words)}) [x0,top,x1,bottom] text --")
            for idx, w in enumerate(words):
                if max_words is not None and idx >= max_words:
                    lines.append(f"... (truncated at {max_words} words)")
                    break
                lines.append(f"{w['x0']:.1f},{w['top']:.1f},{w['x1']:.1f},{w['bottom']:.1f}  {w['text']}")

            # 3) Characters with bounding boxes + font info
            chars = page.chars or []
            lines.append(f"\n-- chars (count={len(chars)}) [x0,top,x1,bottom] 'ch' font size --")
            for idx, c in enumerate(chars):
                if max_chars is not None and idx >= max_chars:
                    lines.append(f"... (truncated at {max_chars} chars)")
                    break
                lines.append(
                    f"{c['x0']:.1f},{c['top']:.1f},{c['x1']:.1f},{c['bottom']:.1f}  "
                    f"{repr(c['text'])}  {c.get('fontname','?')}  {c.get('size','?')}"
                )

            # 4) Graphics/objects summary (helps detect scanned PDFs)
            lines.append(
                f"\n-- objects -- lines={len(page.lines)} rects={len(page.rects)} curves={len(page.curves)} images={len(page.images)}"
            )

    dump = "\n".join(lines)
    if out_txt:
        Path(out_txt).parent.mkdir(parents=True, exist_ok=True)
        Path(out_txt).write_text(dump, encoding="utf-8")
    return dump


def scan_directory(in_dir: str, out_csv: str | None = None):
    """
    Scan all PDFs in a directory and return a DataFrame with the extracted fields.
    Optionally write a CSV for inspection.
    """
    out_csv = os.path.join(os.path.expanduser("~"), "Downloads", "HumanCheck.txt")

    out_cs = os.path.join(os.path.expanduser("~"), "Downloads", "HumanCheck.txt")
    in_dir = Path(in_dir)
    for p in sorted(in_dir.glob("*.pdf")):
        print(f"Processing {p.name}...")
        dump_pdf_structure(p, out_csv, max_words=10000, max_chars=50000)
        dump_pdf_bytes(p, out_cs, width=16)


# --- CLI -------------------------------------------------------------------

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Extract Handelsregister fields from PDFs.")
    ap.add_argument("--in", dest="in_dir", required=True, help="Folder with downloaded PDFs")
    ap.add_argument("--out", dest="out_csv", default=None, help="Optional CSV path to save results")
    args = ap.parse_args()
    print(f"Scanning directory: {args.in_dir}")
    scan_directory(args.in_dir, args.out_csv)
   