"""Command line entry point for interacting with NorthData providers."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Dict, Optional

from dotenv import load_dotenv

from bpauto import excel_io
from bpauto.providers.northdata import NorthDataProvider

logger = logging.getLogger(__name__)

DEFAULT_MAPPING: Dict[str, str] = {
    "register_type": "U",
    "register_no": "V",
    "legal_name": "W",
    "street": "X",
    "zip": "Y",
    "city": "Z",
    "pdf_path": "AA",
    "notes": "AB",
}


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="BP Automation NorthData integration")
    parser.add_argument("--excel", required=True, help="Path to the Excel workbook")
    parser.add_argument("--sheet", required=True, help="Worksheet name")
    parser.add_argument(
        "--start",
        type=int,
        default=3,
        help="Start row index (1-based). Defaults to 3.",
    )
    parser.add_argument("--end", type=int, help="End row index (1-based, inclusive)")
    parser.add_argument("--name-col", default="C", help="Column letter containing the company name")
    parser.add_argument("--zip-col", default=None, help="Column letter containing the ZIP code")
    parser.add_argument(
        "--country-col",
        default=None,
        help="Column letter containing the country",
    )
    parser.add_argument(
        "--mapping-yaml",
        help="YAML file mapping CompanyRecord fields to Excel column letters",
    )
    parser.add_argument("--source", default="api", choices=["api"], help="Data source to use")
    parser.add_argument(
        "--download-ad",
        action="store_true",
        help="Download additional documents such as official register PDFs",
    )
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging output")
    return parser.parse_args()


def _configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")


def _load_mapping(path: Optional[str]) -> Dict[str, str]:
    if not path:
        return dict(DEFAULT_MAPPING)

    mapping_path = Path(path)
    if not mapping_path.exists():
        raise FileNotFoundError(f"Mapping file not found: {mapping_path}")

    try:
        import yaml  # type: ignore
    except ImportError:
        yaml = None

    content = mapping_path.read_text(encoding="utf-8")
    if not content.strip():
        return dict(DEFAULT_MAPPING)

    if yaml:
        data = yaml.safe_load(content)  # type: ignore[attr-defined]
        if not isinstance(data, dict):
            raise ValueError("Mapping YAML must contain a dictionary")
        normalised = {str(k): str(v) for k, v in data.items()}
        return {**DEFAULT_MAPPING, **normalised}

    # Minimal fallback parser for simple key: value lines
    mapping: Dict[str, str] = {}
    for line in content.splitlines():
        if not line.strip() or line.strip().startswith("#"):
            continue
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        mapping[key.strip()] = value.strip()
    return {**DEFAULT_MAPPING, **mapping} if mapping else dict(DEFAULT_MAPPING)


def main() -> int:
    args = _parse_args()
    _configure_logging(args.verbose)
    dotenv_path = Path(".env")
    if dotenv_path.exists():
        load_dotenv(str(dotenv_path))

    mapping = _load_mapping(args.mapping_yaml)

    if args.source != "api":
        raise ValueError(f"Unsupported source: {args.source}")

    provider = NorthDataProvider(download_ad=args.download_ad)

    processed = 0
    success = 0
    no_result = 0

    for row in excel_io.iter_rows(
        excel_path=args.excel,
        sheet=args.sheet,
        start=args.start,
        end=args.end,
        name_col=args.name_col,
        zip_col=args.zip_col,
        country_col=args.country_col,
    ):
        name = row.get("name")
        if not name:
            logger.debug("Skipping row %s: no company name", row["index"])
            continue

        processed += 1

        try:
            record = provider.fetch(
                name=name,
                zip_code=row.get("zip"),
                country=row.get("country"),
            )
        except Exception as exc:  # pragma: no cover - runtime safety
            logger.error("Failed to fetch record for %s: %s", name, exc)
            continue

        notes_value = (record.get("notes") or "").strip().lower()
        if notes_value == "no result":
            no_result += 1
        else:
            success += 1

        excel_io.write_result(
            excel_path=args.excel,
            sheet=args.sheet,
            row_index=row["index"],
            record=record,
            mapping=mapping,
        )

    excel_io.save(args.excel)

    logger.info(
        "Processing summary: processed=%s, success=%s, no_result=%s",
        processed,
        success,
        no_result,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
