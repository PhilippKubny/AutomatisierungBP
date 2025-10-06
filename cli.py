"""Command line entry point for interacting with NorthData providers."""

from __future__ import annotations

import argparse
import logging
import time
from collections.abc import Mapping, MutableMapping
from pathlib import Path

import yaml
from dotenv import load_dotenv

from bpauto import excel_io
from bpauto.providers import NorthDataProvider

logger = logging.getLogger(__name__)

DEFAULT_MAPPING: dict[str, str] = {
    "legal_name": "W",
    "register_type": "U",
    "register_no": "V",
    "street": "X",
    "zip": "Y",
    "city": "Z",
    "country": "AC",
    "pdf_path": "AA",
    "notes": "AB",
    "source": "AD",
}


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="BP Automation NorthData Integration")
    parser.add_argument("--excel", required=True, help="Pfad zur Excel-Arbeitsmappe")
    parser.add_argument("--sheet", required=True, help="Tabellenblatt-Name")
    parser.add_argument(
        "--start", type=int, default=3, help="Startzeile (1-basiert). Standard: 3."
    )
    parser.add_argument("--end", type=int, help="Endzeile (1-basiert, inklusiv)")
    parser.add_argument(
        "--name-col", default="C", help="Spalte mit Firmenname (Standard: C)"
    )
    parser.add_argument(
        "--zip-col", default=None, help="Spalte mit Postleitzahl (optional)"
    )
    parser.add_argument(
        "--country-col", default=None, help="Spalte mit Ländercode (optional)"
    )
    parser.add_argument(
        "--mapping-yaml",
        help="YAML mit Mapping zwischen CompanyRecord-Feldern und Spalten",
    )
    parser.add_argument("--source", default="api", choices=["api"], help="Quelle")
    parser.add_argument(
        "--download-ad",
        action="store_true",
        help="Amtlichen Auszug herunterladen und lokal speichern",
    )
    parser.add_argument(
        "--verbose", action="store_true", help="Ausführliche Log-Ausgabe aktivieren"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Nur Abruf, keine Schreiboperationen (nur für Tests)",
    )
    return parser.parse_args()


def _configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")


def _validate_column(column: str | None) -> str | None:
    if column is None:
        return None
    column = column.strip().upper()
    if not column:
        return None
    if not column.isalpha():
        raise ValueError(f"Ungültiger Spaltenwert: {column}")
    return column


def _load_mapping(path: str | None) -> dict[str, str]:
    mapping: MutableMapping[str, str] = dict(DEFAULT_MAPPING)
    if not path:
        return dict(mapping)

    mapping_path = Path(path)
    if not mapping_path.exists():
        raise FileNotFoundError(f"Mapping-Datei nicht gefunden: {mapping_path}")

    data = yaml.safe_load(mapping_path.read_text(encoding="utf-8"))
    if data is None:
        return dict(mapping)
    if not isinstance(data, Mapping):
        raise ValueError("Mapping YAML muss ein Dictionary enthalten")

    for key, value in data.items():
        str_key = str(key)
        if value is None:
            continue
        str_value = _validate_column(str(value))
        if str_value is None:
            continue
        mapping[str_key] = str_value

    return dict(mapping)


def main() -> int:
    args = _parse_args()
    _configure_logging(args.verbose)

    dotenv_path = Path(".env")
    if dotenv_path.exists():
        load_dotenv(dotenv_path=dotenv_path)

    mapping = _load_mapping(args.mapping_yaml)

    if args.source != "api":
        logger.error("Unbekannte Quelle: %s", args.source)
        return 2

    provider = NorthDataProvider(download_ad=args.download_ad)

    processed = 0
    hits = 0
    no_result = 0
    errors = 0
    start_time = time.perf_counter()

    try:
        name_column = _validate_column(args.name_col)
    except ValueError as exc:
        logger.error("%s", exc)
        return 2
    if not name_column:
        logger.error("Spalte für Firmennamen darf nicht leer sein")
        return 2

    try:
        zip_column = _validate_column(args.zip_col)
    except ValueError as exc:
        logger.error("%s", exc)
        return 2

    try:
        country_column = _validate_column(args.country_col)
    except ValueError as exc:
        logger.error("%s", exc)
        return 2

    rows = excel_io.iter_rows(
        excel_path=args.excel,
        sheet=args.sheet,
        start=args.start,
        end=args.end,
        name_col=name_column,
        zip_col=zip_column,
        country_col=country_column,
    )

    for row in rows:
        processed += 1
        name = row.get("name")
        zip_code = row.get("zip")
        country = row.get("country")

        if not name:
            logger.debug("Überspringe Zeile %s ohne Firmennamen", row.get("index"))
            continue

        try:
            record = provider.fetch(name=name, zip_code=zip_code, country=country)
        except RuntimeError as exc:
            logger.error("Abbruch wegen API-Fehler: %s", exc)
            return 3
        except Exception as exc:  # pragma: no cover - defensive
            logger.exception("Fehler bei der Abfrage für %s: %s", name, exc)
            errors += 1
            continue

        notes = (record.get("notes") or "").lower()
        if "no result" in notes:
            no_result += 1
        else:
            hits += 1

        if not args.dry_run:
            excel_io.write_result(
                excel_path=args.excel,
                sheet=args.sheet,
                row_index=row["index"],
                record=record,
                mapping=mapping,
            )

    if not args.dry_run:
        excel_io.save(args.excel)

    duration = time.perf_counter() - start_time
    logger.info(
        "Verarbeitung abgeschlossen: processed=%s hits=%s no_result=%s errors=%s duration=%.2fs",
        processed,
        hits,
        no_result,
        errors,
        duration,
    )

    if errors:
        return 4
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
