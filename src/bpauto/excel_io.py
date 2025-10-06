"""Hilfsfunktionen für das Lesen und Schreiben von Excel-Arbeitsmappen."""

from __future__ import annotations

import logging
from collections.abc import Iterator
from typing import TypedDict, cast

from openpyxl import load_workbook

from .providers.base import CompanyRecord

logger = logging.getLogger(__name__)

try:  # pragma: no cover - ``openpyxl`` always provides these during runtime
    from openpyxl.workbook import Workbook
    from openpyxl.worksheet.worksheet import Worksheet
except ImportError:  # pragma: no cover - fallback für statische Analyse
    Workbook = object  # type: ignore[assignment]
    Worksheet = object  # type: ignore[assignment]


class RowData(TypedDict, total=False):
    """Representation einer gelesenen Tabellenzeile."""

    index: int
    name: str | None
    zip: str | None
    country: str | None


_WORKBOOK_CACHE: dict[str, Workbook] = {}


def _get_or_load_workbook(excel_path: str) -> Workbook:
    """Gibt eine zwischengespeicherte Arbeitsmappe zurück."""

    workbook = _WORKBOOK_CACHE.get(excel_path)
    if workbook is None:
        logger.debug("Lade Arbeitsmappe: %s", excel_path)
        workbook = load_workbook(excel_path)
        _WORKBOOK_CACHE[excel_path] = workbook
    return workbook


def _normalise_column(column: str | None) -> str | None:
    if column is None:
        return None
    column = column.strip().upper()
    return column or None


def _cell_to_string(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        value_str = value.strip()
    else:
        value_str = str(value).strip()
    return value_str or None


def _read_cell(worksheet: Worksheet, column: str | None, row_index: int) -> str | None:
    column = _normalise_column(column)
    if not column:
        return None
    cell = worksheet[f"{column}{row_index}"]
    return _cell_to_string(cell.value)


def _find_last_row_with_name(
    worksheet: Worksheet, name_column: str, start_row: int
) -> int:
    normalised_name = _normalise_column(name_column)
    if not normalised_name:
        raise ValueError("Name column must be provided")

    max_row = worksheet.max_row
    for row_idx in range(max_row, start_row - 1, -1):
        value = _read_cell(worksheet, normalised_name, row_idx)
        if value is not None:
            return row_idx
    return start_row - 1


def _get_worksheet(workbook: Workbook, sheet: str) -> Worksheet:
    try:
        return workbook[sheet]
    except KeyError as exc:  # pragma: no cover - defensive
        raise ValueError(f"Arbeitsblatt '{sheet}' wurde nicht gefunden") from exc


def iter_rows(
    excel_path: str,
    sheet: str,
    start: int,
    end: int | None,
    name_col: str,
    zip_col: str | None = None,
    country_col: str | None = None,
) -> Iterator[RowData]:
    """Liest Zeilen aus der Arbeitsmappe und liefert bereinigte Werte."""

    workbook = _get_or_load_workbook(excel_path)
    worksheet = _get_worksheet(workbook, sheet)
    normalised_name_col = _normalise_column(name_col)
    if not normalised_name_col:
        raise ValueError("Name column must be provided")

    stop = end if end is not None else _find_last_row_with_name(
        worksheet, normalised_name_col, start
    )

    logger.info(
        "Lese Zeilen %s-%s aus Blatt '%s' (%s)", start, stop, sheet, excel_path
    )

    def _generator() -> Iterator[RowData]:
        yielded = 0
        if stop < start:
            logger.info(
                "Keine Datenzeilen in Blatt '%s' (%s) gefunden", sheet, excel_path
            )
            return

        for row_idx in range(start, stop + 1):
            name_value = _read_cell(worksheet, normalised_name_col, row_idx)
            if name_value is None:
                continue

            row_data: RowData = RowData(
                index=row_idx,
                name=name_value,
                zip=_read_cell(worksheet, zip_col, row_idx),
                country=_read_cell(worksheet, country_col, row_idx),
            )
            yielded += 1
            yield row_data

        logger.info(
            "Verarbeitete Zeilen in Blatt '%s' (%s): %s",
            sheet,
            excel_path,
            yielded,
        )

    return _generator()


def write_result(
    excel_path: str,
    sheet: str,
    row_index: int,
    record: CompanyRecord,
    mapping: dict[str, str],
) -> None:
    """Schreibt Daten aus *record* in die gemappten Spalten."""

    workbook = _get_or_load_workbook(excel_path)
    worksheet = _get_worksheet(workbook, sheet)

    logger.debug(
        "Schreibe Ergebnis für Zeile %s in Blatt '%s' (%s)",
        row_index,
        sheet,
        excel_path,
    )

    record_dict = cast(dict[str, object | None], record)

    for key, column in mapping.items():
        column_letter = _normalise_column(column)
        if not column_letter or key not in record_dict:
            continue
        value = record_dict.get(key)
        if value is None:
            cell_value = ""
        elif isinstance(value, str):
            cell_value = value
        else:
            cell_value = str(value)
        worksheet[f"{column_letter}{row_index}"] = cell_value


def save(excel_path: str) -> None:
    """Persistiert Änderungen auf die Festplatte."""

    workbook = _WORKBOOK_CACHE.get(excel_path)
    if workbook is None:
        logger.debug("Keine Arbeitsmappe im Cache für Pfad: %s", excel_path)
        return
    logger.info("Speichere Arbeitsmappe: %s", excel_path)
    workbook.save(excel_path)


def reset() -> None:
    """Leert den Arbeitsmappen-Cache (hauptsächlich für Tests)."""

    logger.debug("Leere Arbeitsmappen-Cache")
    _WORKBOOK_CACHE.clear()
