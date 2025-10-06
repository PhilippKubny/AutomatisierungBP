"""Utilities for reading from and writing to Excel workbooks."""

from __future__ import annotations

import logging
from typing import Dict, Generator, Optional, TypedDict

from openpyxl import load_workbook
from .providers.base import CompanyRecord

logger = logging.getLogger(__name__)


class RowData(TypedDict):
    index: int
    name: Optional[str]
    zip: Optional[str]
    country: Optional[str]


_WORKBOOK_CACHE: Dict[str, "Workbook"] = {}

try:  # Lazy import for typing only
    from openpyxl.workbook import Workbook
    from openpyxl.worksheet.worksheet import Worksheet
except ImportError:  # pragma: no cover - openpyxl always provides Workbook
    Workbook = object  # type: ignore
    Worksheet = object  # type: ignore


def _get_or_load_workbook(excel_path: str) -> "Workbook":
    """Return a cached workbook instance for the given path."""

    workbook = _WORKBOOK_CACHE.get(excel_path)
    if workbook is None:
        logger.debug("Loading workbook: %s", excel_path)
        workbook = load_workbook(excel_path)
        _WORKBOOK_CACHE[excel_path] = workbook
    return workbook


def _normalise_column(column: Optional[str]) -> Optional[str]:
    if column is None:
        return None
    column = column.strip().upper()
    return column or None


def _cell_to_string(value: object) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, str):
        value_str = value.strip()
    else:
        value_str = str(value).strip()
    return value_str or None


def _read_cell(worksheet: "Worksheet", column: Optional[str], row_index: int) -> Optional[str]:
    column = _normalise_column(column)
    if not column:
        return None
    cell = worksheet[f"{column}{row_index}"]
    return _cell_to_string(cell.value)


def _find_last_row_with_name(
    worksheet: "Worksheet", name_column: str, start_row: int
) -> int:
    name_column = _normalise_column(name_column)
    if not name_column:
        raise ValueError("Name column must be provided")

    max_row = worksheet.max_row
    for row_idx in range(max_row, start_row - 1, -1):
        value = _read_cell(worksheet, name_column, row_idx)
        if value is not None:
            return row_idx
    return start_row - 1


def iter_rows(
    excel_path: str,
    sheet: str,
    start: int,
    end: Optional[int],
    name_col: str,
    zip_col: Optional[str] = None,
    country_col: Optional[str] = None,
) -> Generator[RowData, None, None]:
    """Yield row information between *start* and *end* (inclusive)."""

    workbook = _get_or_load_workbook(excel_path)
    worksheet = workbook[sheet]
    normalised_name_col = _normalise_column(name_col)
    if not normalised_name_col:
        raise ValueError("Name column must be provided")

    if end is not None:
        stop = end
    else:
        stop = _find_last_row_with_name(worksheet, normalised_name_col, start)

    logger.info(
        "Iterating rows from %s to %s on sheet '%s' in workbook '%s'",
        start,
        stop,
        sheet,
        excel_path,
    )

    def _generator() -> Generator[RowData, None, None]:
        yielded = 0
        if stop < start:
            logger.info(
                "No rows to iterate for sheet '%s' in workbook '%s'", sheet, excel_path
            )
            return

        for row_idx in range(start, stop + 1):
            name_value = _read_cell(worksheet, normalised_name_col, row_idx)
            if name_value is None:
                continue

            row_data: RowData = {
                "index": row_idx,
                "name": name_value,
                "zip": _read_cell(worksheet, zip_col, row_idx),
                "country": _read_cell(worksheet, country_col, row_idx),
            }
            yielded += 1
            yield row_data

        logger.info(
            "Iterated %s row(s) from sheet '%s' in workbook '%s'", yielded, sheet, excel_path
        )

    return _generator()


def write_result(
    excel_path: str,
    sheet: str,
    row_index: int,
    record: CompanyRecord,
    mapping: Dict[str, str],
) -> None:
    """Write *record* to the worksheet using the provided column *mapping*."""

    workbook = _get_or_load_workbook(excel_path)
    worksheet = workbook[sheet]

    logger.info(
        "Writing result for row %s on sheet '%s' in workbook '%s'",
        row_index,
        sheet,
        excel_path,
    )

    for key, column in mapping.items():
        column_letter = _normalise_column(column)
        if not column_letter:
            continue
        value = record.get(key)
        if value is None:
            cell_value = ""
        elif isinstance(value, str):
            cell_value = value
        else:
            cell_value = str(value)
        worksheet[f"{column_letter}{row_index}"] = cell_value


def save(excel_path: str) -> None:
    """Persist any pending changes to disk."""

    workbook = _WORKBOOK_CACHE.get(excel_path)
    if workbook is None:
        logger.debug("No workbook cached for path: %s", excel_path)
        return
    logger.debug("Saving workbook: %s", excel_path)
    workbook.save(excel_path)


def reset() -> None:
    """Clear the workbook cache. Intended for test usage."""

    logger.debug("Resetting workbook cache")
    _WORKBOOK_CACHE.clear()
