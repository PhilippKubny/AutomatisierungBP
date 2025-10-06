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
except ImportError:  # pragma: no cover - openpyxl always provides Workbook
    Workbook = object  # type: ignore


def _get_or_load_workbook(excel_path: str) -> "Workbook":
    """Return a cached workbook instance for the given path."""

    workbook = _WORKBOOK_CACHE.get(excel_path)
    if workbook is None:
        logger.debug("Loading workbook: %s", excel_path)
        workbook = load_workbook(excel_path)
        _WORKBOOK_CACHE[excel_path] = workbook
    return workbook


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
    stop = end or worksheet.max_row

    def _cell_value(column: Optional[str], row_index: int) -> Optional[str]:
        if not column:
            return None
        column = column.strip().upper()
        if not column:
            return None
        cell = worksheet[f"{column}{row_index}"]
        value = cell.value
        return value if value is None or isinstance(value, str) else str(value)

    for row_idx in range(start, stop + 1):
        yield RowData(
            index=row_idx,
            name=_cell_value(name_col, row_idx),
            zip=_cell_value(zip_col, row_idx),
            country=_cell_value(country_col, row_idx),
        )


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

    for key, column in mapping.items():
        column = column.strip().upper()
        if not column:
            continue
        value = record.get(key)
        if value is None:
            continue
        cell = worksheet[f"{column}{row_index}"]
        cell.value = value


def save_workbook(excel_path: str) -> None:
    """Persist any pending changes to disk."""

    workbook = _WORKBOOK_CACHE.get(excel_path)
    if workbook is None:
        logger.debug("No workbook cached for path: %s", excel_path)
        return
    logger.debug("Saving workbook: %s", excel_path)
    workbook.save(excel_path)
