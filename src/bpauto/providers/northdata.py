"""NorthData API provider implementation."""

from __future__ import annotations

import logging
import os
from typing import Dict, Optional

from .base import CompanyRecord, Provider

logger = logging.getLogger(__name__)


class NorthDataProvider(Provider):
    """Provider that retrieves company information from the NorthData API."""

    def __init__(self, api_key: Optional[str] = None) -> None:
        self.api_key = api_key or os.getenv("NORTHDATA_API_KEY")
        if not self.api_key:
            raise RuntimeError("NORTHDATA_API_KEY environment variable is required")
        self._session_headers = {"Authorization": f"Bearer {self.api_key}"}

    def _query_api(
        self,
        name: str,
        zip_code: Optional[str] = None,
        country: Optional[str] = None,
    ) -> Dict[str, object]:
        """Placeholder for the actual API query."""

        raise NotImplementedError("NorthData API integration not yet implemented")

    def fetch(
        self,
        name: str,
        zip_code: Optional[str] = None,
        country: Optional[str] = None,
    ) -> CompanyRecord:
        logger.debug(
            "Fetching NorthData record for name=%s, zip=%s, country=%s",
            name,
            zip_code,
            country,
        )
        payload = self._query_api(name=name, zip_code=zip_code, country=country)

        record: CompanyRecord = CompanyRecord(source="northdata_api")

        legal_name = payload.get("legal_name") or payload.get("name")
        if legal_name:
            record["legal_name"] = str(legal_name)

        register = payload.get("register") if isinstance(payload.get("register"), dict) else {}
        if register:
            reg_type = register.get("type")
            reg_no = register.get("number")
            if reg_type:
                record["register_type"] = str(reg_type)
            if reg_no:
                record["register_no"] = str(reg_no)

        address = payload.get("address") if isinstance(payload.get("address"), dict) else {}
        if address:
            for field in ("street", "zip", "city", "country"):
                value = address.get(field)
                if value:
                    record[field] = str(value)

        pdf_path = payload.get("pdf_path") or payload.get("pdf")
        if pdf_path:
            record["pdf_path"] = str(pdf_path)

        notes = payload.get("notes") or payload.get("comment")
        if notes:
            record["notes"] = str(notes)

        return record
