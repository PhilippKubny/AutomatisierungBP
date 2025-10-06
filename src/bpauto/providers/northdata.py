"""NorthData API provider implementation."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Dict, Optional
from urllib.parse import urlparse

import requests
from requests import Response
from tenacity import (RetryError, retry, retry_if_exception_type,
                      stop_after_attempt, wait_exponential)

from .base import CompanyRecord, Provider

logger = logging.getLogger(__name__)

_RETRYABLE_STATUS_CODES = {429, 500, 501, 502, 503, 504, 505}


class _RetryableRequestError(requests.RequestException):
    """Error raised for retryable HTTP status codes."""


class NorthDataProvider(Provider):
    """Provider that retrieves company information from the NorthData API."""

    base_url = "https://api.northdata.de/search"

    def __init__(
        self,
        api_key: Optional[str] = None,
        *,
        download_ad: bool = False,
        download_dir: Optional[str] = None,
        timeout: float = 15.0,
    ) -> None:
        self.api_key = api_key or os.getenv("NORTHDATA_API_KEY")
        if not self.api_key:
            raise RuntimeError("NORTHDATA_API_KEY environment variable is required")

        self._session_headers = {"X-Api-Key": self.api_key}
        self.download_ad = download_ad
        self._download_dir = Path(download_dir) if download_dir else Path.cwd()
        self._timeout = timeout

    def _should_retry(self, response: Response) -> bool:
        return response.status_code in _RETRYABLE_STATUS_CODES or 500 <= response.status_code < 600

    def _query_api(
        self,
        name: str,
        zip_code: Optional[str] = None,
        country: Optional[str] = None,
    ) -> Dict[str, object]:
        params: Dict[str, str] = {"query": name}
        if country:
            params["country"] = country
        if zip_code:
            params["postalCode"] = zip_code

        logger.debug("Querying NorthData API with params: %s", params)

        @retry(
            stop=stop_after_attempt(3),
            wait=wait_exponential(multiplier=1, min=1, max=8),
            retry=retry_if_exception_type(_RetryableRequestError),
            reraise=True,
        )
        def _perform_request() -> Dict[str, object]:
            try:
                response = requests.get(
                    self.base_url,
                    params=params,
                    headers=self._session_headers,
                    timeout=self._timeout,
                )
            except requests.RequestException as exc:  # pragma: no cover - network errors
                logger.error("NorthData API request failed: %s", exc)
                raise _RetryableRequestError(str(exc)) from exc

            if self._should_retry(response):
                logger.warning(
                    "NorthData API returned retryable status %s", response.status_code
                )
                raise _RetryableRequestError(
                    f"Retryable status code: {response.status_code}"
                )

            if response.status_code >= 400:
                logger.error(
                    "NorthData API returned error %s: %s",
                    response.status_code,
                    response.text,
                )
                response.raise_for_status()

            if not response.content:
                logger.debug("NorthData API response is empty")
                return {}

            try:
                return response.json()
            except ValueError as exc:  # pragma: no cover - depends on API response
                logger.error("Invalid JSON from NorthData API: %s", exc)
                raise _RetryableRequestError("Invalid JSON response") from exc

        try:
            return _perform_request()
        except RetryError as exc:  # pragma: no cover - network heavy
            logger.error("NorthData API request failed after retries: %s", exc)
            return {}

    def _best_match(
        self,
        raw: Dict[str, object],
        name: str,
        zip_code: Optional[str] = None,
    ) -> Dict[str, object]:
        results = []
        if isinstance(raw, dict):
            for key in ("result", "results", "hits", "data"):
                value = raw.get(key) if isinstance(raw.get(key), list) else None
                if value:
                    results = value
                    break

        best_entry: Dict[str, object] = {}
        best_score_tuple = (-1.0, -1.0)  # (zip_match, score)

        for entry in results:
            if not isinstance(entry, dict):
                continue

            score = 0.0
            for score_key in ("score", "confidence", "matchScore"):
                raw_score = entry.get(score_key)
                if isinstance(raw_score, (float, int)):
                    score = float(raw_score)
                    break

            entry_zip = (
                entry.get("postalCode")
                or entry.get("zip")
                or (entry.get("address", {}) if isinstance(entry.get("address"), dict) else {}).get("postalCode")
            )
            zip_match = 1.0 if zip_code and entry_zip and str(entry_zip) == str(zip_code) else 0.0

            candidate_tuple = (zip_match, score)
            if candidate_tuple > best_score_tuple:
                best_entry = entry
                best_score_tuple = candidate_tuple

        if not best_entry and isinstance(raw, dict):
            best_entry = raw

        return best_entry

    def _download_pdf(self, url: str, target_dir: str) -> str:
        if not url:
            return ""

        target_path = Path(target_dir)
        target_path.mkdir(parents=True, exist_ok=True)

        logger.debug("Downloading NorthData PDF from %s", url)

        try:
            response = requests.get(url, headers=self._session_headers, timeout=self._timeout)
        except requests.RequestException as exc:  # pragma: no cover - network errors
            logger.error("Failed to download PDF: %s", exc)
            return ""

        if response.status_code != 200:
            logger.warning(
                "Unable to download PDF, status %s", response.status_code
            )
            return ""

        parsed = urlparse(url)
        filename = os.path.basename(parsed.path) or "northdata_document.pdf"
        file_path = target_path / filename

        try:
            file_path.write_bytes(response.content)
        except OSError as exc:  # pragma: no cover - filesystem errors
            logger.error("Failed to write PDF file: %s", exc)
            return ""

        return str(file_path)

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
        try:
            raw = self._query_api(name=name, zip_code=zip_code, country=country)
        except Exception as exc:  # pragma: no cover - defensive
            logger.exception("Error querying NorthData API: %s", exc)
            return CompanyRecord(legal_name=name, notes="error querying northdata", source="northdata_api")

        if not raw:
            return CompanyRecord(legal_name=name, notes="no result", source="northdata_api")

        payload = self._best_match(raw, name=name, zip_code=zip_code)
        if not payload:
            return CompanyRecord(legal_name=name, notes="no result", source="northdata_api")

        record: CompanyRecord = CompanyRecord(source="northdata_api")

        legal_name = payload.get("legal_name") or payload.get("name") or name
        record["legal_name"] = str(legal_name)

        register = payload.get("register") if isinstance(payload.get("register"), dict) else {}
        if register:
            reg_type = register.get("type") or register.get("registerType")
            reg_no = register.get("number") or register.get("registerNumber")
            if reg_type:
                record["register_type"] = str(reg_type)
            if reg_no:
                record["register_no"] = str(reg_no)
        else:
            reg_type = payload.get("registerType")
            reg_no = payload.get("registerNumber")
            if reg_type:
                record["register_type"] = str(reg_type)
            if reg_no:
                record["register_no"] = str(reg_no)

        address = payload.get("address") if isinstance(payload.get("address"), dict) else {}
        if not address:
            address = {
                "street": payload.get("street"),
                "zip": payload.get("zip") or payload.get("postalCode"),
                "city": payload.get("city"),
                "country": payload.get("country"),
            }
        else:
            address = {
                "street": address.get("street") or address.get("streetName"),
                "zip": address.get("zip") or address.get("postalCode"),
                "city": address.get("city"),
                "country": address.get("country"),
            }

        for field in ("street", "zip", "city", "country"):
            value = address.get(field)
            if value:
                record[field] = str(value)

        pdf_url = (
            payload.get("officialExtractUrl")
            or payload.get("officialExtract")
            or payload.get("pdf")
            or payload.get("pdfUrl")
        )

        pdf_path = ""
        if pdf_url:
            if self.download_ad:
                pdf_path = self._download_pdf(pdf_url, str(self._download_dir))
            pdf_path = pdf_path or pdf_url
            record["pdf_path"] = pdf_path

        notes_parts = []
        for score_key in ("score", "confidence", "matchScore"):
            score_val = payload.get(score_key)
            if isinstance(score_val, (float, int)):
                notes_parts.append(f"confidence={score_val:.2f}")
                break

        if zip_code:
            payload_zip = (
                payload.get("postalCode")
                or payload.get("zip")
                or (payload.get("address", {}) if isinstance(payload.get("address"), dict) else {}).get("postalCode")
            )
            if payload_zip:
                match_status = "matched" if str(payload_zip) == str(zip_code) else "mismatched"
                notes_parts.append(f"zip_{match_status}={payload_zip}")

        if notes_parts:
            record["notes"] = ", ".join(notes_parts)

        return record
