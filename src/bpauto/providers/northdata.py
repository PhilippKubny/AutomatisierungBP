"""NorthData API provider implementation."""

from __future__ import annotations

import json
import os
import re
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, cast
from urllib.parse import urlparse

import requests
from requests import Response
from tenacity import (
    RetryCallState,
    RetryError,
    Retrying,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)

from ..utils.logging_setup import setup_logger
from .base import CompanyRecord, Provider

_BASE_LOGGER = setup_logger()
LOGGER = _BASE_LOGGER.getChild("providers.northdata")

_RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}
_DEFAULT_TIMEOUT = (5.0, 30.0)


class _RetryableRequestError(RuntimeError):
    """Error raised for retryable HTTP status codes."""

    def __init__(self, message: str, *, status_code: int | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code


def _log_retry(retry_state: RetryCallState) -> None:
    exc = retry_state.outcome.exception() if retry_state.outcome else None
    retry_obj = retry_state.retry_object
    max_attempts = "?"
    if isinstance(retry_obj, Retrying):
        stop = getattr(retry_obj, "stop", None)
        max_attempts = getattr(stop, "max_attempt_number", "?")
    if isinstance(exc, _RetryableRequestError):
        LOGGER.warning(
            "Retry NorthData API (Versuch %s/%s) nach Status %s",
            retry_state.attempt_number,
            max_attempts,
            exc.status_code,
        )
    else:  # pragma: no cover - rein defensiv
        LOGGER.warning(
            "Retry NorthData API (Versuch %s/%s) wegen %s",
            retry_state.attempt_number,
            max_attempts,
            exc,
        )


def _slugify(value: str) -> str:
    value = value.strip().lower()
    value = re.sub(r"[^a-z0-9]+", "-", value)
    value = re.sub(r"-+", "-", value)
    return value.strip("-") or "northdata-document"


@dataclass(frozen=True)
class _Candidate:
    index: int
    payload: dict[str, Any]
    is_exact_match: bool
    zip_matches: bool
    city_matches: bool
    score: float


class NorthDataProvider(Provider):
    """Provider that retrieves company information from the NorthData API."""

    base_url = "https://www.northdata.de/_api/company/v1/company"

    def __init__(
        self,
        api_key: str | None = None,
        *,
        download_ad: bool = False,
        download_dir: str | None = None,
        timeout: Sequence[float] | None = None,
    ) -> None:
        self.api_key = api_key or os.getenv("NORTHDATA_API_KEY")
        if not self.api_key:
            raise RuntimeError(
                "NORTHDATA_API_KEY muss gesetzt sein (Umgebungsvariable oder Parameter)."
            )

        self._session_headers = {"X-Api-Key": self.api_key}
        self.download_ad = download_ad
        base_download_dir = Path(download_dir) if download_dir else Path.cwd()
        self._download_dir = base_download_dir / ".artifacts" / "extracts"

        if timeout is None:
            self._timeout = _DEFAULT_TIMEOUT
        elif len(timeout) == 1:
            self._timeout = (float(timeout[0]), _DEFAULT_TIMEOUT[1])
        else:
            connect, read = float(timeout[0]), float(timeout[1])
            self._timeout = (connect, read)

    def _should_retry(self, response: Response) -> bool:
        return response.status_code in _RETRYABLE_STATUS_CODES

    def _perform_request(self, params: dict[str, str]) -> dict[str, Any]:
        retrying = Retrying(
            stop=stop_after_attempt(5),
            wait=wait_exponential(multiplier=0.5, min=0.5, max=10),
            retry=retry_if_exception_type(_RetryableRequestError),
            after=_log_retry,
        )

        for attempt in retrying:
            with attempt:
                return self._perform_request_once(params)
        return {}

    def _perform_request_once(self, params: dict[str, str]) -> dict[str, Any]:
        try:
            response = requests.get(
                self.base_url,
                params=params,
                headers=self._session_headers,
                timeout=self._timeout,
            )
        except requests.RequestException as exc:  # pragma: no cover - Netzwerkfehler
            raise _RetryableRequestError(str(exc)) from exc

        if response.status_code in {401, 403}:
            raise RuntimeError(
                f"NorthData API Key ungültig oder fehlend (HTTP {response.status_code})."
            )

        if response.status_code == 404:
            LOGGER.info("NorthData API meldet: keine Ergebnisse (404)")
            return {}

        if self._should_retry(response):
            raise _RetryableRequestError(
                f"Retryable status code: {response.status_code}",
                status_code=response.status_code,
            )

        if response.status_code >= 400:
            LOGGER.error(
                "NorthData API Fehler %s: %s",
                response.status_code,
                response.text,
            )
            response.raise_for_status()

        if not response.content:
            return {}

        try:
            payload = response.json()
        except json.JSONDecodeError as exc:
            raise _RetryableRequestError("Ungültige JSON-Antwort") from exc

        if isinstance(payload, dict):
            return cast(dict[str, Any], payload)

        LOGGER.debug("JSON-Antwort hat unerwartetes Format: %s", type(payload))
        return {}

    def _query_api(
        self,
        name: str,
        zip_code: str | None = None,
        *,
        city: str | None = None,
        country: str | None = None,
    ) -> dict[str, Any]:
        params: dict[str, str] = {"name": name}
        address_parts: list[str] = []
        if zip_code:
            address_parts.append(str(zip_code))
        if city:
            address_parts.append(city)
        if country:
            address_parts.append(country)
        if address_parts:
            params["address"] = " ".join(address_parts)

        LOGGER.debug("Frage NorthData API mit Parametern: %s", params)

        try:
            return self._perform_request(params)
        except RetryError as exc:
            last_exc = exc.last_attempt.exception()
            if isinstance(last_exc, _RetryableRequestError):
                LOGGER.error(
                    "NorthData API mehrfach fehlgeschlagen (Status %s)",
                    last_exc.status_code,
                )
            else:
                LOGGER.error("NorthData API mehrfach fehlgeschlagen: %s", exc)
            return {}

    @staticmethod
    def _extract_results(raw: dict[str, Any]) -> Iterable[dict[str, Any]]:
        for key in ("result", "results", "hits", "data"):
            value = raw.get(key)
            if isinstance(value, list):
                yield from (entry for entry in value if isinstance(entry, dict))
                return
        if raw:
            yield raw

    @staticmethod
    def _normalise_name(value: str | None) -> str | None:
        if value is None:
            return None
        return str(value).strip().casefold() or None

    def _candidate_from_entry(
        self,
        entry: dict[str, Any],
        *,
        query_name: str,
        query_zip: str | None,
        query_city: str | None,
        index: int,
    ) -> _Candidate:
        def _float_score(keys: Sequence[str]) -> float:
            for key in keys:
                raw_score = entry.get(key)
                if isinstance(raw_score, float | int):
                    return float(raw_score)
            return 0.0

        raw_name = entry.get("legalName") or entry.get("legal_name") or entry.get("name")
        if isinstance(raw_name, dict):
            entry_name = raw_name.get("name") or raw_name.get("legalName")
        else:
            entry_name = raw_name
        normalised_entry_name = self._normalise_name(entry_name)
        query_name_norm = self._normalise_name(query_name)
        is_exact_match = bool(query_name_norm and normalised_entry_name == query_name_norm)

        address_block = entry.get("address") if isinstance(entry.get("address"), dict) else {}
        entry_zip = entry.get("postalCode") or entry.get("zip")
        if not entry_zip and isinstance(address_block, dict):
            entry_zip = address_block.get("postalCode") or address_block.get("zip")
        zip_matches = bool(
            query_zip and entry_zip and str(entry_zip).strip() == str(query_zip).strip()
        )

        entry_city = entry.get("city")
        if not entry_city and isinstance(address_block, dict):
            entry_city = address_block.get("city")
        city_matches = bool(
            query_city and entry_city and str(entry_city).strip().casefold()
            == str(query_city).strip().casefold()
        )

        score = _float_score(["score", "confidence", "matchScore"])
        return _Candidate(
            index=index,
            payload=entry,
            is_exact_match=is_exact_match,
            zip_matches=zip_matches,
            city_matches=city_matches,
            score=score,
        )

    def _best_match(
        self,
        raw: dict[str, Any],
        name: str,
        zip_code: str | None = None,
        city: str | None = None,
    ) -> dict[str, Any] | None:
        candidates: list[_Candidate] = []
        for idx, entry in enumerate(self._extract_results(raw)):
            candidates.append(
                self._candidate_from_entry(
                    entry,
                    query_name=name,
                    query_zip=zip_code,
                    query_city=city,
                    index=idx,
                )
            )

        if not candidates:
            return None

        if city:
            city_matches = [c for c in candidates if c.city_matches]
            if city_matches:
                candidates = city_matches

        if zip_code:
            zip_matches = [c for c in candidates if c.zip_matches]
            if zip_matches:
                candidates = zip_matches

        exact_matches = [c for c in candidates if c.is_exact_match]
        pool = exact_matches or candidates

        best = max(pool, key=lambda c: (c.score, -c.index))
        return best.payload

    def _download_pdf(
        self,
        url: str,
        target_dir: Path,
        *,
        name: str,
        register_no: str | None,
    ) -> str:
        if not url:
            return ""

        target_dir.mkdir(parents=True, exist_ok=True)
        try:
            response = requests.get(
                url,
                headers=self._session_headers,
                timeout=self._timeout,
            )
        except requests.RequestException as exc:  # pragma: no cover - Netzwerkfehler
            LOGGER.error("PDF-Download fehlgeschlagen: %s", exc)
            return ""

        if response.status_code != 200:
            LOGGER.warning("PDF-Download nicht möglich, Status %s", response.status_code)
            return ""

        parsed = urlparse(url)
        extension = os.path.splitext(os.path.basename(parsed.path))[1] or ".pdf"
        slug_source = "-".join(filter(None, [name, register_no or ""])) or name
        slug = _slugify(slug_source)

        file_path = target_dir / f"{slug}{extension}"
        counter = 1
        while file_path.exists():
            file_path = target_dir / f"{slug}-{counter}{extension}"
            counter += 1

        try:
            file_path.write_bytes(response.content)
        except OSError as exc:  # pragma: no cover - Dateisystemfehler
            LOGGER.error("PDF-Datei konnte nicht geschrieben werden: %s", exc)
            return ""

        return str(file_path)

    def _normalise_address(self, payload: dict[str, Any]) -> dict[str, Any]:
        address_raw = payload.get("address")
        if isinstance(address_raw, dict):
            return {
                "street": address_raw.get("street") or address_raw.get("streetName"),
                "zip": address_raw.get("zip") or address_raw.get("postalCode"),
                "city": address_raw.get("city"),
                "country": address_raw.get("country"),
            }
        return {
            "street": payload.get("street"),
            "zip": payload.get("zip") or payload.get("postalCode"),
            "city": payload.get("city"),
            "country": payload.get("country"),
        }

    def fetch(
        self,
        name: str,
        zip_code: str | None = None,
        *,
        city: str | None = None,
        country: str | None = None,
    ) -> CompanyRecord:
        LOGGER.info(
            "Rufe NorthData-Daten ab für name=%s zip=%s city=%s country=%s",
            name,
            zip_code,
            city,
            country,
        )
        try:
            raw = self._query_api(
                name=name,
                zip_code=zip_code,
                city=city,
                country=country,
            )
        except RuntimeError:
            raise
        except Exception as exc:  # pragma: no cover - defensive
            LOGGER.exception("Fehler beim NorthData-Request: %s", exc)
            return CompanyRecord(
                legal_name=name,
                notes="fehler bei anfrage",
                source="northdata_api",
            )

        if not raw:
            LOGGER.warning("NorthData: keine Daten für %s (%s)", name, zip_code)
            return CompanyRecord(
                legal_name=name,
                notes="no result",
                source="northdata_api",
            )

        payload = self._best_match(raw, name=name, zip_code=zip_code, city=city)
        if not payload:
            LOGGER.warning("NorthData: keine verwertbaren Treffer für %s", name)
            return CompanyRecord(
                legal_name=name,
                notes="no result",
                source="northdata_api",
            )

        record: CompanyRecord = CompanyRecord(source="northdata_api")

        raw_name = (
            payload.get("legalName")
            or payload.get("legal_name")
            or payload.get("name")
            or name
        )
        if isinstance(raw_name, dict):
            legal_name = raw_name.get("name") or raw_name.get("legalName") or name
        else:
            legal_name = raw_name
        record["legal_name"] = str(legal_name)

        register_info = payload.get("register")
        if isinstance(register_info, dict):
            reg_type = (
                register_info.get("type")
                or register_info.get("registerType")
                or register_info.get("category")
            )
            reg_no = (
                register_info.get("number")
                or register_info.get("registerNumber")
                or register_info.get("id")
            )
            if not reg_type and isinstance(reg_no, str):
                reg_type = reg_no.split()[0] if reg_no.strip() else None
        else:
            reg_type = payload.get("registerType")
            reg_no = payload.get("registerNumber")
        if reg_type:
            record["register_type"] = str(reg_type)
        if reg_no:
            reg_no_str = str(reg_no)
            if (
                reg_type
                and isinstance(reg_no, str)
                and reg_no_str.strip().upper().startswith(str(reg_type).strip().upper())
            ):
                reg_no_clean = reg_no_str.strip()[len(str(reg_type).strip()):].strip()
                record["register_no"] = reg_no_clean or reg_no_str.strip()
            else:
                record["register_no"] = reg_no_str

        address = self._normalise_address(payload)
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
        if isinstance(pdf_url, str) and pdf_url:
            if self.download_ad:
                pdf_path = self._download_pdf(
                    pdf_url,
                    self._download_dir,
                    name=record.get("legal_name", name),
                    register_no=record.get("register_no"),
                )
            record["pdf_path"] = pdf_path or pdf_url

        notes_parts: list[str] = []
        for score_key in ("score", "confidence", "matchScore"):
            score_val = payload.get(score_key)
            if isinstance(score_val, float | int):
                notes_parts.append(f"confidence={float(score_val):.2f}")
                break

        if zip_code:
            payload_zip = (
                payload.get("postalCode")
                or payload.get("zip")
                or (
                    payload.get("address", {}) if isinstance(payload.get("address"), dict) else {}
                ).get("postalCode")
            )
            if payload_zip:
                match_status = (
                    "matched" if str(payload_zip).strip() == str(zip_code).strip() else "mismatched"
                )
                notes_parts.append(f"zip_{match_status}={payload_zip}")

        if not notes_parts:
            notes_parts.append("fetched")

        record["notes"] = ", ".join(notes_parts)
        return record
