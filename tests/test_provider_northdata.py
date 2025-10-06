from typing import Any

import pytest

from bpauto.providers.northdata import NorthDataProvider


class _DummyResponse:
    status_code = 200

    def __init__(self, payload: dict[str, Any]) -> None:
        self._payload = payload
        self.content = b"{}"

    def json(self) -> dict[str, object]:
        return self._payload

    def raise_for_status(self) -> None:  # pragma: no cover - compatibility only
        return None


@pytest.fixture
def fake_response() -> _DummyResponse:
    payload = {
        "result": [
            {
                "legalName": "Example GmbH",
                "register": {"type": "HRB", "number": "12345"},
                "address": {
                    "street": "Musterstraße 1",
                    "postalCode": "80333",
                    "city": "München",
                    "country": "DE",
                },
                "score": 0.91,
            }
        ]
    }
    return _DummyResponse(payload)


def test_requires_api_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("NORTHDATA_API_KEY", raising=False)
    with pytest.raises(RuntimeError):
        NorthDataProvider(download_ad=False)


def test_fetch_transforms_payload(
    monkeypatch: pytest.MonkeyPatch, fake_response: _DummyResponse
) -> None:
    monkeypatch.setenv("NORTHDATA_API_KEY", "dummy-key")

    captured_params: list[dict[str, str]] = []

    def fake_get(
        url: str,
        params: dict[str, str],
        headers: dict[str, str],
        timeout: tuple[float, float],
    ) -> _DummyResponse:
        captured_params.append(params)
        return fake_response

    monkeypatch.setattr("bpauto.providers.northdata.requests.get", fake_get)

    provider = NorthDataProvider(download_ad=False)
    record = provider.fetch(name="Example GmbH", zip_code="80333", country="DE")

    assert record["legal_name"] == "Example GmbH"
    assert record["register_type"] == "HRB"
    assert record["register_no"] == "12345"
    assert record["zip"] == "80333"
    assert record["city"] == "München"
    assert record["country"] == "DE"
    assert "confidence=0.91" in record["notes"]

    assert captured_params[0]["query"] == "Example GmbH"
    assert captured_params[0]["postalCode"] == "80333"
    assert captured_params[0]["country"] == "DE"
