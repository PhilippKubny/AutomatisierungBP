"""Simple smoke test for NorthData provider."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

from bpauto.providers import NorthDataProvider


def main() -> int:
    dotenv_path = Path(".env")
    if dotenv_path.exists():
        load_dotenv(dotenv_path=dotenv_path)

    try:
        provider = NorthDataProvider(download_ad=False)
    except RuntimeError as exc:
        print(f"Konfiguration unvollständig: {exc}", file=sys.stderr)
        return 2

    try:
        record = provider.fetch(name="Siemens AG", zip_code="80333", country="DE")
    except RuntimeError as exc:
        print(f"Authentifizierungsfehler: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # pragma: no cover - Netzwerkfehler o.ä.
        print(f"NorthData-Anfrage fehlgeschlagen: {exc}", file=sys.stderr)
        return 2

    def _default(obj: Any) -> Any:
        if isinstance(obj, set):
            return sorted(obj)
        raise TypeError(f"Nicht serialisierbar: {obj!r}")

    print(json.dumps(record, indent=2, sort_keys=True, ensure_ascii=False, default=_default))

    notes = record.get("notes")
    if isinstance(notes, str) and "no result" in notes.lower():
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
