"""Simple smoke test for NorthData provider."""

from __future__ import annotations

import json
import sys
from pathlib import Path

from dotenv import load_dotenv

from bpauto.providers import NorthDataProvider


def main() -> int:
    dotenv_path = Path(".env")
    if dotenv_path.exists():
        load_dotenv(dotenv_path=dotenv_path)

    provider = NorthDataProvider(download_ad=False)
    record = provider.fetch(name="Siemens AG", zip_code="80333", country="DE")

    print(json.dumps(record, indent=2, sort_keys=True, ensure_ascii=False))

    notes = record.get("notes")
    if isinstance(notes, str) and "no result" in notes.lower():
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
