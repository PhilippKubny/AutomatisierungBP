from __future__ import annotations

import sys
from argparse import Namespace

from bpauto import cli


def test_main_skips_rows_with_unsupported_country(monkeypatch):
    rows = [
        {
            "index": 3,
            "name": "Allowed GmbH",
            "zip": "10115",
            "city": "Berlin",
            "country": "DE",
            "street": None,
            "house_number": None,
        },
        {
            "index": 4,
            "name": "Blocked GmbH",
            "zip": "75008",
            "city": "Paris",
            "country": "US",
            "street": None,
            "house_number": None,
        },
    ]

    def fake_iter_rows(**_: object):
        for row in rows:
            yield row

    fetch_calls: list[dict[str, object]] = []

    class DummyProvider:
        def __init__(self, download_ad: bool):
            self.download_ad = download_ad

        def fetch(self, **kwargs: object) -> dict[str, object]:
            fetch_calls.append(kwargs)
            return {"notes": ""}

    args = Namespace(
        excel="dummy.xlsx",
        sheet="Sheet1",
        start=3,
        end=None,
        name_col="C",
        zip_col=None,
        city_col=None,
        country_col="J",
        street_col=None,
        house_number_col=None,
        mapping_yaml=None,
        source="api",
        download_ad=False,
        verbose=False,
        dry_run=True,
    )

    monkeypatch.setattr(cli, "_parse_args", lambda: args)
    monkeypatch.setattr(cli.excel_io, "iter_rows", fake_iter_rows)
    monkeypatch.setattr(cli, "NorthDataProvider", DummyProvider)

    exit_code = cli.main()

    assert exit_code == 0
    assert len(fetch_calls) == 1
    assert fetch_calls[0]["country"] == "DE"


def test_parse_args_sets_default_country_column(monkeypatch):
    argv = [
        "bpauto",
        "--excel",
        "dummy.xlsx",
        "--sheet",
        "Sheet1",
    ]
    monkeypatch.setattr(sys, "argv", argv)

    args = cli._parse_args()

    assert args.country_col == "J"
