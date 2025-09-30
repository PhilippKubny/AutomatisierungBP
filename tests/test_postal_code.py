import pytest

from PlayHandelsregister import build_parser, looks_like_same_zip


def parse_args(argv):
    parser = build_parser()
    return parser.parse_args(argv)


def test_cli_postal_code_flags_batch_mode():
    args = parse_args([
        "--postal-code",
        "--postal-code-col",
        "AB",
        "--excel",
        "TestBP.xlsx",
    ])
    assert args.postal_code is True
    assert args.postal_code_col == "AB"
    assert args.plz == ""


def test_cli_postal_code_flags_single_shot():
    args = parse_args([
        "--postal-code",
        "--plz",
        "45128",
        "-s",
        "Firma",
    ])
    assert args.postal_code is True
    assert args.plz == "45128"


@pytest.mark.parametrize(
    "text, plz, expected",
    [
        ("Musterstraße 1, 45128 Essen", "45128", True),
        ("Musterstraße 1, 45129 Essen", "45128", False),
        ("", "45128", False),
        ("Ohne Postleitzahl", "", False),
    ],
)
def test_looks_like_same_zip(text, plz, expected):
    assert looks_like_same_zip(text, plz) is expected
