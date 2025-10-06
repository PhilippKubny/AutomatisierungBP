# AutomatisierungBP

AutomatisierungBP bündelt Skripte und Integrationen, um Handelsregister-Recherchen, PDF-Extraktionen und Excel-Updates für Business-Partner-Prüfungen zu automatisieren. Die Sammlung kombiniert einen Playwright-Scraper, einen NorthData-API-Provider sowie robuste Excel-/PDF-Helfer in einer installierbaren Python-Distribution.

## Highlights

- Einheitliches CLI `bpauto` für Batch-Updates aus Excel-Dateien.
- Wiederverwendbare Provider-Schnittstelle zur Anbindung externer Datenquellen (z. B. NorthData).
- Werkzeuge für PDF-Parsing und Excel-Schreiboperationen, inklusive Fehlerhandling.
- Konsistente Logging-Ausgabe via `bpauto.utils.logging_setup`.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -e .[dev]
```

Optionale Tools wie Playwright-Browser oder zusätzliche Provider-Abhängigkeiten müssen separat installiert werden.

## Konfiguration

Kopiere `.env.example` nach `.env` und hinterlege den NorthData-API-Schlüssel:

```dotenv
NORTHDATA_API_KEY=dein_api_key
```

## Verwendung

```bash
bpauto --excel "Liste.xlsx" --sheet "Tabelle1" --source api
```

Wichtige Optionen:

- `--start` / `--end`: Zeilenbereich (1-basiert) begrenzen.
- `--mapping-yaml`: eigenes Mapping zwischen `CompanyRecord`-Feldern und Excel-Spalten.
- `--download-ad`: amtliche Ausdrucke (AD) als PDF speichern.
- `--dry-run`: Daten nur abrufen, keine Excel-Schreibvorgänge.
- `--verbose`: detailliertere Log-Ausgabe.

## Repository Structure & Responsibilities

```
AutomatisierungBP/
├── pyproject.toml            # Projekt- und Build-Metadaten
├── README.md                 # Dieses Dokument
├── .env.example              # Vorlage für sensible Einstellungen
├── .gitignore                # Ignore-Regeln (Virtualenvs, Artefakte, usw.)
├── mappings/                 # Beispielhafte Mapping-Dateien für Excel
├── src/
│   └── bpauto/
│       ├── __init__.py       # Paket-Exports
│       ├── cli.py            # CLI-Einstiegspunkt
│       ├── excel_io.py       # Lesen/Schreiben von Excel-Tabellen
│       ├── pdf_scanner.py    # PDF-Auswertung
│       ├── handelsregister.py# Playwright-Scraper für handelsregister.de
│       ├── providers/
│       │   ├── __init__.py   # Provider-Exports
│       │   ├── base.py       # Gemeinsame Typen/Protokolle
│       │   └── northdata.py  # NorthData-API-Integration
│       └── utils/
│           └── logging_setup.py  # Zentrales Logger-Setup
├── tests/
│   ├── test_excel_io.py      # Unit-Tests für Excel-Helfer
│   ├── test_provider_northdata.py # Tests für den API-Provider
│   ├── test_handelsregister.py    # Parser-/Scraper-Regressionsfälle
│   └── smoke_fetch.py        # Einfache Smoke-CLI für NorthData
├── vendor/
│   └── bp_api/               # Eingebettete Legacy-Komponenten
└── Makefile                  # Hilfstasks für Tests/Linting
```

## How to add new Providers (e.g. Orbis, NorthData, Scraper)

1. Lege ein neues Modul unter `src/bpauto/providers/` an und implementiere das `Provider`-Protokoll aus `base.py`.
2. Nutze das Logging über `setup_logger()` für nachvollziehbare Ausgaben und Fehlerbehandlung.
3. Registriere den Provider in `src/bpauto/providers/__init__.py`, damit er per `bpauto.providers` importierbar wird.
4. Ergänze falls nötig neue Mappings oder Konfigurationsparameter im CLI (`src/bpauto/cli.py`).
5. Schreibe Tests in `tests/`, idealerweise mit Mocking für API-Aufrufe, um Netzwerkabhängigkeiten zu vermeiden.
6. Dokumentiere Besonderheiten oder zusätzliche Umgebungsvariablen im README.

## Entwicklung & Tests

```bash
make format       # black/ruff Formatierung
make lint         # Ruff Lints
make test         # pytest + optionale Smoke-Tests
```

Das CLI kann lokal per Module-Run getestet werden:

```bash
python -m bpauto.cli --help
```

## Lizenz & Hinweise

- Verwende `.artifacts/` für temporäre Downloads (wird automatisch ignoriert).
- Sensible Schlüssel niemals ins Repo commiten – `.env` ist in `.gitignore` eingetragen.
