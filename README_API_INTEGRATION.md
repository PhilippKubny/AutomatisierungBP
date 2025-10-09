# NorthData API Integration

Diese Dokumentation beschreibt den Excel-gestützten Abruf von Unternehmensdaten über die NorthData-API.

## Architekturüberblick

Die Automatisierung besteht aus drei Bausteinen:

1. **Excel-I/O (`bpauto.excel_io`)** – liest Eingabedaten aus einer Arbeitsmappe, schreibt Ergebnisse in definierte Spalten und übernimmt das Persistieren.
2. **Provider-Schnittstelle (`bpauto.providers.base.Provider`)** – definiert das Abstraktionslayer für Datenquellen und das gemeinsame `CompanyRecord`-Format.
3. **NorthData-Provider (`bpauto.providers.northdata.NorthDataProvider`)** – implementiert die konkrete API-Anbindung inkl. Retry-Logik, Fehlerbehandlung und optionalem Download amtlicher Registerauszüge.

Der CLI-Einstieg (`cli.py`) orchestriert den Ablauf: Zeilen iterieren, Provider ansprechen, Ergebnisse zurückschreiben und am Ende die Arbeitsmappe speichern.

## Setup & Installation

1. **Virtuelle Umgebung anlegen und aktivieren**
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # PowerShell: .venv\Scripts\Activate.ps1
   ```

2. **Projekt (inkl. Dev-Tools) installieren**
   ```bash
   pip install -e .[dev]
   ```

3. **Environment-Variablen setzen**
   Der NorthData API Key ist Pflicht. Optional kann eine `.env`-Datei genutzt werden.
   *Bash / zsh*
   ```bash
   export NORTHDATA_API_KEY="dein_key"
   ```
   *PowerShell*
   ```powershell
   $Env:NORTHDATA_API_KEY = "dein_key"
   ```

4. **Konfiguration prüfen**
   ```bash
   python -m tests.smoke_fetch
   ```
   Der Befehl gibt ein JSON mit dem gefundenen Datensatz aus. Falls kein Treffer oder keine Verbindung möglich ist, beendet sich das Skript mit Exitcode `1` (kein Ergebnis) bzw. `2` (Konfigurations-/Netzwerkproblem).

## Excel-Lauf starten

```bash
python cli.py \
  --excel "Liste BP Cleaning Kreditoren.xlsx" \
  --sheet "Tabelle1" \
  --start 3 \
  --name-col C \
  --street-col F \
  --house-number-col G \
  --city-col H \
  --zip-col I \
  --country-col J \
  --mapping-yaml mappings/example_mapping.yaml \
  --download-ad
```

* `--end` kann optional angegeben werden; ohne Angabe wird bis zur letzten Zeile mit Firmennamen gelesen.
* `--dry-run` führt den Abruf aus, schreibt aber nichts in Excel (praktisch für Tests).
* `--street-col` und `--house-number-col` werden eingelesen und protokolliert. Für den NorthData-Aufruf wird jedoch ausschließlich
  der Ort (`--city-col`) als `address`-Parameter kombiniert mit dem Firmennamen (`name`) verwendet. Platzhalter wie `#` oder `-`
  werden automatisch ignoriert.

Der CLI-Lauf fasst am Ende die Anzahl verarbeiteter Zeilen, Treffer sowie `no result`-Fälle zusammen. Exitcodes:

| Code | Bedeutung |
| ---- | --------- |
| 0    | Erfolgreiche Verarbeitung |
| 2    | Falsche Parameter (z. B. ungültige Spalte) oder unbekannte Quelle |
| 3    | Authentifizierungsfehler gegenüber der NorthData-API |
| 4    | Allgemeine Laufzeitfehler (z. B. Provider-Ausnahme) |

## Mapping-Dateien

Die Datei `mappings/example_mapping.yaml` weist die Felder des `CompanyRecord` konkreten Excel-Spalten (inkl. Mehrbuchstaben-Spalten wie `AA`, `AB`, `AC`) zu. Eigene Mappings können erstellt werden; nicht gesetzte oder leere Werte werden ignoriert. Pflichtfelder sind mindestens `legal_name` und `notes`, damit Ergebnisse nachvollziehbar bleiben.

## Tests & Qualitätssicherung

* `make lint` – führt `ruff` aus
* `make type` – führt `mypy` aus
* `make test` – führt die Pytest-Unit-Tests aus (inkl. Excel-I/O-Test)
* `make smoke` – ruft den Smoke-Test auf (`Exit 2`, wenn kein API-Key vorliegt)

## Troubleshooting

| Problem | Lösung |
| ------- | ------ |
| **401/403** – Fehlermeldung „API Key ungültig oder fehlend“ | API-Schlüssel prüfen, ggf. neue `.env` einlesen oder Environment neu setzen. |
| **429** – Too Many Requests | Kurz warten; der Provider wiederholt Anfragen automatisch bis zu fünfmal mit exponential Backoff. |
| **Fehlende Spalten** | Mapping-Datei und CLI-Parameter prüfen. Spaltenwerte müssen reine Buchstaben (z. B. `AA`) sein. |
| **Windows-Pfade / Umlautprobleme** | Pfade in Anführungszeichen setzen und PowerShell-UTF-8 sicherstellen (`chcp 65001`). |
| **Keine Treffer** | `notes`-Spalte enthält „no result“. Mögliche Ursachen: falscher Firmenname, andere Rechtsform oder fehlende PLZ. |

## Sicherheit

* API-Schlüssel niemals commiten.
* `.env` ist in `.gitignore` hinterlegt.
* Downloads amtlicher Auszüge landen unter `./.artifacts/extracts` und sollten regelmäßig bereinigt werden.

## Weiterführende Hinweise

* Der Provider setzt auf HTTP-Timeouts (`connect=5s`, `read=30s`) und wiederholt fehlerhafte Anfragen bei 429/5xx.
* Alle Logs sind auf Deutsch gehalten, um Betriebsauswertungen zu erleichtern.
* Die CLI speichert die Arbeitsmappe erst nach Abschluss aller Zeilen – bei Abbrüchen bleiben ursprüngliche Daten erhalten.
