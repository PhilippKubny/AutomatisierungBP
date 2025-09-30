# Automatisierung Handelsregister – README

## Inhaltsverzeichnis
1. [Projektüberblick](#projektüberblick)
2. [Funktionen im Überblick](#funktionen-im-überblick)
3. [Projektstruktur & Skripte](#projektstruktur--skripte)
4. [Systemvoraussetzungen](#systemvoraussetzungen)
5. [Installation & Setup](#installation--setup)
   1. [Repository beziehen](#repository-beziehen)
   2. [Python-Umgebung vorbereiten](#python-umgebung-vorbereiten)
   3. [Playwright-Browser installieren](#playwright-browser-installieren)
6. [Konfiguration](#konfiguration)
   1. [Excel-Layout](#excel-layout)
   2. [Download-Verzeichnisse](#download-verzeichnisse)
7. [Ausführung](#ausführung)
   1. [Batch-Modus mit Excel](#batch-modus-mit-excel)
   2. [Single-Shot-Suche](#single-shot-suche)
   3. [Wichtige CLI-Optionen](#wichtige-cli-optionen)
   4. [PDF-Nachbearbeitung](#pdf-nachbearbeitung)
8. [Troubleshooting & Hinweise](#troubleshooting--hinweise)
9. [Tests & Qualitätssicherung](#tests--qualitätssicherung)
10. [Lizenz & Haftungsausschluss](#lizenz--haftungsausschluss)

---

## Projektüberblick
Dieses Projekt automatisiert Rechercheaufgaben im deutschen Handelsregister. Mittels [Playwright](https://playwright.dev/) werden die Eingabemasken der "Erweiterten Suche" bedient, Treffer ausgewertet und – sofern gewünscht – aktuelle Auszüge (AD-PDFs) heruntergeladen. Die gewonnenen Daten werden anschließend analysiert und in eine Excel-Arbeitsmappe zurückgeschrieben.

Der Ablauf gliedert sich in drei Hauptschritte:
1. Auslesen der Eingabedaten (Excel oder Einzeleingabe)
2. Automatisierte Web-Interaktion mit handelsregister.de
3. Parsing heruntergeladener PDFs und Rückschreiben der Ergebnisse

## Funktionen im Überblick
- Headless-Browser-Automation via Playwright
- Zwei Betriebsmodi: Batch-Verarbeitung aus Excel sowie Single-Shot-Suche
- Optionale PDF-Downloads mit benutzerdefiniertem Speicherort
- PDF-Parsing (Registerart, Registernummer, Adresse) mittels `pdfplumber`
- Rückschreiben der Ergebnisse in definierte Excel-Spalten inklusive Änderungskennzeichnung
- Robuste Wiederholungslogik bei Netz-/UI-Fehlern sowie Debug-Ausgaben

## Projektstruktur & Skripte
- **`PlayHandelsregister.py`** – Hauptskript für die automatisierte Recherche inklusive Excel-Anbindung und PDF-Download.
- **`PDFScanner.py`** – CLI-Tool zum nachträglichen Auslesen heruntergeladener PDF-Auszüge (Registerart, Nummer, Anschrift usw.).
- **`PDFdump.py`** – Hilfsprogramm zur strukturierten Analyse problematischer PDFs (Text- und Byte-Dumps für Debugging).
- **`tests/`** – Enthält `pytest`-basierte Tests für Kernfunktionen.

Ergänzende Helferskripte wie `exel.py` sind optional und richten sich an fortgeschrittene Nutzer:innen, die den Datenexport bzw. das Troubleshooting automatisieren möchten.

## Systemvoraussetzungen
- **Betriebssystem:** Windows, macOS oder Linux (Playwright wird unter allen großen Plattformen unterstützt)
- **Python:** Version 3.10 oder höher (Projekt wurde mit 3.11 getestet)
- **Browser-Abhängigkeiten:** Chromium/Firefox/WebKit via Playwright
- **Microsoft Excel oder kompatible Software** zum Verwalten der Eingabedateien

## Installation & Setup

### Repository beziehen
```bash
# via HTTPS
git clone https://github.com/<Ihr-Account>/AutomatisierungBP.git
cd AutomatisierungBP
```

### Python-Umgebung vorbereiten
Es wird dringend empfohlen, eine virtuelle Umgebung zu verwenden:
```bash
python -m venv .venv
source .venv/bin/activate   # PowerShell: .venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r requirements.txt
```

### Playwright-Browser installieren
Playwright bringt eigene Browser-Builds mit. Installieren Sie diese nach der Paketinstallation einmalig:
```bash
playwright install chromium
# Unter Linux können zusätzlich Systemabhängigkeiten notwendig sein:
# playwright install-deps chromium
```

> Hinweis: Für grafische Ausführung (Debugging) muss ggf. ein Display-Server vorhanden sein.

## Konfiguration

### Excel-Layout
Für die Batch-Verarbeitung erwartet `PlayHandelsregister.py` bestimmte Spalten. Die Defaults lassen sich bei Bedarf per CLI-Option anpassen.

| Zweck                         | Default-Spalte | Option                           |
|------------------------------|----------------|----------------------------------|
| Name 1 (Suchbegriff)         | `C`            | `--name-col` (Input) / `--name1-col` (Update)
| Registernummer               | `U`            | `--regno-col`
| SAP-Lieferantennummer        | `A`            | `--sap-supplier-col`
| SAP-Kundennummer             | `B`            | `--sap-customer-col`
| Land                         | `J`            | `--country-col`
| Straße / Hausnummer          | `X` / `Y`      | `--street-col`, `--house-number-col`
| Postleitzahl / Ort           | `AA` / `Z`     | `--postal-code-col`, `--city-col`
| Dokumentpfad                 | `P`            | `--doc-path-col`
| Änderungskennzeichen         | `Q`            | `--changes-check-col`
| Datum letzte Prüfung         | `S`            | `--date-check-col`
| Registerart                  | `V`            | `--register-type-col`

Weitere Spalten (z. B. `Name2`, `Name3`) können über die entsprechenden Optionen gesetzt werden. Der Start von Datenzeilen wird standardmäßig ab Zeile 3 erwartet (Header in Zeile 1–2).

### Download-Verzeichnisse
- Standardmäßig werden PDFs in `~/Downloads/BP` abgelegt. Dies lässt sich mit `--outdir` ändern.
- Für Protokolleinträge wird (falls nötig) `~/Downloads/HumanCheck.txt` verwendet.

## Ausführung

### Batch-Modus mit Excel
Verarbeiten Sie mehrere Einträge aus einer Excel-Tabelle:
```bash
python PlayHandelsregister.py \
    --excel "~/Downloads/TestBP.xlsx" \
    --sheet "Tabelle1" \
    --start 3 \
    --end 30 \
    --download-ad \
    --postal-code \
    --postal-code-col AA
```
Wichtige Hinweise:
- `--excel` aktiviert den Batch-Modus und ist Pflichtparameter.
- `--start` und `--end` referenzieren 1-basierte Zeilennummern (inklusive).
- `--download-ad` löst den PDF-Download aus. Ohne diese Option werden nur Treffer ausgewertet.
- `--postal-code` verwendet die Postleitzahl aus der Excel-Tabelle zur Suche; mit `--postal-code-col` kann die Spalte angepasst werden.

### Single-Shot-Suche
Für Einzelfälle ohne Excel-Datei:
```bash
python PlayHandelsregister.py \
    -s "THYSSENKRUPP SCHULTE GMBH" \
    --download-ad \
    --postal-code \
    --plz 45128
```
- `-s/--schlagwoerter` ist der Pflicht-Suchbegriff.
- `--plz` liefert die Postleitzahl für Single-Shot-Suchen, sobald `--postal-code` gesetzt ist.
- `--register-number`, `--sap-number` und `--row-number` bleiben optional, steigern aber die Treffer- und Rückschreibqualität.

### Wichtige CLI-Optionen
- `-d/--debug`: Ausführliche Konsolenausgabe (inkl. HTML-Snippets bei Fehlern)
- `--headful`: Startet den Browser sichtbar (Standard: headless)
- `--download-ad`: Aktiviert PDF-Downloads
- `--outdir`: Zielverzeichnis für PDF-Dateien
- `--schlagwortOptionen {all|min|exact}`: Steuerung der Volltextsuche
- `--postal-code`: Aktiviert die Postleitzahl-Filterung; kombiniert mit `--postal-code-col` (Batch) bzw. `--plz` (Single-Shot)

Eine vollständige Übersicht erhalten Sie mit `python PlayHandelsregister.py --help`.

### PDF-Nachbearbeitung
Bereits heruntergeladene AD-PDFs lassen sich mit den beigefügten Utilities effizient weiterverarbeiten:

```bash
# Kerndaten mehrerer PDFs extrahieren und als CSV speichern
python PDFScanner.py --in "~/Downloads/BP" --out "~/Downloads/handelsregister-daten.csv"

# Strukturdump für die Fehlersuche erzeugen (gibt Ausgaben nach stdout aus)
python PDFdump.py --in "~/Downloads/BP"
```

`PDFScanner.py` liefert pro Datei Registertyp, Nummer, Firma sowie aufbereitete Adressbestandteile. `PDFdump.py` erstellt ausführliche Text- und Strukturdumps, um Layout-Probleme oder OCR-Scans zu identifizieren. Die Outputs lassen sich optional in der Datei `~/Downloads/HumanCheck.txt` ablegen.

## Troubleshooting & Hinweise
- **Mehrere Treffer:** Die Automatisierung verarbeitet nur eindeutige Treffer. Mehrfachtreffer werden im Excel-Protokoll vermerkt.
- **Timeouts / UI-Änderungen:** Das Skript versucht bis zu drei Wiederholungen. Bestehen Probleme weiter, prüfen Sie manuell die Website-Struktur.
- **PDF-Parsing-Fehler:** Unbekannte Dokumentlayouts werden als "Unexpected format" protokolliert.
- **Downloads unter Windows:** Stellen Sie sicher, dass der Pfad Schreibrechte besitzt und nicht durch Sicherheitssoftware blockiert wird.
- **Headful-Debugging:** Verwenden Sie `--headful`, um die Browserinteraktion nachzuvollziehen.

## Tests & Qualitätssicherung
- Unit-Tests werden mit `pytest` bereitgestellt. Führen Sie sie aus, um lokale Änderungen zu verifizieren:
  ```bash
  pytest
  ```
- Stellen Sie sicher, dass alle PDF- und Excel-Pfade auf Ihrem System existieren, bevor Sie produktiv testen.

## Lizenz & Haftungsausschluss
Die Nutzung der Handelsregister-Daten unterliegt den Bedingungen von handelsregister.de. Verwenden Sie dieses Skript verantwortungsbewusst und beachten Sie geltende Nutzungsrichtlinien, Datenschutz- sowie Compliance-Vorgaben Ihres Unternehmens.

