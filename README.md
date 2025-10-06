# Handelsregister Automation Toolkit

AutomatisierungBP automatisiert die Recherche im deutschen Handelsregister und verarbeitet die Ergebnisse für interne Business-Partner-Prüfungen. Das Tool füllt Web-Formulare per Playwright, lädt amtliche Ausdrucke (AD-PDFs) herunter, extrahiert deren Inhalte und schreibt die Daten zurück in Excel-Tabellen.

## Inhaltsverzeichnis

- [Überblick](#überblick)
- [Hauptfunktionen](#hauptfunktionen)
- [Projektaufbau](#projektaufbau)
- [Voraussetzungen](#voraussetzungen)
- [Installation](#installation)
- [Konfiguration](#konfiguration)
- [Anwendung](#anwendung)
  - [Stapelverarbeitung mit Excel](#stapelverarbeitung-mit-excel)
  - [Einzelabfragen](#einzelabfragen)
  - [Gemeinsame Argumente](#gemeinsame-argumente)
- [Ausgabe und Ergebnisdateien](#ausgabe-und-ergebnisdateien)
- [Troubleshooting](#troubleshooting)
- [Weiterentwicklung](#weiterentwicklung)

## Überblick

Das Skript `PlayHandelsregister.py` bildet die manuelle Recherche auf [handelsregister.de](https://www.handelsregister.de/rp_web/welcome.xhtml) nach. Es navigiert zur erweiterten Suche, setzt die übergebenen Parameter, lädt bei eindeutigen Treffern die AD-PDFs herunter und aktualisiert definierte Zellen einer Excel-Arbeitsmappe. Unterstützende Module kümmern sich um das Lesen/Schreiben der Excel-Datei (`exel.py`) sowie die Texterkennung aus AD-PDFs (`PDFScanner.py`).

## Hauptfunktionen

- Automatisiertes Ausfüllen der erweiterten Suche im Handelsregister.
- Verarbeitung von Excel-Tabellen zur Stapelprüfung mehrerer Unternehmen.
- Optionaler Download der amtlichen Ausdrucke (AD) pro Treffer.
- Auslesen relevanter Stammdaten (Firma, Registernummer, Adresse) aus AD-PDFs.
- Rückschreiben der ermittelten Informationen in frei konfigurierbare Excel-Spalten.
- Debug- und Wiederholungslogik zur Minimierung von Ausfällen durch UI-Änderungen oder Rate Limits.

## Projektaufbau

| Datei | Zweck |
| --- | --- |
| `PlayHandelsregister.py` | Zentrales CLI-Skript, steuert Playwright-Browser, Download und Excel-Update. |
| `exel.py` | Hilfsfunktionen zum Lesen und Aktualisieren der Excel-Datei. |
| `PDFScanner.py` | Erkennung relevanter Textpassagen in AD-PDFs und Strukturierung der Daten. |
| `PDFdump.py` | Debugging-Werkzeug zum Analysieren neuer/abweichender PDF-Strukturen. |
| `test_handelsregister.py` | Ausgangspunkt für automatisierte Tests (z. B. Parser/Scanner). |

## Voraussetzungen

- Python 3.11 oder neuer.
- Google Chrome oder Microsoft Edge ist nicht erforderlich; Playwright bringt eigene Browser-Bundles mit.
- Schreib-/Leserechte auf das Excel- und Download-Verzeichnis.

## Installation

1. Repository klonen oder als ZIP herunterladen.
2. (Empfohlen) Virtuelle Umgebung anlegen.
3. Abhängigkeiten installieren:

   ```bash
   pip install -r requirements.txt
   playwright install
   ```

   Der zweite Befehl installiert die benötigten Browser-Binaries für Playwright.

## Konfiguration

### Excel-Vorbereitung

- Das Tool erwartet eine Excel-Datei mit Kopfzeilen in den ersten beiden Zeilen. Standardmäßig beginnt die Verarbeitung ab Zeile 3.
- Relevante Spalten können über CLI-Argumente konfiguriert werden. Standardwerte:

  | Zweck | Standardspalte |
  | --- | --- |
  | SAP-Lieferantennummer | `A` |
  | SAP-Kundennummer | `B` |
  | Firmenname (Eingabe) | `C` |
  | Land | `J` |
  | Postleitzahl (Eingabe) | `I` |
  | Registertyp (Ausgabe) | `V` |
  | Registernummer (Ausgabe) | `U` |
  | Dokumentpfad | `P` |
  | Änderungsbedarf | `Q` |
  | Datum letzte Prüfung | `S` |
  | Ergebniszählung/Log (`name1-4`) | `T` |
  | Adressdaten (Straße, Hausnummer, Stadt, PLZ) | `X`, `Y`, `Z`, `AA` |

- Eigene Spaltenzuordnungen lassen sich über die entsprechenden `--*-col` Argumente anpassen.

### Download-Verzeichnis

- Ohne Angabe wird im Benutzerverzeichnis `~/Downloads/BP` angelegt.
- Alternativ kann mit `--outdir` ein eigener Pfad gesetzt werden.

## Anwendung

Das CLI unterscheidet zwei Betriebsmodi: Stapelverarbeitung mit Excel (`--excel`) und Einzelabfragen. In beiden Fällen empfiehlt es sich, bei UI-Problemen den Debug-Modus zu aktivieren (`--debug`).

### Stapelverarbeitung mit Excel

```bash
python PlayHandelsregister.py \
  --excel "C:\\Users\\User\\Downloads\\TestBP.xlsx" \
  --sheet "Tabelle1" \
  --start 25 \
  --end 30 \
  --download-ad \
  --postal \
  --outdir "C:\\Users\\User\\Downloads\\BP" \
  --debug
```

- `--start` und `--end` definieren den verarbeiteten Zeilenbereich (inklusiv).
- `--postal` ergänzt die Suche um Postleitzahlen (falls in der Excel hinterlegt).
- `--download-ad` speichert AD-PDFs für eindeutige Treffer.
- Fehler wie Mehrfachtreffer oder nicht gefundene Firmen werden in der konfigurierten Ergebnis-Spalte protokolliert.

### Einzelabfragen

```bash
python PlayHandelsregister.py \
  --schlagwoerter "THYSSENKRUPP SCHULTE GMBH" \
  --register-number "26718" \
  --sap-number "2203241" \
  --row-number "352" \
  --download-ad \
  --postal \
  --outdir "C:\\Users\\User\\Downloads\\BP"
```

- `--schlagwoerter` ist Pflicht (Firmenname bzw. Stichworte).
- `--sap-number` und `--row-number` dienen zur Zuordnung in der Excel-Datei und sind im Einzelmodus erforderlich.
- `--register-number` erhöht die Treffergenauigkeit, ist aber optional.

### Gemeinsame Argumente

| Argument | Beschreibung |
| --- | --- |
| `-d`, `--debug` | Zusätzliche Konsolenausgaben für Fehlersuche. |
| `--headful` | Öffnet den Browser sichtbar (Standard: headless). |
| `--schlagwortOptionen {all,min,exact}` | Steuerung der Suchlogik (alle Stichworte, mindestens eines, exakte Übereinstimmung). |
| `--postal` | Aktiviert Postleitzahl-Suche (setzt voraus, dass ein Wert vorliegt). |
| `--download-ad` | Lädt AD-PDFs in das Ausgabeverzeichnis. |
| `--outdir PATH` | Zielordner für Downloads und Prüfdateien. |
| `--start/--end` | Zeilenbereich bei Excel-Stapelverarbeitung. |

Eine vollständige Auflistung aller Argumente liefert `python PlayHandelsregister.py --help`.

## Ausgabe und Ergebnisdateien

- **AD-PDFs** werden als `<SAP- oder interner Schlüssel>_<Firmenname>_<YYYY-MM-DD>.pdf` im gewählten Ausgabeverzeichnis gespeichert.
- **Excel-Updates**: Bei eindeutigen Treffern aktualisiert das Skript Name, Adresse, Registernummer, Registertyp, Downloadpfad sowie Status- und Datumsfelder.
- **Protokollierung**: Mehrfachtreffer oder Fehlermeldungen werden in der Ergebnis-Spalte (Standard `T`) mitgezählt bzw. beschrieben (`Unexpected format`, `0` für keine Treffer etc.).

## Troubleshooting

- Bei UI-Änderungen der Webseite hilft `--debug`, um HTML-Auszüge zu erhalten.
- Tritt ein Timeout auf, versucht das Skript bis zu drei Wiederholungen. Danach empfiehlt sich eine Pause, um mögliche Rate Limits zu umgehen.
- Nicht erkannte PDF-Layouts können mit `PDFdump.py` analysiert und anschließend in `PDFScanner.py` ergänzt werden.

## Weiterentwicklung

- Tests für PDF-Parsing und Excel-Schreiboperationen erweitern (`test_handelsregister.py`).
- Unterstützung weiterer Dokumenttypen (neben AD) evaluieren.
- Resilienz gegen UI-Änderungen durch robustere Selektoren und visuelle Prüfungen steigern.

