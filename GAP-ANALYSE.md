# Gap-Analyse: Excel (xlsm) vs. SPFx WebPart

> Stand: 02.04.2026  
> Vergleich der Features aus `Z_Übersicht Technische Anfragen.xlsm` mit dem aktuellen SPFx-WebPart `ta-management`.

---

## Vollständig abgedeckt ✅

| Feature (Excel) | SPFx-Umsetzung |
|---|---|
| **TA-Erfassung** mit Formular | `NeueTa.tsx` — Formular mit Kunde, Anwendung, Material, Kategorie, Priorität, Budget, Wunschtermin etc. |
| **Projekt-Auto-Fill** (Kunde → Anwendungen/Budget/Material) | CSV-Parser in `SharePointService.getProjekte()` liest `Projektliste.csv`, Dropdowns werden dynamisch befüllt wie das VBA-Hilfsblatt |
| **Statusautomatik** (5 Status basierend auf geplantem Termin) | `evaluateStatuses()` — gleiche Logik: kein Termin → `Termin planen`, vergangen → `überfällig`, nah dran → `prüfen`, Zukunft → `läuft planmäßig`, erledigt → `abgeschlossen` |
| **TA-Nummernvergabe** (automatisch) | `getNextTaNumber()` — Format `TA-YYYY-NNN` (verbessert gegenüber dem alten `JJMMNN`-Format) |
| **Termineingabe** (Datum + Verantwortlicher) | `TaDetail.tsx` — Datumsfeld + People-Picker für Verantwortlichen |
| **Terminverschiebung** mit Grund | `Modal.tsx` — 13 Verschiebungsgründe (erweitert von 6 im Excel), alter Termin wird in `field_22` gespeichert |
| **Verzögerungsgrund** bei verspäteter Erstplanung | `TaDetail.tsx` — Delay-Modal mit 6 Gründen (`INITIAL_DELAY_REASONS`), konfigurierbar über `DelayThresholdDays` |
| **KPI-Dashboard** (Überfällig/Planmäßig/Prüfen/Planen) | `Dashboard.tsx` — 4 KPI-Tiles + letzte Aktivitäten |
| **Alle TAs filtern/suchen** | `AlleTas.tsx` — Textsuche + Status-Filter + Prioritäts-Sortierung |
| **Budgetkontrolle** (Budget/IST/Plan/Differenz) | `TaDetail.tsx` — zeigt `field_17`–`field_20` an (Budget, IST-Kosten, geplante Kosten, Differenz) |
| **Kategorien-Verwaltung** | `Settings.tsx` — CRUD auf `TA_Kategorien`-Liste (dynamisch statt hardcoded) |
| **App-Konfiguration** (Delay Threshold) | `Settings.tsx` — `TA_Config`-Liste, key-value |
| **Mehrbenutzerfähigkeit** | Inhärent gelöst — SharePoint-Listen statt Dateisperre |
| **Benutzer-Zuordnung** (Ersteller) | Automatisch via `sp.web.currentUser` (besser als `Environ("UserName")`) |
| **Verantwortlicher-Suche** | People-Picker via PnPjs `clientPeoplePickerSearchUser` (statt hardcodierter Mitarbeiterliste) |

---

## Teilweise abgedeckt / anders gelöst ⚠️

### 1. Verzögerungsgründe (Verschiebung)
- **Excel**: 6 Gründe (Rückfragen, Selbstorganisation, Komplexität, Kapazität, Kein Budget, Technische Probleme)
- **SPFx**: 13 Gründe (`VERSCHIEBUNG_GRUENDE`) — bewusste fachliche Erweiterung
- **Status**: OK — alle alten 6 sind enthalten

### 2. "prüfen"-Schwelle
- **Excel**: `WORKDAY(TODAY(),3)` → **3 Werktage** Fenster
- **SPFx**: `planDate.getTime() === today.getTime()` → nur **gleicher Tag**
- **Status**: ⚠️ **Bug** — Die 3-Werktage-Logik fehlt!

### 3. Kosten bearbeiten
- **Excel**: Direkte Eingabe in Spalten W/X/Y
- **SPFx**: `TaDetail.tsx` zeigt Kosten nur **zur Anzeige** an, keine Eingabefelder
- **Status**: ⚠️ Eingabefelder fehlen

---

## Fehlt komplett ❌

### 1. E-Mail-Benachrichtigungen
- **Excel**: Automatische Outlook-Mails bei:
  - TA-Erstellung → Mail an `awt@barlog.de` + CC `ablage@barlog.de`
  - Statuswechsel `läuft planmäßig → prüfen` → Erinnerungs-Mail an Verantwortlichen
  - Statuswechsel `prüfen → überfällig` → Eskalations-Mail
- **SPFx**: Keine E-Mails
- **Empfohlene Lösung**: Power Automate Flow auf Listenänderungen oder Graph API
- **Aufwand**: Mittel

### 2. Ordner-Erstellung auf Netzlaufwerk
- **Excel**: Erstellt automatisch `S:\03_Dienstleistungen\01_AWT\03_Technische Anfragen\TA XXXXXX KundeName Anwendung`
- **SPFx**: Kein Zugriff auf On-Prem-Netzlaufwerke
- **Klärungsbedarf**: Wird das Feature noch benötigt? Alternativen:
  - SharePoint-Dokumentenbibliothek mit Auto-Ordner
  - Power Automate + On-Premises Data Gateway
- **Aufwand**: Mittel–Hoch (je nach Lösung)

### 3. Termintreue-Auswertung / Statistik
- **Excel**: 4 Differenzberechnungen in Werktagen (Spalten AD–AG):
  - IST-Erledigungstermin vs. Wunschtermin
  - Wunschtermin vs. Geplanter Termin
  - IST vs. Ursprünglich geplanter Termin
  - IST vs. Verschobener Termin
  - Grün/Rot-Farbcodierung + Statistik-Blatt
- **SPFx**: Keine Auswertungsansicht
- **Empfohlene Lösung**: Neue `Statistik`-View mit Berechnungen + ggf. Charts
- **Aufwand**: Mittel

### 4. Erledigungstermin
- **Excel**: Eigenes Feld (Spalte I) für das tatsächliche Erledigungsdatum
- **SPFx**: Status wird auf `abgeschlossen` gesetzt, aber **kein** explizites Erledigungsdatum gespeichert
- **Empfohlene Lösung**: Neues Feld in der SharePoint-Liste + beim Abschließen automatisch `today()` eintragen
- **Aufwand**: Klein

### 5. Kunden-Anwendungen-Verwaltung (Settings-UI)
- **Excel**: Hilfsblatt wird beim Öffnen automatisch befüllt
- **SPFx**: `IKundenAnwendungItem`-Interface + Service-Methoden (`getKundenAnwendungen`, `addKundenAnwendung`, `deleteKundenAnwendung`) existieren bereits, aber die **Settings-UI hat keinen Tab dafür**
- **Empfohlene Lösung**: CRUD-Oberfläche für `TA_Kunden_Anwendungen` in `Settings.tsx` ergänzen
- **Aufwand**: Klein

---

## Priorisierte Maßnahmen

### Muss (funktionale Lücken)
| # | Maßnahme | Aufwand |
|---|---|---|
| 1 | **"prüfen"-Schwelle auf 3 Werktage** anpassen in `evaluateStatuses()` | Klein |
| 2 | **Erledigungsdatum** beim Abschließen einer TA automatisch speichern | Klein |
| 3 | **Kosten-Eingabefelder** in `TaDetail.tsx` ergänzen | Klein |

### Soll (Feature-Parität mit Excel)
| # | Maßnahme | Aufwand |
|---|---|---|
| 4 | **E-Mail-Benachrichtigungen** via Power Automate | Mittel |
| 5 | **Termintreue-Statistik** als neue View | Mittel |
| 6 | **Kunden-Anwendungen-UI** in Settings | Klein |

### Klärungsbedarf
| # | Maßnahme | Frage |
|---|---|---|
| 7 | **Netzlaufwerk-Ordner** | Wird das Feature noch gebraucht oder reicht SharePoint? |

---

## Bewusst verbessert (kein Handlungsbedarf)

- **Mehrbenutzerfähigkeit** — SharePoint-Listen statt Dateisperre auf Netzlaufwerk
- **People-Picker** — statt hardcodierter VBA-Mitarbeiterliste
- **Dynamische Kategorien** — SharePoint-Liste statt VBA-Code
- **TA-Nummernformat** — `TA-YYYY-NNN` statt `JJMMNN`
- **Mehr Verschiebungsgründe** — 13 vs. 6 im Excel
- **Konfigurierbarer Delay-Threshold** — statt hardcoded im VBA
- **Kein `Application.Quit`** — App bleibt nach TA-Erstellung offen
- **Kein fragiler Zustandsspeicher** — SharePoint-Felder statt versteckte Spalte AK
