# CaseFlow — Konfigurationsbogen für Neukunden

Dieses Dokument erfasst alle kundenspezifischen Einstellungen, die für die Inbetriebnahme von CaseFlow benötigt werden. Ausgefüllt vom Kunden, Grundlage für das Deployment-Projekt.

---

## 1. SharePoint-Umgebung

| Feld | Wert | Beispiel |
|---|---|---|
| SharePoint-Tenant-URL | | `barlog.sharepoint.com` |
| Site-URL (wo CaseFlow läuft) | | `https://barlog.sharepoint.com/sites/TechnischeAnfragen` |
| Site-Name (für Workbench) | | `TechnischeAnfragen` |

## 2. Benennung (Entity Labels)

CaseFlow nennt einen Vorgang standardmäßig **"Anfrage"** (Mehrzahl: **"Anfragen"**). Das kann pro Kunde geändert werden.

| Feld | Wert | Beispiele |
|---|---|---|
| Bezeichnung (Einzahl) | | `Anfrage`, `Ticket`, `Case`, `Vorgang`, `Reklamation` |
| Bezeichnung (Mehrzahl) | | `Anfragen`, `Tickets`, `Cases`, `Vorgänge`, `Reklamationen` |

## 3. SharePoint-Listen

CaseFlow benötigt 4 SharePoint-Listen. Der Kunde kann bestehende Listen verwenden oder neue anlegen lassen.

### 3.1 Hauptliste (Cases)

| Feld | Wert | Standard |
|---|---|---|
| Listenname | | `CaseFlow_Cases` |
| Feld "Titel/Vorgangsnummer" (SP-Name) | | `Title` |
| Feld "Status" (SP-Name) | | `Status` |
| Feld "Kunde" (SP-Name) | | `field_8` |
| Feld "Endkunde" (SP-Name) | | `field_9` |
| Feld "Kontaktnummer" (SP-Name) | | `field_10` |
| Feld "Anwendung" (SP-Name) | | `field_11` |
| Feld "Material" (SP-Name) | | `field_12` |
| Feld "Priorität" (SP-Name) | | `field_13` |
| Feld "Potenzial" (SP-Name) | | `field_14` |
| Feld "Chance" (SP-Name) | | `field_15` |
| Feld "Kategorie" (SP-Name) | | `field_16` |
| Feld "Budget bei Start" (SP-Name) | | `field_17` |
| Feld "IST Kosten" (SP-Name) | | `field_18` |
| Feld "Geplante Kosten" (SP-Name) | | `field_19` |
| Feld "Differenz" (SP-Name) | | `field_20` |
| Feld "Grund für Verschiebung" (SP-Name) | | `field_21` |
| Feld "Ursprünglicher Termin" (SP-Name) | | `field_22` |
| Feld "Erstellungsdatum" (SP-Name) | | `field_4` |
| Feld "Wunschtermin" (SP-Name) | | `field_5` |
| Feld "Geplanter Termin" (SP-Name) | | `field_6` |
| Feld "Erledigungsdatum" (SP-Name) | | `Erledigungsdatum` |
| Feld "Bemerkung" (SP-Name) | | `field_2` |
| Feld "Aufgabenstellung" (SP-Name) | | `Aufgabenstellung` |
| Lookup-Feld "Ersteller" (SP-Name) | | `Ersteller` |
| Lookup-Feld "Verantwortlicher" (SP-Name) | | `Verantwortlicher` |
| Feld "SOP" (SP-Name) | | `SOP` |
| Feld "SegCode" (SP-Name) | | `SegCode` |
| Feld "AntwortIn" (SP-Name) | | `AntwortIn` |
| Feld "Zielpreis" (SP-Name) | | `Zielpreis` |
| Feld "Projektnummer" (SP-Name) | | `Projektnummer` |

> **Hinweis:** Die SP-Namen können vom Kunden über das Property Pane (WebPart-Einstellungen) als JSON überschrieben werden, z.B. `{"field_8": "Customer", "field_6": "PlannedDate"}`.

### 3.2 Kategorien-Liste

| Feld | Wert | Standard |
|---|---|---|
| Listenname | | `CaseFlow_Categories` |
| Felder | | `Title` (Titel), `Email` (E-Mail-Empfänger) |

### 3.3 Kunden-Anwendungen-Liste

| Feld | Wert | Standard |
|---|---|---|
| Listenname | | `CaseFlow_CustomerApplications` |
| Felder | | `Title` (Kunde), `Anwendung` |

### 3.4 Config-Liste

| Feld | Wert | Standard |
|---|---|---|
| Listenname | | `CaseFlow_Config` |
| Felder | | `Title` (Key), `Value` |

## 4. Status-Definitionen

Standardmäßig 5 Status. Der Kunde kann Namen und Farben ändern oder weitere Stati hinzufügen.

| Status-Key | Label (Standard) | Farbe (Hex) | Kunde |
|---|---|---|---|
| `Termin planen` | Termin planen | `#3B82F6` (blau) | |
| `läuft planmäßig` | läuft planmäßig | `#10B981` (grün) | |
| `prüfen` | prüfen | `#F59E0B` (gelb) | |
| `überfällig` | überfällig | `#EF4444` (rot) | |
| `abgeschlossen` | abgeschlossen | `#6B7280` (grau) | |

**Eigene Status-Definition (JSON):**
```json
[
  {"key": "Termin planen", "label": "...", "color": "#..."},
  {"key": "läuft planmäßig", "label": "...", "color": "#..."},
  ...
]
```

## 5. Thresholds / Regeln

| Parameter | Standard | Kunde | Beschreibung |
|---|---|---|---|
| `DelayThresholdDays` | `2` | | Tage nach Erstellung, ab wann ein Verzögerungsgrund beim Planen erforderlich ist |
| `PruefenTage` | `3` | | Werktage vor geplantem Termin, ab wann Status "prüfen" wird |

## 6. Projektdaten (Auto-Vervollständigung)

| Feld | Wert | Beispiel |
|---|---|---|
| CSV-Datei-Pfad (serverrelativ) | (leer = deaktiviert) | `/sites/TechnischeAnfragen/Projektliste/Projektliste.csv` |

> Die CSV wird semikolongetrennt erwartet, mit deutscher Zahlenformatierung (Komma-Dezimaltrenner, Punkt-Tausendertrenner). Spalten werden dynamisch anhand der Header-Zeile erkannt.

## 7. White-Labeling / Design

| Feld | Standard | Kunde |
|---|---|---|
| Primärfarbe (Hex) | `#3B82F6` (blau) | |
| Hintergrundfarbe (Hex) | `#f8fafc` | |
| Logo-URL | (leer = kein Logo) | |

## 8. Power Automate Integration

| Workflow | Aktiv? | Beschreibung |
|---|---|---|
| E-Mail bei Anlage | ☐ Ja ☐ Nein | Benachrichtigung an Verantwortlichen + Kategorie-Empfänger |
| E-Mail bei Terminplanung | ☐ Ja ☐ Nein | Bestätigungsmail |
| E-Mail bei Verschiebung | ☐ Ja ☐ Nein | Mail mit Grund + neuem Datum |
| E-Mail bei Abschluss | ☐ Ja ☐ Nein | Abschluss-Benachrichtigung |

## 9. Sonstiges

| Thema | Notiz |
|---|---|
| Gewünschtes Nummernformat | (Standard: `TA-YYYY-NNN`) |
| Benötigte Rollen / Berechtigungen | |
| Test-Zeitraum / Go-Live-Datum | |
| Ansprechpartner Kunde (Name, E-Mail) | |
| Ansprechpartner Technik (Name, E-Mail) | |

---

**Ausgefüllt am:** __________ **Durch:** __________
