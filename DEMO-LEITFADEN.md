# TA-Management v1 — Demo-Leitfaden

> Leitfaden für die Vorstellung der ersten Version des TA-Management WebParts beim Kunden (Barlog).

---

## 1. Einstieg & Motivation

- Das bisherige Excel-Tool (`Z_Übersicht Technische Anfragen.xlsm`) wird abgelöst
- **Vorteile gegenüber Excel**:
  - Mehrbenutzerfähig (keine Dateisperre)
  - Direkt in SharePoint — kein Öffnen/Speichern/Schließen
  - Automatische Statusaktualisierung in Echtzeit
  - Kein VBA-Sicherheitsrisiko, keine Makro-Warnungen
- Zugriff über SharePoint → keine Installation nötig

---

## 2. Navigation (6 Bereiche)

| Icon | Bereich | Funktion |
|---|---|---|
| 🏠 Home | **Dashboard** | KPI-Übersicht, letzte Aktivitäten |
| ➕ Neue TA | **Neue TA** | Technische Anfrage anlegen |
| 📋 Alle TAs | **Alle TAs** | Suchen, filtern, sortieren |
| 📊 Statistik | **Statistik** | Termintreue-Auswertung |
| 📅 Termine | **Termine** | Terminplanung & -übersicht |
| ⚙️ Einstellungen | **Einstellungen** | Kategorien, Konfiguration |

---

## 3. Demo-Ablauf (empfohlene Reihenfolge)

### Schritt 1: Dashboard zeigen

- 4 KPI-Kacheln: **Überfällig** / **Prüfen** / **Planmäßig** / **Termin planen**
- Letzte Aktivitäten (zuletzt geänderte TAs)
- → *„Auf einen Blick sehen, wo Handlungsbedarf besteht"*

### Schritt 2: Neue TA anlegen

- Projekt-Nr. eingeben → **Kunde, Anwendung, Material werden automatisch befüllt** (aus Projektliste.csv)
- Kategorie auswählen (dynamisch aus SharePoint-Liste, nicht hartcodiert)
- Priorität setzen (Hoch / Mittel / Niedrig)
- Wunschtermin + Budget eintragen
- TA-Nummer wird automatisch vergeben: `TA-2026-001`
- → *„Weniger Tipparbeit, keine Fehler bei der Zuordnung"*

### Schritt 3: Alle TAs durchsuchen

- **Volltextsuche** über alle Felder
- **Statusfilter** (alle 5 Status einzeln filterbar)
- Sortierung nach Priorität
- Klick auf TA → Detailansicht
- → *„Sofort finden, was man sucht"*

### Schritt 4: TA-Detail & Termin planen

- Geplanten Termin setzen + **Verantwortlichen** per People-Picker zuweisen
- Status springt automatisch auf `läuft planmäßig`
- **Verzögerungsgrund**: Wird automatisch abgefragt, wenn der geplante Termin zu weit vom Erstellungsdatum abweicht (Schwelle konfigurierbar)
- **Kosten bearbeiten**: Budget bei Start, IST-Kosten, geplante Kosten direkt editierbar
- → *„Dokumentierte Nachvollziehbarkeit, warum Termine so gewählt wurden"*

### Schritt 5: Terminverschiebung vorführen

- Geplanten Termin ändern → **Modal mit 13 Verschiebungsgründen** erscheint
- Alter Termin wird automatisch als „Ursprünglicher Termin" gesichert
- → *„Jede Verschiebung wird transparent dokumentiert"*

### Schritt 6: Statusautomatik erklären

Die 5 Status werden **automatisch** berechnet — kein manuelles Setzen nötig:

| Status | Regel |
|---|---|
| `Termin planen` | Kein geplanter Termin gesetzt |
| `läuft planmäßig` | Geplanter Termin > 3 Werktage in der Zukunft |
| `prüfen` | Geplanter Termin innerhalb der nächsten **3 Werktage** |
| `überfällig` | Geplanter Termin liegt in der Vergangenheit |
| `abgeschlossen` | Manuell abgeschlossen → Erledigungsdatum wird automatisch gesetzt |

→ *„Genau wie im Excel, aber zuverlässiger und für alle Nutzer gleichzeitig aktuell"*

### Schritt 7: Statistik & Termintreue

- **4 KPIs**: Termintreue (%), Abgeschlossen, Verschoben, Ø Abweichung (Werktage)
- **Statusverteilung** als gestapeltes Balkendiagramm
- **Verschiebungsgründe** als Ranking
- **Detail-Tabelle**: Pro TA die Abweichungen in Werktagen (IST vs. Wunsch / Plan / Original)
- Zeitraumfilter: Alle / Aktuelles Jahr / Aktuelles Quartal
- → *„Dieselbe Auswertung wie das Statistik-Blatt im Excel, aber live und interaktiv"*

### Schritt 8: Einstellungen

- **Kategorien verwalten**: Hinzufügen, Löschen (dynamisch statt VBA-Code ändern)
- **Delay-Threshold**: Konfigurierbar, ab wie vielen Tagen ein Verzögerungsgrund bei der Erstplanung gefragt wird
- → *„Fachabteilung kann selbst anpassen, ohne IT-Änderung"*

---

## 4. Was ist neu / besser als im Excel?

| Thema | Excel (vorher) | WebPart (jetzt) |
|---|---|---|
| Mehrbenutzerzugriff | Dateisperre | Gleichzeitig nutzbar |
| Verantwortlicher | Hardcodierte VBA-Liste | People-Picker aus Active Directory |
| TA-Nummern | `260401` (kryptisch) | `TA-2026-001` (sprechend) |
| Verschiebungsgründe | 6 Gründe | 13 Gründe (feinere Auswertung) |
| Kategorien | Im VBA-Code | In SharePoint-Liste (selbst verwaltbar) |
| Konfiguration | Hardcoded | Einstellungen-Seite |
| Makro-Sicherheit | VBA-Warnmeldungen | Keine Makros nötig |

---

## 5. Was kommt in v2? (Ausblick)

| Feature | Lösung | Status |
|---|---|---|
| **E-Mail-Benachrichtigungen** | Power Automate Flow | Geplant |
| Erstellung → Mail an AWT + Ablage | | |
| Erinnerung bei `prüfen` | | |
| Eskalation bei `überfällig` | | |
| **Kunden-Anwendungen verwalten** | Settings-UI | Backend fertig, UI folgt |
| **Netzlaufwerk-Ordner** | Klärungsbedarf | SharePoint-Bibliothek als Alternative? |

---

## 6. Offene Fragen an den Kunden

- [ ] Wird die automatische Ordner-Erstellung auf `S:\` noch benötigt, oder reicht eine SharePoint-Dokumentenbibliothek?
- [ ] Sollen E-Mail-Empfänger konfigurierbar sein oder fix auf `awt@barlog.de`?
- [ ] Gibt es weitere Kategorien oder Verschiebungsgründe, die ergänzt werden sollen?
- [ ] Soll es eine Export-Funktion (CSV/Excel) für die Statistik geben?
