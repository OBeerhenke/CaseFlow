# CaseFlow — Test-Anleitung für SharePoint

So testest du CaseFlow in deiner SharePoint-Umgebung.

---

## 1. Build erstellen

```bash
cd ~/Projects/CaseFlow
npm run build
```

Ausgabe: `sharepoint/solution/case-flow.sppkg`

> **Hinweis:** Bei Node 25 (deine Version) gibt es Engine-Warnungen, aber der Build läuft durch. Sollte es Probleme geben, installiere Node 22: `nvm install 22 && nvm use 22 && npm install && npm run build`.

## 2. App hochladen

1. Gehe zu deinem SharePoint **App Catalog**:
   - `https://{dein-tenant}.sharepoint.com/sites/AppCatalog`
   - Oder: `https://{dein-tenant}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webPartMarketplace/:/WebPartMarketplace/appCatalog`
2. Navigiere zu **"Apps for SharePoint"** → **"App Catalog"**
3. Klicke auf **"Upload"** → wähle `case-flow.sppkg`
4. Hake **"Make this solution available to all sites"** an (skipFeatureDeployment: true)
5. Klicke **"Deploy"**

## 3. SharePoint-Listen anlegen

CaseFlow braucht 4 Listen. Du hast zwei Optionen:

### Option A: Setup-Script (empfohlen)

```bash
# Im caseflow-setup/ Ordner
cd caseflow-setup
npm install

# Trockentest (macht nichts, zeigt nur an was passieren würde)
node index.js \
  --siteUrl "https://{dein-tenant}.sharepoint.com/sites/{deine-site}" \
  --appId "{deine-App-ID}" \
  --appSecret "{dein-App-Secret}" \
  --dry-run

# Wirklich ausführen
node index.js \
  --siteUrl "https://{dein-tenant}.sharepoint.com/sites/{deine-site}" \
  --appId "{deine-App-ID}" \
  --appSecret "{dein-App-Secret}"
```

> **App-ID/Secret:** Du brauchst eine SharePoint-Add-In-Registrierung (AppPrincipal). Erstelle eine über `https://{tenant}-admin.sharepoint.com/_layouts/15/appregnew.aspx` → "Generate" → speichere ID + Secret.

### Option B: Manuell

Erstelle 4 leere SharePoint-Listen in deiner Site:

| Listenname | Typ | Spalten |
|---|---|---|
| `CaseFlow_Cases` | Benutzerdefiniert | Siehe Konfigurationsbogen (ca. 30 Felder inkl. Lookups Ersteller/Verantwortlicher) |
| `CaseFlow_Categories` | Benutzerdefiniert | `Title` (Text), `Email` (Text) |
| `CaseFlow_CustomerApplications` | Benutzerdefiniert | `Title` (Text), `Anwendung` (Text) |
| `CaseFlow_Config` | Benutzerdefiniert | `Title` (Text), `Value` (Text) |

> **Tipp:** Die Config-Liste muss initial einen Eintrag haben:
> - `Title = "ConfigSchemaVersion"` → `Value = "1.1"`

## 4. WebPart zu einer Seite hinzufügen

1. Gehe zu deiner SharePoint-Site
2. Erstelle eine neue **"Wiki"- oder "Modern"-Seite** (oder nutze die Workbench)
3. Füge eine **neue WebPart-Zone / Abschnitt** ein
4. Wähle **"CaseFlow"** aus der WebPart-Liste

> Alternativ: Workbench direkt öffnen:
> ```
> https://{dein-tenant}.sharepoint.com/sites/{deine-site}/_layouts/15/workbench.aspx
> ```

## 5. WebPart konfigurieren

1. Klicke auf das **Stift-Icon** (Edit) → dann auf das **CaseFlow-WebPart**
2. Es öffnet sich das **Property Pane** (rechte Seitenleiste)
3. **Standard-Listen** sollten passen, wenn du die gleichen Namen wie oben verwendet hast
4. Falls deine Felder anders heißen, trage das **Feld-Mapping JSON** ein:
   ```json
   {"field_8": "MeinKundenfeld", "field_6": "MeinTerminfeld"}
   ```
5. Klicke **"Apply"** → das WebPart lädt neu

## 6. Funktionen testen

- **Dashboard**: KPI-Kacheln sichtbar? (überfällig, planen, planmäßig, prüfen)
- **Neuen Case anlegen**: Formular öffnet sich? Speichern funktioniert?
- **Alle Cases**: Liste sichtbar? Sortieren/Filtern?
- **Termin planen**: Datum wählbar? Status wechselt?
- **Detail-Ansicht**: Bearbeiten, Verschieben, Abschließen?
- **Analytics**: Statistik wird angezeigt?
- **Einstellungen**: Kategorien + App-Einstellungen editierbar?

## 7. Power Automate testen (optional)

Importiere die Flow-Vorlagen aus `power-automate/templates/`:
1. Gehe zu **https://make.powerautomate.com**
2. Wähle die richtige Umgebung (dein Tenant)
3. **"My flows"** → **"Import"** → wähle `.zip`
4. Verbinde die SharePoint-Connector-Schritte mit deinen Listen

## 8. Häufige Probleme

| Problem | Lösung |
|---|---|
| WebPart zeigt "Loading" ewig | Prüfe ob die Config-Liste existiert und `ConfigSchemaVersion` enthält |
| Feld XYZ wird nicht gespeichert | Property Pane Feld-Mapping prüfen: der SP-Feldname muss stimmen |
| Keine Projektdaten-Autovervollständigung | `ProjectsCsvPath` in den Settings eintragen |
| 403 / Access Denied | App Catalog-Deployment mit skipFeatureDeployment nötig, oder User hat keine Berechtigung |
| Build failed (Node-Version) | Mit Node 22 bauen (`nvm use 22`) |

## 9. Workbench vs. echte Seite

**Workbench** (`/_layouts/15/workbench.aspx`):
- Schnellster Weg zum Testen
- Arbeitet gegen echte SharePoint-Daten
- Kein Deployment nötig (lädt lokal von localhost:4321)

**Echte SharePoint-Seite** nach Deployment:
- Zeigt das Verhalten wie für Endkunden
- Erfordert .sppkg Upload in App Catalog
- Erfordert `gulp trust-dev-cert` für HTTPS (oder Zertifikat manuell akzeptieren)
