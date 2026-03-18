# Claude Code Context — TA-Management

## What This Project Is

An SPFx 1.22.2 WebPart that manages **Technische Anfragen** (Technical Inquiries) for Barlog. It runs as a single-page app inside SharePoint Online at `barlog.sharepoint.com/sites/TechnischeAnfragen`. The UI uses React 17 + Fluent UI 8 with a glassmorphism SCSS design.

## Architecture

- **Entry point**: `src/webparts/taManagement/TaManagementWebPart.ts` — SPFx lifecycle, initializes PnPjs, renders root React component
- **Root component**: `src/webparts/taManagement/components/TaManagement.tsx` — state machine, data loading, view routing (6 views via `AppView` enum)
- **Data layer**: `src/webparts/taManagement/services/SharePointService.ts` — singleton service, all SharePoint REST/PnPjs v2 CRUD operations
- **Types**: `src/webparts/taManagement/models/types.ts` — all interfaces, enums, constants

### Views (AppView enum)

| View | Component | |
|---|---|---|
| Dashboard | `Dashboard.tsx` | KPI tiles + recent activity |
| NeueTa | `NeueTa.tsx` | Create form, project auto-fill from CSV |
| TerminPlanen | `TerminPlanen.tsx` | TAs needing scheduling |
| TaDetail | `TaDetail.tsx` | Detail/edit, schedule/reschedule/complete |
| AlleTas | `AlleTas.tsx` | Filterable list of all TAs |
| Settings | `Settings.tsx` | Admin: categories + config |

### SharePoint Lists

- `Technische_Anfragen` — main TA items (22+ fields including lookups)
- `TA_Kategorien` — category helper table
- `TA_Kunden_Anwendungen` — customer ↔ application mapping
- `TA_Config` — key-value config (`DelayThresholdDays`)
- CSV file `Projektliste.csv` read from SharePoint doc library (semicolon-delimited, German number format)

## Key Patterns

### Status Evaluation
`evaluateStatuses()` in `SharePointService` auto-computes status by comparing `Geplanter_x0020_Termin` to today:
- No date → `Termin planen`
- Past date → `überfällig`
- Within threshold → `prüfen`
- Future → `läuft planmäßig`
- Completed items are skipped

### TA Number Generation
Format `TA-YYYY-NNN` — `getNextTaNumber()` queries all TAs for current year, finds max, increments.

### Reschedule Tracking
When postponing: original date stored in `field_22`, reason in `field_21`. 13 predefined reasons in `VERSCHIEBUNG_GRUENDE` constant.

### Delay Reason
If scheduling happens > N days after creation (configurable via `DelayThresholdDays`), user must pick from `INITIAL_DELAY_REASONS`.

## Build & Dev

```bash
npm install              # Install deps (requires Node >=22.14 <23)
npm start                # Dev server on https://localhost:4321
npm run build            # Tests + production .sppkg package
npm run clean            # Clean artifacts
```

Build output: `sharepoint/solution/ta-management.sppkg`

## Important Notes

- PnPjs v2 (not v3/v4) — import paths are `@pnp/sp/presets/all`, `@pnp/sp/profiles`
- SharePoint field internal names use encoded spaces (`_x0020_`) — see `types.ts` for mappings
- The CSV parser handles German number format (comma decimals, period thousands)
- `skipFeatureDeployment: true` — solution can be deployed tenant-wide
- SCSS uses `@import '~@fluentui/react/dist/sass/mixins'` for Fluent design tokens
- All German UI strings are hardcoded (no i18n beyond SPFx loc scaffolding)

## Code Style

- TypeScript strict mode via SPFx build rig
- React class components (not hooks) for the root component; functional components for leaf UI
- Singleton pattern for `SharePointService` (`SharePointService.getInstance()`)
- Async/await throughout for SharePoint calls with try/catch error handling
