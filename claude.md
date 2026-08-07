# Claude Code Context — CaseFlow

## What This Project Is

CaseFlow is a **generic, configurable case/inquiry management product** for SharePoint Online, built as an SPFx 1.22.2 WebPart. It is being refactored from a single-customer solution (originally built for Barlog as "TA-Management" / Technische Anfragen) into a multi-tenant product sellable to German Mittelstand customers across industries.

Core workflow pattern: **Erfassen → Zuordnen → Planen → Überwachen → Abschließen** (create → assign → schedule → track → complete), with a status traffic-light, postponement-reason tracking, and analytics. The UI uses React 17 + Fluent UI 8 with a glassmorphism SCSS design.

Product plan and ticket breakdown: see `.scratch/caseflow/issues/` in the source project (`ta-management` repo) for the roadmap — Config Layer + Field Mapping, StatusEvaluation Engine, White-Labeling, Setup-Script, Power Automate Templates, Branding.

## Current Status

**Ticket 1 (Repo-Setup + CaseFlow-Grundgerüst) is complete.** This is a pure rename/rebrand of the original codebase — all business logic and behavior is unchanged. List names, field mappings, and status/reason labels are still hardcoded (default) values, not yet tenant-configurable. That is Ticket 2's scope.

## Architecture

- **Entry point**: `src/webparts/caseFlow/CaseFlowWebPart.ts` — SPFx lifecycle, initializes PnPjs, renders root React component
- **Root component**: `src/webparts/caseFlow/components/App.tsx` — state machine, data loading, view routing (7 views via `AppView` enum)
- **Data layer**: `src/webparts/caseFlow/services/SharePointService.ts` — singleton service, all SharePoint REST/PnPjs v2 CRUD operations
- **Types**: `src/webparts/caseFlow/models/types.ts` — all interfaces, enums, constants

### Views (AppView enum)

| View | Component | |
|---|---|---|
| Dashboard | `Dashboard.tsx` | KPI tiles + recent activity |
| NewCase | `NewCase.tsx` | Create form, project auto-fill from CSV |
| Schedule | `Schedule.tsx` | Cases needing scheduling |
| CaseDetail | `CaseDetail.tsx` | Detail/edit, schedule/reschedule/complete |
| CaseList | `CaseList.tsx` | Filterable list of all cases |
| Analytics | `Analytics.tsx` | Punctuality / delay analytics |
| Settings | `Settings.tsx` | Admin: categories + config |

### SharePoint Lists (hardcoded defaults — Ticket 2 makes these configurable)

- `CaseFlow_Cases` — main case items (22+ fields including lookups)
- `CaseFlow_Categories` — category helper table
- `CaseFlow_CustomerApplications` — customer ↔ application mapping
- `CaseFlow_Config` — key-value config (`DelayThresholdDays`)
- CSV file (project data) read from SharePoint doc library (semicolon-delimited, German number format)

## Key Patterns

### Status Evaluation
`evaluateStatuses()` in `SharePointService` auto-computes status by comparing the planned date to today:
- No date → `Termin planen`
- Past date → `überfällig`
- Within threshold → `prüfen`
- Future → `läuft planmäßig`
- Completed items are skipped

(Ticket 3 extracts this into a pure, testable `StatusEvaluation Engine` with configurable status values/thresholds.)

### Case Number Generation
`getNextCaseNumber()` queries all cases for current year, finds max, increments.

### Reschedule Tracking
When postponing: original date stored in `field_22`, reason in `field_21`. Predefined reasons in `VERSCHIEBUNG_GRUENDE` constant (currently Barlog-specific business vocabulary — becomes configurable per customer in Ticket 2).

### Delay Reason
If scheduling happens > N days after creation (configurable via `DelayThresholdDays`), user must pick from `INITIAL_DELAY_REASONS`.

## Build & Dev

```bash
npm install              # Install deps (requires Node >=22.14 <23; warns but works on newer Node)
npm start                # Dev server on https://localhost:4321 (update config/serve.json tenant URL first)
npm run build            # Tests + production .sppkg package
npm run clean            # Clean artifacts
```

Build output: `sharepoint/solution/case-flow.sppkg`

## Important Notes

- PnPjs v2 (not v3/v4) — import paths are `@pnp/sp/presets/all`, `@pnp/sp/profiles`
- SharePoint field internal names use encoded spaces / are auto-generated (`field_N`) — see `types.ts` for mappings
- The CSV parser handles German number format (comma decimals, period thousands)
- `skipFeatureDeployment: true` — solution can be deployed tenant-wide
- SCSS uses `@import '~@fluentui/react/dist/sass/mixins'` for Fluent design tokens
- All German UI strings/business vocabulary (Kunde, Termin, Bemerkung, etc.) are intentionally kept — target market is German Mittelstand, this is not Barlog-specific and stays as-is. Only Barlog-specific identity (tenant URLs, list names, product name, code identifiers) was renamed.
- No i18n beyond SPFx loc scaffolding

## Code Style

- TypeScript strict mode via SPFx build rig
- React class components (not hooks) for the root component; functional components for leaf UI
- Singleton pattern for `SharePointService` (`SharePointService.instance`)
- Async/await throughout for SharePoint calls with try/catch error handling
- Variable naming: `caseItem` for a single case (not `case` — reserved word), `cases` for arrays
