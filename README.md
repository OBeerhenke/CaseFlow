# CaseFlow — Vorgangs-Management für SharePoint

A **SharePoint Framework (SPFx) WebPart** for configurable case/inquiry management: **Erfassen → Zuordnen → Planen → Überwachen → Abschließen**. Provides a single-page dashboard with KPI tracking, deadline management, automatic status evaluation, and full CRUD operations against SharePoint Online lists — no additional licensing, servers, or user management required.

![SPFx 1.22.2](https://img.shields.io/badge/SPFx-1.22.2-green.svg)
![React 17](https://img.shields.io/badge/React-17.0.1-blue.svg)
![Node >=22.14](https://img.shields.io/badge/Node-%3E%3D22.14-brightgreen.svg)

---

## Features

- **KPI Dashboard** — At-a-glance tiles showing overdue, on-track, review, and plan counts with recent activity feed
- **Create Cases** — New case form with auto-generated numbering and project data auto-fill from CSV
- **Deadline Scheduling** — Schedule deadlines with responsible person assignment via People Picker
- **Auto Status Evaluation** — Compares planned dates to today and auto-updates status (`überfällig` / `prüfen` / `läuft planmäßig` / `Termin planen`)
- **Reschedule with Reason** — Postpone deadlines with mandatory reason selection (predefined reasons tracked for reporting)
- **Filterable Case List** — Full-text search + status/priority filters across all cases
- **Settings Panel** — Manage categories (Kategorien) and app configuration (delay threshold)
- **Multi-Host Support** — Runs in SharePoint, Teams Personal App, Teams Tab, and SharePoint Full Page

> **Status:** This repository is being refactored from a single-customer solution into a configurable, multi-tenant product. See `.scratch/caseflow/issues/` in the source project for the roadmap (Config Layer + Field Mapping, White-Labeling, Setup-Script, Power Automate Templates).

## Tech Stack

| Layer | Technology |
|---|---|
| Framework | SharePoint Framework (SPFx) 1.22.2 |
| UI | React 17 + Fluent UI React 8 |
| Language | TypeScript ~5.8 |
| SharePoint API | PnPjs v2 (`@pnp/sp`) |
| Build | Heft (Rush Stack) via `@rushstack/heft` |
| Styling | SCSS modules (glassmorphism design) |
| Testing | Jest (heft-jest) |

## Project Structure

```
src/webparts/caseFlow/
├── CaseFlowWebPart.ts       # SPFx entry point (init, render, theme)
├── components/
│   ├── App.tsx              # Root SPA component (state, routing, data)
│   ├── Dashboard.tsx        # KPI tiles + recent activity
│   ├── NewCase.tsx          # New case creation form
│   ├── CaseList.tsx         # Filterable/searchable case list
│   ├── CaseDetail.tsx       # Detail view (edit, schedule, complete)
│   ├── Schedule.tsx         # Cases needing deadline scheduling
│   ├── Analytics.tsx        # Punctuality / delay analytics
│   ├── Settings.tsx         # Admin: categories + config
│   ├── NavBar.tsx           # Bottom navigation
│   ├── KpiTile.tsx          # Reusable KPI tile
│   ├── StatusPill.tsx       # Color-coded status badge
│   └── Modal.tsx            # Reschedule modal
├── models/
│   └── types.ts             # Interfaces, enums, constants
└── services/
    └── SharePointService.ts # Singleton — all SharePoint CRUD ops
```

## SharePoint Lists

| List | Purpose |
|---|---|
| `CaseFlow_Cases` | Main case items (CRUD, up to 2000 items) |
| `CaseFlow_Categories` | Category helper table |
| `CaseFlow_CustomerApplications` | Customer ↔ Application mapping |
| `CaseFlow_Config` | Key-value app configuration |
| CSV: project data file | Project data (semicolon-delimited, read from SharePoint document library) |

> List names, field mappings, and status definitions are currently hardcoded defaults. Making these tenant-configurable is tracked as a follow-up (Config Layer + Field Mapping).

## Prerequisites

- **Node.js** >= 22.14.0, < 23.0.0
- **Heft CLI** — `npm install -g @rushstack/heft`
- Access to a SharePoint Online tenant

## Getting Started

```bash
# Clone the repository
git clone <repo-url>
cd CaseFlow

# Install dependencies
npm install

# Start dev server (serves on https://localhost:4321)
npm start
```

Then open the SharePoint Workbench (update `config/serve.json` with your tenant/site first):
```
https://{your-tenant}.sharepoint.com/sites/{your-site}/_layouts/15/workbench.aspx
```

## Build & Deploy

```bash
# Production build (runs tests + packages solution)
npm run build

# Clean build artifacts
npm run clean
```

The production build produces `sharepoint/solution/case-flow.sppkg` which can be uploaded to the SharePoint App Catalog. The solution supports tenant-wide deployment (`skipFeatureDeployment: true`).

## Utility Scripts

| Script | Purpose |
|---|---|
| `error-check.js` | Check SharePoint API error responses |
| `parse-har.js` | Parse HAR files to find HTTP 400 responses (debugging) |
| `queryProjekteFields.js` | Query list fields via PnPjs Node (requires `SP_ID` / `SP_SECRET` env vars) |

## Configuration

The WebPart reads its configuration from the `CaseFlow_Config` SharePoint list:

| Key | Description |
|---|---|
| `DelayThresholdDays` | Number of days after case creation before a delay reason is required when scheduling |

## References

- [SharePoint Framework Documentation](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
- [PnPjs v2 Documentation](https://pnp.github.io/pnpjs/)
- [Fluent UI React](https://developer.microsoft.com/fluentui)
- [Heft Build System](https://heft.rushstack.io/)
