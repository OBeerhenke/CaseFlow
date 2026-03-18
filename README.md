# TA-Management — Technische Anfragen Dashboard

A **SharePoint Framework (SPFx) WebPart** for managing **Technische Anfragen** (Technical Inquiries) at Barlog. Provides a single-page dashboard with KPI tracking, deadline management, automatic status evaluation, and full CRUD operations against SharePoint Online lists.

![SPFx 1.22.2](https://img.shields.io/badge/SPFx-1.22.2-green.svg)
![React 17](https://img.shields.io/badge/React-17.0.1-blue.svg)
![Node >=22.14](https://img.shields.io/badge/Node-%3E%3D22.14-brightgreen.svg)

---

## Features

- **KPI Dashboard** — At-a-glance tiles showing overdue, on-track, review, and plan counts with recent activity feed
- **Create TAs** — New TA form with auto-generated `TA-YYYY-NNN` numbering and project data auto-fill from CSV
- **Deadline Scheduling** — Schedule deadlines with responsible person assignment via People Picker
- **Auto Status Evaluation** — Compares planned dates to today and auto-updates status (`überfällig` / `prüfen` / `läuft planmäßig` / `Termin planen`)
- **Reschedule with Reason** — Postpone deadlines with mandatory reason selection (13 predefined reasons tracked for reporting)
- **Filterable TA List** — Full-text search + status/priority filters across all TAs
- **Settings Panel** — Manage categories (Kategorien) and app configuration (delay threshold)
- **Multi-Host Support** — Runs in SharePoint, Teams Personal App, Teams Tab, and SharePoint Full Page

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
src/webparts/taManagement/
├── TaManagementWebPart.ts          # SPFx entry point (init, render, theme)
├── components/
│   ├── TaManagement.tsx            # Root SPA component (state, routing, data)
│   ├── Dashboard.tsx               # KPI tiles + recent activity
│   ├── NeueTa.tsx                  # New TA creation form
│   ├── AlleTas.tsx                 # Filterable/searchable TA list
│   ├── TaDetail.tsx                # Detail view (edit, schedule, complete)
│   ├── TerminPlanen.tsx            # TAs needing deadline scheduling
│   ├── Settings.tsx                # Admin: categories + config
│   ├── NavBar.tsx                  # Bottom navigation (5 tabs)
│   ├── KpiTile.tsx                 # Reusable KPI tile
│   ├── StatusPill.tsx              # Color-coded status badge
│   └── Modal.tsx                   # Reschedule modal
├── models/
│   └── types.ts                    # Interfaces, enums, constants
└── services/
    └── SharePointService.ts        # Singleton — all SharePoint CRUD ops
```

## SharePoint Lists

| List | Purpose |
|---|---|
| `Technische_Anfragen` | Main TA items (CRUD, up to 2000 items) |
| `TA_Kategorien` | Category helper table |
| `TA_Kunden_Anwendungen` | Customer ↔ Application mapping |
| `TA_Config` | Key-value app configuration |
| CSV: `Projektliste.csv` | Project data (semicolon-delimited, read from SharePoint document library) |

## Prerequisites

- **Node.js** >= 22.14.0, < 23.0.0
- **Heft CLI** — `npm install -g @rushstack/heft`
- Access to the target SharePoint Online tenant (`barlog.sharepoint.com`)

## Getting Started

```bash
# Clone the repository
git clone https://github.com/OBeerhenke/ta-management.git
cd ta-management

# Install dependencies
npm install

# Start dev server (serves on https://localhost:4321)
npm start
```

Then open the SharePoint Workbench:
```
https://barlog.sharepoint.com/sites/TechnischeAnfragen/_layouts/15/workbench.aspx
```

## Build & Deploy

```bash
# Production build (runs tests + packages solution)
npm run build

# Clean build artifacts
npm run clean
```

The production build produces `sharepoint/solution/ta-management.sppkg` which can be uploaded to the SharePoint App Catalog. The solution supports tenant-wide deployment (`skipFeatureDeployment: true`).

## Utility Scripts

| Script | Purpose |
|---|---|
| `error-check.js` | Check SharePoint API error responses |
| `parse-har.js` | Parse HAR files to find HTTP 400 responses (debugging) |
| `queryProjekteFields.js` | Query "Projekte" list fields via PnPjs Node (requires `SP_ID` / `SP_SECRET` env vars) |

## Configuration

The WebPart reads its configuration from the `TA_Config` SharePoint list:

| Key | Description |
|---|---|
| `DelayThresholdDays` | Number of days after TA creation before a delay reason is required when scheduling |

## References

- [SharePoint Framework Documentation](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
- [PnPjs v2 Documentation](https://pnp.github.io/pnpjs/)
- [Fluent UI React](https://developer.microsoft.com/fluentui)
- [Heft Build System](https://heft.rushstack.io/)