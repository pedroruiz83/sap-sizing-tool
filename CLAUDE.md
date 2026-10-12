# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SAP Sizing Analysis Tool — a browser-based single-page application for infrastructure sizing analysis. No build step, no framework, no server required: open `index.html` directly in a browser.

## Running the Tool

```bash
# Any of the following work:
start sizing-tool/index.html          # Windows default browser
python -m http.server 8080            # from sizing-tool/ directory
npx serve sizing-tool                 # if Node.js is available
```

## Project Structure

```
sizing-tool/
├── index.html                    # Shell: header, sidebar nav, content mount point
├── css/styles.css                # All styles — CSS custom properties + component classes
├── js/
│   ├── app.js                    # SizingApp core: routing, state, shared utilities
│   └── modules/
│       ├── systemLandscape.js    # Module 01 — system inventory
│       ├── techWorkloadApp.js    # Module 02 — application layer metrics
│       ├── techWorkloadDb.js     # Module 03 — database layer metrics
│       ├── appWorkload.js        # Module 04 — business/application workload
│       └── sizingProjection.js   # Module 05 — aggregated sizing recommendations
```

## Architecture

**Routing and state** live in `js/app.js` as the `SizingApp` singleton. Each navigation click calls `SizingApp.navigateTo(moduleId)`, which renders the target module's HTML into `#content-area` and calls `module.init(app)`.

**Module contract** — every module file exposes a plain object with two methods:
```js
const MyModule = {
    render()     { return '<html string>'; },  // called first
    init(app)    { /* wire up events */    },  // called after innerHTML set
};
```

**Shared utilities** (defined at the bottom of `app.js`, available globally):
- `parseUploadedFile(file, callback)` — handles CSV (PapaParse) and XLSX (SheetJS)
- `initUploadZone(zoneId, inputId, onFile)` — wires drag-and-drop + click for an upload zone
- `showFileLoaded(containerId, fileName)` — shows green success banner
- `statCard(label, value, unit)` — returns stat card HTML string
- `downloadXlsx(rows, sheetName, fileName)` — exports array-of-objects to XLSX
- `detectCol(headers, regex)` — finds first matching header by regex
- `countBy(rows, col)` — frequency map of a column's values
- `percentile(arr, p)` — Pth percentile of a numeric array
- `CHART_COLORS` — 8-colour SAP-palette array for Chart.js

**CDN dependencies** (loaded in `index.html`, no install needed):
- Chart.js 4.4.0
- PapaParse 5.4.1
- SheetJS 0.20.0

## Styling Conventions

All CSS custom properties are in `:root` in `styles.css`. Key tokens:
- `--brand` `--brand-dark` `--brand-light` — primary blue
- `--header-bg` `--sidebar-bg` — dark nav surfaces
- `--bg-page` `--bg-card` — content backgrounds
- `--border` `--border-strong` — borders

Component classes follow a flat naming pattern: `.card`, `.upload-zone`, `.stats-grid`, `.stat-card`, `.charts-grid`, `.chart-card`, `.data-table`, `.badge-*`, `.btn`, `.btn-blue`.

## Adding a New Module

1. Create `js/modules/myModule.js` with the `render()` / `init(app)` interface.
2. Add a `<script src="js/modules/myModule.js">` tag in `index.html` (before `app.js`).
3. Register in `MODULES` and `MODULE_ORDER` inside `js/app.js`.
4. Add a `<li class="nav-item" data-module="my-module">` entry in the sidebar in `index.html`.

## File Upload Conventions

Each module should:
- Use `initUploadZone` + `parseUploadedFile` for all file handling.
- Call `app.setModuleData(moduleId, { rows, fileName, headers, analysis })` after processing.
- Call `app.markCompleted(moduleId)` once analysis is ready.
- Provide a **Download Analysis** button using `downloadXlsx`.
- Include a **Next** button that calls `app.nextModule()`.

Column detection is intentionally flexible — use `detectCol(headers, /regex/i)` to find columns by pattern rather than exact name, so the tool works with varied export formats.
