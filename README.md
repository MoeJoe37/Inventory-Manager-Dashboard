# Inventory Manager

A browser-based inventory monitoring dashboard for Odoo-style inventory exports. The app imports XLSX/CSV inventory files, analyzes stock, value, locations, categories, product status, and can compare the current inventory file against another XLSX file through a dedicated comparison dashboard.

The project is fully static: it runs directly from `index.html` or from GitHub Pages without a backend, database, build step, or external JavaScript dependency.

---

## Table of contents

- [Overview](#overview)
- [Main features](#main-features)
- [Comparison features](#comparison-features)
- [Supported import format](#supported-import-format)
- [How to use](#how-to-use)
- [GitHub Pages deployment](#github-pages-deployment)
- [Project structure](#project-structure)
- [Browser support](#browser-support)
- [Data privacy](#data-privacy)
- [Troubleshooting](#troubleshooting)
- [Version notes](#version-notes)

---

## Overview

Inventory Manager is designed for fast inventory review without installing a server application. It reads inventory spreadsheets in the browser, normalizes Arabic and English Odoo export headers, removes grouped summary rows to avoid double-counting, and provides interactive KPIs, filters, charts, and export tools.

The dashboard supports two separate workflows:

1. **Normal inventory import**: import one inventory XLSX/CSV file and analyze it.
2. **Inventory comparison**: import a second XLSX file and compare it against the currently loaded inventory data.

The comparison tools are isolated from the normal import dashboard, so comparison KPIs and charts do not modify the normal dashboard KPIs, charts, filters, or exports.

---

## Main features

### Inventory import and analysis

- Import inventory files in **XLSX**, **XLS**, or **CSV** format.
- Supports Arabic and English Odoo `stock.quant` style exports.
- Automatically detects the correct worksheet when an XLSX file contains multiple sheets.
- Skips grouped Odoo summary rows to prevent duplicated totals.
- Supports files with or without a removal date column.
- Ignores optional company columns when present.
- Shows clear import progress and validation messages.

### Dashboard KPIs

Default KPI cards include:

- Products
- Out-of-stock products
- Total stock quantity
- Low-stock rows
- Negative-stock rows
- Total inventory value
- Locations

Optional KPI cards can be enabled from **More options**, including:

- Displayed rows
- Available quantity
- Reserved quantity
- Categories
- Units of measure
- Batches / serial numbers
- Rows with removal dates
- Average unit value
- Reserved value

### Interactive KPI filtering

KPI cards are clickable. Clicking a KPI filters the table to the rows represented by that metric, such as low-stock rows, negative rows, out-of-stock rows, or rows with reserved quantity.

### Search and filters

- Multi-tag search.
- Real-time search suggestions.
- Location filter.
- Category filter.
- Status filter.
- Removal-date filter.
- Unit-of-measure filter.
- Low-stock threshold control.
- Product totals mode to combine the same product across locations.
- Reset filters button.

### Location popup

When a product exists in multiple locations, the Location column can show a location-count badge. Clicking the badge opens a popup with the locations and their summed quantities. Selecting a location from the popup automatically applies it to the location filter.

### Charts

The normal dashboard includes default and optional charts. Optional charts are hidden until selected.

Default charts:

- Value by category
- Product count by category

Optional charts:

- Quantity by location
- Top products by quantity
- Stock status distribution
- Value by location
- Available vs reserved by category
- Products by category
- Unit-of-measure distribution
- Reserved quantity by location
- Low-stock products
- Unit value by product

### Table tools

- Sortable inventory table.
- Resizable columns.
- Column show/hide menu.
- Pagination.
- Export filtered results as CSV.
- Export filtered results as XLSX.

### Interface

- Arabic and English UI.
- Right-to-left and left-to-right layout support.
- Light mode and dark mode.
- Browser tab favicon and in-page app icon using `icon.png`.

---

## Comparison features

The comparison workflow lets the user load a second XLSX file and compare it with the currently imported main inventory file.

### How comparison works

1. Import the main inventory file first.
2. Open **Settings**.
3. Click **Import comparison XLSX**.
4. Select another XLSX file.
5. The app compares the second file against the current imported data.

The comparison file must be XLSX. The main inventory import can still be XLSX, XLS, or CSV.

### Comparison matching logic

Rows are matched using the inventory identity fields from the imported data, such as product, category, location, lot/serial number, removal date, and unit of measure. Matching rows are then compared by stock and value metrics.

The comparison output classifies rows as:

- **New**: exists in the comparison file but not in the current file.
- **Removed**: exists in the current file but not in the comparison file.
- **Changed**: exists in both files but has changed metrics.
- **Unchanged**: exists in both files with no detected metric change.
- **Increased**: quantity or value increased.
- **Decreased**: quantity or value decreased.

### Comparison KPI selector

Comparison KPIs have their own selector and are separate from the normal dashboard KPIs.

Only the main comparison KPIs are visible by default:

- Comparison rows
- Changes
- New rows
- Removed rows
- Quantity delta
- Value delta

Additional comparison KPIs are hidden by default and can be enabled from the **Comparison metrics** dropdown:

- Increases
- Decreases
- Unchanged rows
- Matched rows
- Quantity changed
- Available changed
- Reserved changed
- Value changed
- Quantity-only changes
- Value-only changes
- Quantity + value changes
- Value gains
- Value losses
- Quantity gain rows
- Quantity loss rows
- Unit value changed
- Affected products
- Affected locations
- Current quantity
- Comparison quantity
- Absolute quantity movement
- Quantity change percentage
- Current value
- Comparison value
- Absolute value movement
- Value change percentage
- Available delta
- Reserved delta
- Available change percentage
- Reserved change percentage
- New quantity
- Removed quantity
- New value
- Removed value
- Current imported rows
- Comparison imported rows

### Clickable comparison KPIs

Comparison KPI cards are clickable. Clicking a KPI filters the comparison table to the relevant rows. Examples:

- Click **New** to show only rows that exist in the comparison file.
- Click **Removed** to show rows missing from the comparison file.
- Click **Quantity changed** to show rows where on-hand quantity changed.
- Click **Value gains** to show rows where value increased.
- Click **Affected locations** to inspect changed locations.

The **Show all comparison** button clears the active comparison view and returns to the full comparison table.

### Comparison charts

Comparison charts have their own selector and are all hidden by default. The user can choose exactly which comparison charts to show.

Available comparison charts include:

- Comparison row status
- Quantity movement by status
- Value movement by status
- Top quantity gains
- Top quantity losses
- Top value gains
- Top value losses
- Current vs comparison quantity by category
- Current vs comparison value by category
- Changed rows by location
- Quantity movement by location
- Value movement by location
- Changed rows by category
- Quantity movement by category
- Value movement by category

Supported comparison chart styles include:

- Doughnut charts
- Vertical bar charts
- Horizontal bar charts
- Grouped bar charts
- Line charts

Comparison charts react to the active comparison view where applicable.

### Comparison exports

The comparison result can be exported separately from the normal inventory table.

Available comparison exports:

- Export comparison as CSV
- Export comparison as XLSX

The comparison export includes detailed fields such as current quantity, comparison quantity, deltas, percentage changes, value changes, unit value changes, category, location, and source row counts.

---

## Supported import format

The importer supports Arabic and English Odoo-style inventory headers.

| Internal field | English  | Required |
|---|---|---|
| Product | Product | Yes |
| Product category | Product Category | Yes |
| Location | Location | Yes |
| Lot / serial number | Lot/Serial Number | Yes |
| Removal date | Removal Date | No |
| On-hand quantity | Inventoried Quantity / On Hand Quantity | Yes |
| Available quantity | Available Quantity | Yes |
| Unit of measure | Unit of Measure | Yes |
| Value | Value | Yes |

Optional company columns are accepted and ignored during import.

Supported company headers include:

- `Company`
- `Company Name`


---

## How to use

### Local usage

1. Download or clone the project.
2. Open `index.html` in a modern browser.
3. Click **Settings**.
4. Click **Import file**.
5. Select the inventory XLSX/CSV file.
6. Use KPIs, filters, charts, and the table to analyze the imported data.
7. Export filtered results using **Export CSV** or **Export XLSX**.

### Comparison usage

1. Import the main inventory file first.
2. Open **Settings**.
3. Click **Import comparison XLSX**.
4. Select the second inventory XLSX file.
5. Review the comparison KPIs, table, and optional comparison charts.
6. Click comparison KPIs to inspect specific row groups.
7. Export the comparison using **Export comparison CSV** or **Export comparison XLSX**.

---

### Notes for hosted usage

- After updating the app, hard-refresh the browser if an old version still appears.
- The project already uses versioned CSS/JS references in `index.html` to reduce cache issues on GitHub Pages.
- The comparison import control is implemented as a real file-picker label so it works reliably on hosted pages.

---

## Project structure

```text
.
├── index.html   # Main HTML layout and UI structure
├── styles.css   # Dashboard styling, themes, responsive layout, and controls
├── app.js       # Import parsing, normalization, KPIs, charts, filters, exports, and comparison logic
├── icon.png     # Browser favicon and in-page app icon
└── README.md    # Project documentation
```

---

## Browser support

Use a current desktop browser for best results:

- Google Chrome
- Microsoft Edge
- Mozilla Firefox
- Brave

XLSX support depends on modern browser APIs for reading compressed spreadsheet files. If a browser cannot open compressed XLSX files, use a current browser version or import CSV for the normal inventory workflow.

The comparison importer requires XLSX.

---

## Data privacy

Inventory Manager runs entirely in the browser. Imported files are processed locally by the browser session. The project does not include a backend, database, tracking script, or network upload function.

When hosted on GitHub Pages, the static files are served by GitHub Pages, but the imported inventory spreadsheets are still processed locally in the user's browser.

---

## Troubleshooting

### The import button does not open the file picker on GitHub Pages

Hard-refresh the page, then make sure the deployed repository contains the latest `index.html`, `styles.css`, and `app.js`. Browser caching can keep older files active after deployment.

### The comparison button does nothing

Import the main inventory file first. The comparison importer is blocked until the current inventory dataset exists.

### The comparison file is rejected

The comparison file must be XLSX. CSV is supported for the normal import only.

### Imported totals look duplicated

Make sure the file is an Odoo-style export or follows the supported headers. The app skips recognized grouped summary rows, but unusual custom summary rows may still need to be removed from the spreadsheet before import.

### Some charts are not visible

Normal optional charts and all comparison charts are hidden by default. Open the chart selector and enable the charts you want to display.

### Some KPIs are not visible

Normal optional KPIs are controlled from **More options**. Comparison KPIs are controlled from the comparison KPI selector.

---

## Version notes

This release includes:

- GitHub Pages-safe comparison import control.
- Pointer cursor behavior for the comparison import control.
- Comparison-only KPI selector.
- Comparison-only chart selector.
- Hidden-by-default advanced comparison KPIs.
- Hidden-by-default comparison charts.
- Clickable comparison KPI filtering.
- Expanded comparison metrics, deltas, percentages, and source row counts.
- Separate comparison CSV/XLSX export.
- Preserved normal import dashboard behavior.
