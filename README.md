# Inventory Manager

A fully browser-based inventory dashboard for XLSX / CSV stock files, including grouped Odoo `stock.quant` exports.

## What changed in this version

- Made every KPI card clickable, including main KPIs and optional KPIs.
- Clicking a KPI opens its referenced inventory rows in the same searchable/filterable/exportable table.
- The selected KPI view is highlighted and can be cleared with **Clear view**.
- Kept the visible **Expand chart** button. Every chart card has a large modal view.
- Added **Optional KPIs** under **More options** so users can choose extra KPI cards without cluttering the dashboard.
- Optional KPI cards include:
  - Visible rows
  - Total available quantity
  - Total reserved quantity
  - Category count
  - Unit-of-measure count
  - Batch / serial count
  - Removal-dated rows
  - Average unit value
  - Reserved value
- Optional KPI selections are saved in browser local storage with the language/theme preferences.
- Kept negative-stock KPI, clickable stock KPI cards, export, light/dark mode, Arabic/English support, and grouped Odoo import support.

## How to use

1. Open `index.html` in a modern browser.
2. Click **Settings**.
3. Use **Import file** to import the XLSX / CSV inventory file.
4. Use search, filters, chart expansion, table sorting, and export buttons.
5. Open **More options** to enable extra KPI cards.
6. Click any KPI card to inspect the rows referenced by that KPI.
7. Use search, filters, and **Export XLSX** while inside a KPI view to export the currently shown result.

## Required import headers

The importer expects these Arabic headers because they match the Odoo export format:

- المنتج
- فئة المنتج
- الموقع
- رقم الدفعة/الرقم التسلسلي
- تاريخ الإزالة
- الكمية في المخزون
- الكمية المتوفرة
- وحدة القياس
- القيمة

Grouped Odoo summary rows are automatically skipped so dashboard totals do not double-count inventory.

## Files

- `index.html`
- `styles.css`
- `app.js`
- `README.md`

No backend, database, build step, or external JavaScript library is required.
