# Inventory Manager

A fully browser-based inventory dashboard for XLSX / CSV stock files, including grouped Odoo `stock.quant` exports.

## What changed in this version

- Fixed the startup issue caused by the optional KPI renderer.
- The dashboard now loads normally again after adding the chart selector.
- Removed **Removal date from** and **Removal date to** from the **More options** area.
- Kept the **Select charts** dropdown:
  - `Value by category` stays visible by default.
  - `Quantity distribution by category` stays visible by default.
  - All other charts start disabled and can be multi-selected manually.
- Kept compact money labels on money charts, for example `10M`, `1B`, and `155.5M`.
- Kept expandable chart buttons, clickable KPI cards, optional KPIs, export, search, filters, light/dark mode, Arabic/English support, and grouped Odoo import support.

## How to use

1. Open `index.html` in a modern browser.
2. Click **Settings**.
3. Use **Import file** to import the XLSX / CSV inventory file.
4. Use **Select charts** to activate any extra charts you want on the dashboard.
5. Use search, filters, chart expansion, table sorting, and export buttons.
6. Click **Total by product** to see one combined row per product across all locations.
7. Open **More options** to enable extra KPI cards.
8. Click any KPI card to inspect the rows referenced by that KPI.
9. Use **Export XLSX** or **Export CSV** to export the currently filtered result.

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

### Table columns

- Open **Columns / الأعمدة** above the table to hide or show any table column.
- Drag the edge of any column header to resize it.
- Double-click a column resize edge to reset that column width.
- Column visibility and custom widths are saved in the browser.
