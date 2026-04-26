# Inventory Manager

A fully browser-based inventory dashboard for XLSX / CSV stock files, including grouped Odoo `stock.quant` exports.

## What changed in this version

- Money-based charts now use compact rounded labels, such as `10M`, `1B`, or `155.5M`, instead of long full values. This applies to chart axes, bar labels, and chart tooltips.

- Added a **Total by product** filter/toggle.
  - When enabled, the table and charts combine every product across all locations.
  - On-hand quantity, available quantity, reserved quantity, and value are summed per product.
  - Category, location, batch/serial, unit, and removal-date fields are summarized in the row instead of duplicating the product across locations.
- Removed chart cards related to:
  - Batch / serial number counts
  - Removal-date timeline
- Kept the expandable chart button on every remaining chart.
- Kept clickable KPI cards, optional KPIs, export, search, filters, light/dark mode, Arabic/English support, and grouped Odoo import support.

## How to use

1. Open `index.html` in a modern browser.
2. Click **Settings**.
3. Use **Import file** to import the XLSX / CSV inventory file.
4. Use search, filters, chart expansion, table sorting, and export buttons.
5. Click **Total by product** to see one combined row per product across all locations.
6. Open **More options** to enable extra KPI cards.
7. Click any KPI card to inspect the rows referenced by that KPI.
8. Use **Export XLSX** or **Export CSV** to export the currently filtered result.

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
