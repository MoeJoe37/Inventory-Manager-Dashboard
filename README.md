# Inventory Manager

A fully browser-based inventory dashboard for XLSX / CSV stock files, including grouped Odoo `stock.quant` exports.

## What changed in this version

- Added support for the new English Odoo `stock.quant` export format.
- Added support for the Arabic Odoo `stock.quant` export that includes the extra company column.
- The optional `Company` / `الشركة` / `Company Name` field is ignored during import and is not required.
- `Removal date` remains supported when present, but the importer can now load files that do not include it.
- Kept compact money labels on money charts, expandable chart buttons, clickable KPI cards, optional KPIs, export, search, filters, light/dark mode, Arabic/English UI support, column hiding/resizing, and grouped Odoo summary-row skipping.

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

## Supported import headers

The importer supports these Arabic and English Odoo `stock.quant` headers:

| Internal field | Arabic header | English header | Required |
|---|---|---|---|
| Product | المنتج | Product | Yes |
| Product category | فئة المنتج | Product Category | Yes |
| Location | الموقع | Location | Yes |
| Lot / serial number | رقم الدفعة/الرقم التسلسلي | Lot/Serial Number | Yes |
| Removal date | تاريخ الإزالة | Removal Date | No |
| On-hand quantity | الكمية في المخزون | Inventoried Quantity / On Hand Quantity | Yes |
| Available quantity | الكمية المتوفرة | Available Quantity | Yes |
| Unit of measure | وحدة القياس | Unit of Measure | Yes |
| Value | القيمة | Value | Yes |

The optional company column is ignored if present. Supported company headers include `Company`, `Company Name`, `الشركة`, and `اسم الشركة`.

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
