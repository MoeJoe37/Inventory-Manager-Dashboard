# Inventory Manager

A fully browser-based inventory dashboard for XLSX / CSV stock files, including grouped Odoo `stock.quant` exports.

## What changed in this version

- Moved **Import file**, **XLSX template**, and **CSV template** into the **Settings** menu.
- Removed the topbar import/template buttons so the dashboard takes less vertical and horizontal space.
- Kept the empty first-run screen with quick import/template buttons.
- Made the import success message a temporary floating toast instead of a page section that pushes the dashboard down.
- Improved chart rendering in light mode by forcing canvas redraws after theme changes and applying theme-aware canvas backgrounds/colors.
- Kept Light/Dark mode and Arabic/English toggles in Settings.
- Kept compatibility with the exact attached Odoo-style grouped XLSX format.

## How to use

1. Open `index.html` in a modern browser.
2. Click **Settings**.
3. Use **Import file** to import the XLSX / CSV inventory file.
4. Use search, filters, chart clicks, table sorting, and export buttons.
5. Use **Settings** to switch theme/language or download XLSX/CSV templates.

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

## v6 fix

- Fixed the import/start card so it disappears completely after a successful import and does not reserve dashboard space.
