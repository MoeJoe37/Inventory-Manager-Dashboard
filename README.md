# Inventory Manager

A fully browser-based inventory dashboard for XLSX / CSV stock files, including grouped Odoo `stock.quant` exports.

## What changed in this version

- Made the dashboard more compact and responsive so it fits the page better at normal browser zoom.
- Added a **See more charts** button under the main chart area.
- Added detailed extra charts for value by location, available vs reserved by category, product count by category, unit-of-measure distribution, reserved quantity by location, lowest available products, batch count by category, and unit value by product.
- Kept **Import file**, **XLSX template**, and **CSV template** inside the **Settings** menu.
- Kept the import/start card hidden after a successful import so it does not take dashboard space.
- Kept chart rendering working in both light mode and dark mode.
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
