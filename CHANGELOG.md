# Changelog

## 2026-09-02

### Added
- Rows for products flagged not purchasable in Unleashed (`Is Purchasable = No` in the Product List export) are highlighted red on `UNL_Order_Sheet` and every `<SUPPLIER>_ALL` tab, with a new `UNL Purchasable` column (Z) and a `Calculation_Logic` entry.
- Preview shows a count of not-purchasable products and a red/yellow legend.
- `constants/order_sheet.py` for Product List column names, YES/NO values, and highlight colours.

### Changed
- Version bumped to v4.9.
- Red (not purchasable) takes precedence over yellow (special order) when a row qualifies for both.
- Sample Data Format tab 4 now includes the `Is Purchasable` column.

## 2026-07-16

### Fixed
- `ProcessOrderMacro` now pulls the order quantity (and include filter and Total Weight) from column `V` (Searay Order) instead of `T` (Need To Order), so the supplier PO form reflects the hand-entered "Searay Order" quantities rather than the system suggestion.

## 2026-07-15

### Fixed
- Uploading a Windows-1252/Excel CSV export no longer crashes with `UnicodeDecodeError` (byte 0x93); CSV reads now try utf-8-sig, then cp1252, then latin-1 via `utils/csv_loader.read_csv_bytes`.

## 2026-06-04

### Fixed
- `ProcessOrderMacro` now reads the order quantity from column `T` (Need To Order) and notes from `W` (Comments); it previously read `U` (Base Unit), so "Qty (pcs)" showed "Each" and Total Weight came out blank. Filter now requires a numeric `T > 0`.

## 2026-06-02

### Fixed
- Last PO Date/Qty now parse Unleashed dates as day-first (DD/MM/YYYY), so POs dated after the 12th are no longer dropped and the most-recent transaction is reported.

## 2026-05-27

### Fixed
- Stock Ordering Sheet hanging on Render with no preview shown — the per-row bin-grouping loop in `views/stock_ordering_sheet.py` was O(n²) and could exhaust Render free tier's 0.1 CPU / 512MB before reaching the dataframe preview.

### Changed
- Vectorised the bin-grouping in `views/stock_ordering_sheet.py` to an O(n log n) merge + stable sort while preserving bin grouping and within-bin need-desc order.
- Cached CSV loading and the full data-processing pipeline via `@st.cache_data` so editing the supplier text input no longer re-reads 9MB of CSVs.
- Excel workbook generation is now gated behind a "Build Excel for download" button instead of running on every script rerun.
- Hoisted `workbook.add_format()` calls out of the per-row write loops (main sheet and per-supplier ALL sheet).
- Wrapped data loading in `st.spinner` so slow hosts show progress instead of a blank page below the upload widgets.
- Cached the empty 6-tab sample workbook so it isn't rebuilt on every rerun.

## 2026-05-25

### Added
- Invoice → Excel with RRP tool — drag-and-drop a Searay PDF tax invoice + Pricing Rules workbook to produce a CSV and formula-bearing XLSX with the RRP column populated; freight rows and backorder items are excluded.
- `run_invoice.sh` launcher and `pdfplumber` / `openpyxl` dependencies.
- `utils/pdf_invoice.py`, `utils/rrp.py`, `utils/output.py`, `constants/invoice.py` supporting modules.

### Changed
- Reorganised as a multi-page Streamlit app via `st.navigation()`. `app.py` is now a thin navigation entry; the existing UNL Order Sheet generator lives at `views/stock_ordering_sheet.py` and the new invoice tool at `views/invoice_to_excel_with_rrp.py`.
- Sidebar labels: home page is "Stock Ordering Sheet" (was "app"), and the invoice page is "Invoice to Excel with RRP" (was "Invoice Converter"). URLs are `/stock-ordering-sheet` and `/invoice-to-excel-with-rrp`.
