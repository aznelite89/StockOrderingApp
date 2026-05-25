# Changelog

## 2026-05-25

### Added
- Invoice → Excel with RRP tool — drag-and-drop a Searay PDF tax invoice + Pricing Rules workbook to produce a CSV and formula-bearing XLSX with the RRP column populated; freight rows and backorder items are excluded.
- `run_invoice.sh` launcher and `pdfplumber` / `openpyxl` dependencies.
- `utils/pdf_invoice.py`, `utils/rrp.py`, `utils/output.py`, `constants/invoice.py` supporting modules.

### Changed
- Reorganised as a multi-page Streamlit app via `st.navigation()`. `app.py` is now a thin navigation entry; the existing UNL Order Sheet generator lives at `views/stock_ordering_sheet.py` and the new invoice tool at `views/invoice_to_excel_with_rrp.py`.
- Sidebar labels: home page is "Stock Ordering Sheet" (was "app"), and the invoice page is "Invoice to Excel with RRP" (was "Invoice Converter"). URLs are `/stock-ordering-sheet` and `/invoice-to-excel-with-rrp`.
