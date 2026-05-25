# Changelog

## 2026-05-25

### Added
- Invoice → Excel with RRP tool (`invoice_app.py`) — drag-and-drop a Searay PDF tax invoice + Pricing Rules workbook to produce a CSV and formula-bearing XLSX with the RRP column populated; freight rows and backorder items are excluded.
- `run_invoice.sh` launcher and `pdfplumber` / `openpyxl` dependencies.
- `utils/pdf_invoice.py`, `utils/rrp.py`, `utils/output.py`, `constants/invoice.py` supporting modules.
