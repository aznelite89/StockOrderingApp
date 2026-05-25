"""Entry point — registers the multi-page navigation.

Render's auto-detect still finds this file as the Streamlit entry, so
deployment doesn't need a start-command change. Page labels are decoupled
from filenames here via `st.Page(..., title=...)`.
"""

import streamlit as st


stock_ordering_sheet = st.Page(
    "views/stock_ordering_sheet.py",
    title="Stock Ordering Sheet",
    default=True,
    url_path="stock-ordering-sheet",
)
invoice_to_excel_with_rrp = st.Page(
    "views/invoice_to_excel_with_rrp.py",
    title="Invoice to Excel with RRP",
    url_path="invoice-to-excel-with-rrp",
)

navigation = st.navigation([stock_ordering_sheet, invoice_to_excel_with_rrp])
navigation.run()
