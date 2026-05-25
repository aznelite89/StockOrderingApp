"""Streamlit app: convert a Searay PDF tax invoice → CSV + XLSX with RRP.

Companion to app.py (the UNL Order Sheet Generator). This app handles the
"Invoice to Excel with RRP Pricing Rules" workflow for Speirs Jewellers
(and other clients with their own Pricing Rules tab).
"""

from __future__ import annotations

import streamlit as st

from constants.invoice import (
    CHECK_RRP_SENTINEL,
    HeaderField,
    LineColumn,
)
from utils.output import build_csv, build_xlsx
from utils.pdf_invoice import ParsedInvoice, parse_invoice_pdf
from utils.rrp import compute_rrp

st.set_page_config(page_title="Invoice → Excel with RRP", layout="wide")
st.title("Invoice → Excel with RRP")
st.caption(
    "Drag-and-drop a Searay tax-invoice PDF, supply the Pricing Rules workbook, "
    "and download the customer-ready CSV + XLSX with the RRP column filled in."
)

with st.expander("How this works", expanded=False):
    st.markdown(
        """
        1. **PDF invoice** — the Searay tax invoice exported as text-based PDF.
        2. **Pricing Rules workbook (.xlsx)** — must contain a `Pricing Rules`
           tab. The tab is copied through to the output workbook so the
           customer can see the formula source. The actual markup/rounding
           logic used today is hard-coded to match the Speirs sample:
           `IFS(I<380, CEILING(I*2.8,10)-1, I<=1400, IF(I*2.5<300, ..., ...),
           I>1400, "Check RRP")`.
        3. **Customer Code** — the short customer ID (e.g. `SJ003`). It is
           not present on the PDF; enter it per export.

        Freight rows and any items under "Items on backorder:" are excluded
        from the output, matching the Speirs sample.
        """
    )

col_left, col_right = st.columns(2)
with col_left:
    pdf_file = st.file_uploader("1. PDF invoice", type=["pdf"])
with col_right:
    rules_file = st.file_uploader(
        "2. Pricing Rules workbook (.xlsx)", type=["xlsx"]
    )

customer_code = st.text_input("3. Customer Code", value="", placeholder="e.g. SJ003")

st.divider()

if not (pdf_file and rules_file and customer_code):
    st.info("Upload both files and enter a Customer Code to generate the export.")
    st.stop()

with st.spinner("Parsing PDF…"):
    parsed: ParsedInvoice = parse_invoice_pdf(pdf_file)

if parsed.line_items.empty:
    st.error(
        "No line items were extracted from the PDF. Check that this is a "
        "text-based Searay tax invoice (not a scanned image)."
    )
    st.stop()

df = parsed.line_items.copy()
df[LineColumn.RRP] = df[LineColumn.TOTAL_AFTER_DISCOUNT].apply(compute_rrp)

st.subheader("Preview")
meta_cols = st.columns(4)
meta_cols[0].metric("Customer", parsed.header[HeaderField.CUSTOMER_NAME] or "—")
meta_cols[1].metric("Order Reference", parsed.header[HeaderField.ORDER_REFERENCE] or "—")
meta_cols[2].metric(
    "Invoice Date",
    parsed.header[HeaderField.INVOICE_DATE].strftime("%d/%m/%Y")
    if parsed.header[HeaderField.INVOICE_DATE] else "—",
)
meta_cols[3].metric("Line items", len(df))

if parsed.excluded_freight or parsed.excluded_backorder:
    st.caption(
        f"Excluded from output: {parsed.excluded_freight} freight row(s), "
        f"{parsed.excluded_backorder} backorder row(s)."
    )

if parsed.warnings:
    for w in parsed.warnings:
        st.warning(w)

check_rrp_rows = df[df[LineColumn.RRP] == CHECK_RRP_SENTINEL]
if not check_rrp_rows.empty:
    codes = ", ".join(check_rrp_rows[LineColumn.CODE].tolist())
    st.warning(
        f"{len(check_rrp_rows)} row(s) need manual RRP review (Total After "
        f"Discount > $1,400): {codes}"
    )

st.dataframe(df, use_container_width=True, hide_index=True)

rules_bytes = rules_file.getvalue()
csv_bytes = build_csv(parsed, df, customer_code)
xlsx_bytes = build_xlsx(parsed, df, customer_code, rules_bytes,
                        on_warning=st.warning)

inv_ref = parsed.header[HeaderField.ORDER_REFERENCE] or "invoice"
base_name = f"{parsed.header[HeaderField.CUSTOMER_NAME] or 'Customer'}_Invoice_{inv_ref}"
base_name = base_name.replace("/", "-").replace("\\", "-").replace(" ", "_")

st.subheader("Downloads")
d_cols = st.columns(2)
with d_cols[0]:
    st.download_button(
        label="Download CSV",
        data=csv_bytes,
        file_name=f"{base_name}.csv",
        mime="text/csv",
    )
with d_cols[1]:
    st.download_button(
        label="Download XLSX (with formulas)",
        data=xlsx_bytes,
        file_name=f"{base_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
