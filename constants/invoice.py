"""Shared constants for the invoice-to-Excel tool."""

OUTPUT_SHEET_NAME = "Searay Pty Ltd_Invoice as Excel"
RULES_SHEET_NAME = "Pricing Rules"

FREIGHT_CODE = "FREIGHT"
BACKORDER_HEADING = "Items on backorder:"


class HeaderField:
    CUSTOMER_NAME = "Customer Name"
    CUSTOMER_CODE = "Customer Code"
    INVOICE_DATE = "Invoice date"
    ORDER_REFERENCE = "Order Reference"


class LineColumn:
    LINE_NO = "Line No"
    CODE = "Code"
    DESCRIPTION = "Description"
    QTY = "Qty"
    PRICE = "Price"
    TOTAL_BEFORE_DISCOUNT = "Total Before Discount"
    DISC = "Disc"
    PRICE_AFTER_DISCOUNT = "Price pg / pp After Discount"
    TOTAL_AFTER_DISCOUNT = "Total After Discount (Ex-GST)"
    RRP = "RRP"


LINE_COLUMN_ORDER = [
    LineColumn.LINE_NO,
    LineColumn.CODE,
    LineColumn.DESCRIPTION,
    LineColumn.QTY,
    LineColumn.PRICE,
    LineColumn.TOTAL_BEFORE_DISCOUNT,
    LineColumn.DISC,
    LineColumn.PRICE_AFTER_DISCOUNT,
    LineColumn.TOTAL_AFTER_DISCOUNT,
    LineColumn.RRP,
]

# Layout: header fields occupy rows 1..4 (cells A1..B4); blank row 5;
# table header row is row 6 (Excel 1-indexed) which is xlsxwriter row 5.
HEADER_ROW_COUNT = 4
TABLE_HEADER_ROW_EXCEL = 6  # 1-indexed
FIRST_DATA_ROW_EXCEL = 7    # 1-indexed
FIRST_DATA_ROW_INDEX = 6    # 0-indexed for xlsxwriter

# Column letters in the output sheet (A..J), used to build live formulas
# that match the sample workbook (sheet1.xml in
# "WITH RRP RULES - Searay Pty Ltd_Invoice as Excel_Speirs (web#2971).xlsx").
class OutputColumn:
    LINE_NO = "A"
    CODE = "B"
    DESCRIPTION = "C"
    QTY = "D"
    PRICE = "E"
    TOTAL_BEFORE_DISCOUNT = "F"
    DISC = "G"
    PRICE_AFTER_DISCOUNT = "H"
    TOTAL_AFTER_DISCOUNT = "I"
    RRP = "J"


# Exact IFS formula copied from sample cell J7. {r} is the 1-indexed row.
RRP_FORMULA_TEMPLATE = (
    "=IFS("
    "I{r}<380,CEILING(I{r}*2.8,10)-1,"
    "I{r}<=1400,IF(I{r}*2.5<300,CEILING(I{r}*2.5,10)-1,CEILING(I{r}*2.5,50)-1),"
    "I{r}>1400,\"Check RRP\""
    ")"
)

TOTAL_BEFORE_DISCOUNT_FORMULA = "=E{r}*D{r}"
PRICE_AFTER_DISCOUNT_FORMULA = "=E{r}*(1-G{r})"
SUM_TOTAL_BEFORE_DISCOUNT_FORMULA = "=SUM(F{first}:F{last})"

CHECK_RRP_SENTINEL = "Check RRP"

PDF_DATE_FORMATS = ("%d/%m/%Y", "%d-%m-%Y")
