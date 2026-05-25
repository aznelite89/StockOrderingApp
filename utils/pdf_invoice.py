"""Extract header fields and line items from Searay tax-invoice PDFs.

The PDF text-extraction order does not match the visible left-to-right column
order: a description that wraps onto two lines causes the extractor to dump
the description before/after the price columns. We work around that by
extracting words with x/y coordinates, detecting column anchors from each
page's table header, and assigning words to columns by nearest x-anchor.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import IO, Optional

import pandas as pd
import pdfplumber

from constants.invoice import (
    BACKORDER_HEADING,
    FREIGHT_CODE,
    HeaderField,
    LINE_COLUMN_ORDER,
    LineColumn,
)

# Column tolerance for assigning a word to a header anchor.
COLUMN_TOLERANCE = 18.0  # px

# Vertical tolerance for grouping words into the same visual row.
ROW_Y_TOLERANCE = 3.0

# Labels expected in the line-item table header, in left-to-right visible
# order. We map each PDF header label to our internal column name.
HEADER_LABEL_TO_COLUMN = {
    "Ln":          LineColumn.LINE_NO,
    "Code":        LineColumn.CODE,
    "Description": LineColumn.DESCRIPTION,
    "Units":       "UNITS",
    "Qty":         LineColumn.QTY,
    "Price":       LineColumn.PRICE,
    "Disc":        LineColumn.DISC,
}
# The "Price pg/pp" header is stacked vertically ("Price pg /\npp") and
# "Total Ex-GST" is stacked ("Total ExGST"). We detect them via the joined
# "ExGSTPrice" token that appears on the header row.
PRICE_PG_PP_LABEL = "pp"          # on row below the main header
TOTAL_EXGST_LABEL = "ExGSTPrice"  # joined header text spans the column


@dataclass
class ParsedInvoice:
    header: dict
    line_items: pd.DataFrame
    excluded_freight: int = 0
    excluded_backorder: int = 0
    parsed_sub_total: Optional[float] = None
    pdf_sub_total: Optional[float] = None
    warnings: list = field(default_factory=list)


def parse_invoice_pdf(file: IO) -> ParsedInvoice:
    """Parse a Searay tax-invoice PDF.

    Returns a ParsedInvoice with extracted header, included line items, and
    counts of excluded freight/backorder lines.
    """
    with pdfplumber.open(file) as pdf:
        all_pages_words = [_extract_words(page) for page in pdf.pages]
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    header = _extract_header(full_text)

    line_items: list[dict] = []
    excluded_freight = 0
    excluded_backorder = 0

    # Anchor positions from the first page are reused for later pages that
    # don't have their own table header (Searay's PDFs only print the
    # header once on page 1).
    carried_anchors: Optional[dict] = None
    for words in all_pages_words:
        page_rows, freight, backorder, carried_anchors = (
            _extract_line_items_from_page(words, carried_anchors)
        )
        line_items.extend(page_rows)
        excluded_freight += freight
        excluded_backorder += backorder

    parsed_sub_total = sum(
        (row.get(LineColumn.TOTAL_AFTER_DISCOUNT) or 0) for row in line_items
    )
    pdf_sub_total = _extract_sub_total(full_text)

    df = pd.DataFrame(line_items, columns=LINE_COLUMN_ORDER)
    df[LineColumn.TOTAL_BEFORE_DISCOUNT] = (
        df[LineColumn.QTY] * df[LineColumn.PRICE]
    )

    warnings: list = []
    if pdf_sub_total is not None:
        diff = abs(parsed_sub_total - pdf_sub_total)
        # The PDF sub-total includes freight; allow up to ~$250 of freight
        # before flagging. Anything larger suggests the parser dropped rows.
        if diff > 250:
            warnings.append(
                f"Parsed line-item total ${parsed_sub_total:,.2f} differs from "
                f"PDF Sub Total ${pdf_sub_total:,.2f} by ${diff:,.2f}. "
                "Verify all rows were extracted correctly."
            )

    return ParsedInvoice(
        header=header,
        line_items=df,
        excluded_freight=excluded_freight,
        excluded_backorder=excluded_backorder,
        parsed_sub_total=parsed_sub_total,
        pdf_sub_total=pdf_sub_total,
        warnings=warnings,
    )


# Word extraction & row grouping ------------------------------------------

def _extract_words(page) -> list[dict]:
    return page.extract_words(use_text_flow=False, keep_blank_chars=False)


def _group_into_rows(words: list[dict]) -> list[list[dict]]:
    rows: list[list[dict]] = []
    for w in sorted(words, key=lambda x: (x["top"], x["x0"])):
        if rows and abs(w["top"] - rows[-1][0]["top"]) <= ROW_Y_TOLERANCE:
            rows[-1].append(w)
        else:
            rows.append([w])
    return rows


# Column-anchor detection -------------------------------------------------

def _detect_column_anchors(rows: list[list[dict]]) -> Optional[dict]:
    """Find x-anchors of each line-item column by inspecting the table header.

    Returns a dict mapping column-name → x0, plus '_header_top' = y of header
    row and '_desc_x_max' = x of the right edge of the Description column
    (just before Units). Returns None if the header row isn't found.
    """
    # The backorder section also has a row of "Ln Code Description Units Qty"
    # but it lacks Price/Disc. Require Price too so we only pick the
    # line-item table header.
    header_row = None
    for row in rows:
        texts = {w["text"] for w in row}
        if {"Ln", "Code", "Description", "Units", "Price", "Disc"} <= texts:
            header_row = row
            break
    if header_row is None:
        return None

    anchors: dict = {}
    for w in header_row:
        col = HEADER_LABEL_TO_COLUMN.get(w["text"])
        if col is not None and col not in anchors:
            anchors[col] = w["x0"]

    # Detect Price pg/pp and Total Ex-GST positions. They appear as a
    # stacked header just above the main header line, with x0 anchors that
    # we can read from the "pp" word (Price pg/pp) and from the column
    # immediately right of Disc. We approximate Total Ex-GST as the
    # right-most numeric column anchor seen in the first data row.
    for row in rows:
        for w in row:
            if w["text"] == PRICE_PG_PP_LABEL and abs(w["top"] - header_row[0]["top"]) < 15:
                # "pp" sits below "Price pg /" — its x0 is roughly the column anchor.
                anchors.setdefault(LineColumn.PRICE_AFTER_DISCOUNT, w["x0"] - 8)

    # Total Ex-GST column: find a value-row word at the right-most x. We use
    # the first row below the header that has a number-looking value near
    # the right edge.
    if LineColumn.TOTAL_AFTER_DISCOUNT not in anchors:
        anchors[LineColumn.TOTAL_AFTER_DISCOUNT] = _guess_rightmost_money_x(
            rows, header_row[0]["top"]
        )
    # If Price pg/pp still missing, place it midway between Disc and Total Ex-GST.
    if (LineColumn.PRICE_AFTER_DISCOUNT not in anchors
            and LineColumn.DISC in anchors
            and LineColumn.TOTAL_AFTER_DISCOUNT in anchors):
        anchors[LineColumn.PRICE_AFTER_DISCOUNT] = (
            (anchors[LineColumn.DISC] + anchors[LineColumn.TOTAL_AFTER_DISCOUNT]) / 2
        )

    anchors["_header_top"] = header_row[0]["top"]
    # Description right edge is just before Units anchor.
    if "UNITS" in anchors:
        anchors["_desc_x_max"] = anchors["UNITS"] - 2
    else:
        anchors["_desc_x_max"] = 330.0
    return anchors


def _guess_rightmost_money_x(rows, header_top) -> float:
    """Find the x0 of the right-most numeric word in the first data row."""
    for row in rows:
        if row[0]["top"] <= header_top + 5:
            continue
        money = [w for w in row if re.fullmatch(r"[\d,]+\.\d{2}", w["text"])]
        if money:
            return max(w["x0"] for w in money)
    return 528.0  # sensible fallback


def _column_for(word: dict, anchors: dict) -> Optional[str]:
    x0 = word["x0"]
    x1 = word.get("x1", x0)
    desc_x_min = anchors.get(LineColumn.DESCRIPTION, 135.0)
    desc_x_max = anchors.get("_desc_x_max", 328.0)
    # Both edges must sit inside the description range; otherwise the word
    # straddles into the Units column (e.g. "16.5Each" joined token).
    if desc_x_min - 4 <= x0 and x1 <= desc_x_max:
        return LineColumn.DESCRIPTION
    best_col, best_dist = None, COLUMN_TOLERANCE
    for col, anchor in anchors.items():
        if col.startswith("_") or col == LineColumn.DESCRIPTION:
            continue
        dist = abs(x0 - anchor)
        if dist <= best_dist:
            best_col, best_dist = col, dist
    return best_col


# Line item extraction ----------------------------------------------------

def _extract_line_items_from_page(
    words: list[dict],
    carried_anchors: Optional[dict] = None,
) -> tuple[list[dict], int, int, Optional[dict]]:
    rows = _group_into_rows(words)
    anchors = _detect_column_anchors(rows)
    if anchors is None:
        if carried_anchors is None:
            return [], 0, 0, None
        # Reuse last page's anchors; treat the entire page as line-item region.
        anchors = dict(carried_anchors)
        anchors["_header_top"] = -1.0  # accept rows from top of page
    header_y = anchors["_header_top"]
    backorder_y = _find_backorder_y(rows)
    subtotal_y = _find_subtotal_y(rows)

    # End of line-item region on this page: above Sub Total (and we still
    # parse backorder rows below it to count them, until we hit page end).
    end_y = subtotal_y if subtotal_y is not None else float("inf")

    # Identify anchor rows: rows that contain a Ln number AND a Code value.
    # When the Code wraps onto an adjacent y-band (e.g. "9KCH375AQ19.5C"
    # printed above "M"), we look up to ROW_Y_TOLERANCE * 3 pixels
    # above/below the Ln row to find the Code-column word(s).
    anchor_rows: list[tuple[float, dict]] = []
    for idx, row in enumerate(rows):
        if row[0]["top"] <= header_y + 2:
            continue
        if row[0]["top"] >= end_y:
            if backorder_y is None or row[0]["top"] < backorder_y:
                continue
        parsed = _try_parse_anchor_row(row, anchors, rows)
        if parsed is not None:
            anchor_rows.append((row[0]["top"], parsed))


    if not anchor_rows:
        return [], 0, 0, anchors

    # Compute y-bands: each anchor's band runs from midpoint with previous
    # anchor down to midpoint with next anchor. The very first band extends
    # up to the header (so wrapped descriptions ABOVE the code line are
    # captured), and the last band extends to "Sub Total" (or page end).
    anchor_tops = [t for t, _ in anchor_rows]
    # Typical row height is the median gap between adjacent anchors.
    if len(anchor_tops) >= 2:
        gaps = [anchor_tops[i + 1] - anchor_tops[i] for i in range(len(anchor_tops) - 1)]
        gaps.sort()
        median_gap = gaps[len(gaps) // 2]
    else:
        median_gap = 20.0
    # The last band's lower bound is capped at top + 1.5 × median row height
    # so we don't pull in footer text far below.
    bands: list[tuple[float, float]] = []
    for i, top in enumerate(anchor_tops):
        upper = (anchor_tops[i - 1] + top) / 2 if i > 0 else header_y + 1
        if i + 1 < len(anchor_tops):
            lower = (top + anchor_tops[i + 1]) / 2
        else:
            lower = min(end_y, top + 1.5 * median_gap)
        bands.append((upper, lower))

    # Assign description words within each band. Use _column_for() so the
    # same x0/x1 logic applies as for the anchor row (rejects tokens like
    # "16.5Each" that straddle the description→Units boundary).
    for (top, line), (upper, lower) in zip(anchor_rows, bands):
        desc_words = []
        for row in rows:
            for w in row:
                if (upper <= w["top"] < lower
                        and _column_for(w, anchors) == LineColumn.DESCRIPTION):
                    desc_words.append(w)
        desc_words.sort(key=lambda w: (round(w["top"]), w["x0"]))
        line[LineColumn.DESCRIPTION] = " ".join(w["text"] for w in desc_words)

    # Filter freight / backorder.
    line_items: list[dict] = []
    freight_count = 0
    backorder_count = 0
    for top, line in anchor_rows:
        if line[LineColumn.CODE].upper() == FREIGHT_CODE:
            freight_count += 1
            continue
        if backorder_y is not None and top > backorder_y:
            backorder_count += 1
            continue
        line_items.append(line)

    return line_items, freight_count, backorder_count, anchors


def _try_parse_anchor_row(
    row: list[dict],
    anchors: dict,
    all_rows: list[list[dict]] = (),
) -> Optional[dict]:
    by_col: dict[str, list[dict]] = {}
    for w in row:
        col = _column_for(w, anchors)
        if col is None or col == LineColumn.DESCRIPTION:
            continue
        by_col.setdefault(col, []).append(w)

    ln_words = by_col.get(LineColumn.LINE_NO, [])
    if not ln_words:
        return None

    code_words = by_col.get(LineColumn.CODE, [])
    if not code_words and all_rows:
        # Code may be on a slightly different y (wrapped product code that
        # printed above the Ln line). Search nearby rows for Code-column
        # words within ±3 × row-tolerance pixels of this row's top.
        ln_top = row[0]["top"]
        for other in all_rows:
            if other is row:
                continue
            if abs(other[0]["top"] - ln_top) > ROW_Y_TOLERANCE * 3:
                continue
            for w in other:
                if _column_for(w, anchors) == LineColumn.CODE:
                    code_words.append(w)
    if not code_words:
        return None

    ln_text = "".join(w["text"] for w in sorted(ln_words, key=lambda w: w["x0"]))
    if not re.fullmatch(r"\d+", ln_text):
        return None

    code = "".join(
        w["text"] for w in sorted(code_words, key=lambda w: (w["top"], w["x0"]))
    )

    qty = _parse_number(by_col.get(LineColumn.QTY, []))
    price = _parse_number(by_col.get(LineColumn.PRICE, []))
    disc = _parse_disc(by_col.get(LineColumn.DISC, []))
    price_after_disc = _parse_number(by_col.get(LineColumn.PRICE_AFTER_DISCOUNT, []))
    total_after_disc = _parse_number(by_col.get(LineColumn.TOTAL_AFTER_DISCOUNT, []))

    return {
        LineColumn.LINE_NO: int(ln_text),
        LineColumn.CODE: code,
        LineColumn.DESCRIPTION: "",   # filled by band assignment
        LineColumn.QTY: qty,
        LineColumn.PRICE: price,
        LineColumn.TOTAL_BEFORE_DISCOUNT: None,  # filled later (Qty * Price)
        LineColumn.DISC: disc,
        LineColumn.PRICE_AFTER_DISCOUNT: price_after_disc,
        LineColumn.TOTAL_AFTER_DISCOUNT: total_after_disc,
        LineColumn.RRP: None,
    }


def _parse_number(words: list[dict]) -> Optional[float]:
    if not words:
        return None
    raw = "".join(w["text"] for w in sorted(words, key=lambda w: w["x0"]))
    raw = raw.replace(",", "").replace("$", "").strip()
    try:
        return float(raw)
    except ValueError:
        return None


def _parse_disc(words: list[dict]) -> Optional[float]:
    if not words:
        return None
    raw = "".join(w["text"] for w in sorted(words, key=lambda w: w["x0"])).strip()
    if raw.endswith("%"):
        try:
            return float(raw[:-1]) / 100.0
        except ValueError:
            return None
    return _parse_number(words)


def _find_backorder_y(rows: list[list[dict]]) -> Optional[float]:
    target = BACKORDER_HEADING.rstrip(":")
    for row in rows:
        joined = " ".join(w["text"] for w in sorted(row, key=lambda w: w["x0"]))
        if target in joined:
            return row[0]["top"]
    return None


def _find_subtotal_y(rows: list[list[dict]]) -> Optional[float]:
    for row in rows:
        joined = " ".join(w["text"] for w in sorted(row, key=lambda w: w["x0"]))
        if joined.startswith("Sub Total"):
            return row[0]["top"]
    return None


def _extract_sub_total(full_text: str) -> Optional[float]:
    m = re.search(r"Sub\s+Total\s+([\d,]+\.\d{2})", full_text)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except ValueError:
        return None


# Header extraction --------------------------------------------------------

def _extract_header(full_text: str) -> dict:
    customer_name = _extract_label_value(full_text, "Customer Name")
    invoice_number = _extract_label_value(full_text, "Invoice Number")
    invoice_date_str = _extract_label_value(full_text, "Invoice Date")
    invoice_date = _parse_date(invoice_date_str) if invoice_date_str else None

    return {
        HeaderField.CUSTOMER_NAME: customer_name or "",
        HeaderField.INVOICE_DATE: invoice_date,
        HeaderField.ORDER_REFERENCE: invoice_number or "",
    }


def _extract_label_value(full_text: str, label: str) -> Optional[str]:
    pattern = rf"{re.escape(label)}:\s*([^\n]+)"
    matches = re.findall(pattern, full_text)
    if not matches:
        return None
    raw = matches[-1].strip()
    # Cut at the next label that may appear inline (e.g. "Speirs Jewellers Payment Terms").
    raw = re.split(r"\s+[A-Z][A-Za-z ]*:", raw)[0].strip()
    return raw or None


def _parse_date(raw: str) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None
