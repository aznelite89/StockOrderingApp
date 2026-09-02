# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Run

Single-file Streamlit app. Launch with:

```bash
./run.sh
# expands to:
uv run --with streamlit --with pandas --with xlsxwriter streamlit run app.py
```

`requirements.txt` lists `streamlit pandas xlsxwriter requests`, but `run.sh` does not include `requests` — add `--with requests` if any new code imports it. There is no test suite, linter, or build step.

`launcher.bat` is the Windows equivalent and just runs `streamlit run app.py`.

## What the app does

Generates a weekly purchase-order worksheet (the "UNL Order Sheet") for Searay from six Unleashed CSV exports. Output is an `.xlsx` with a main sheet, a per-supplier order tab for each supplier the user enters, a `*_ALL` tab per supplier, and a `Calculation_Logic` reference sheet.

The six inputs and where they come from in Unleashed are documented in-app (and mirrored in the `Calculation_Logic` sheet); the upload widgets are ordered to match:

1. **PO Product Data** — `read_csv(..., skiprows=1)` (first row is a title banner).
2. **PO Sales Data** — `skiprows=1`. Monthly columns `MM1..MM12` live at positions 13:16 for the "last 3 months" slice (`sales_df.columns[13:16]`); column order matters.
3. **Warehouse Data** — no skip. Renames `*Product Code` → `Product Code`, `*SOH` → `SOH`, then sums SOH per product.
4. **Product List** — no skip. Consumes `Weight` and `Is Purchasable` (Unleashed's Purchasable toggle; Yes/No). If the column is absent every row is treated as purchasable.
5. **Transaction Detail** — `skiprows=1`. Used to derive `Last PO Date` / `Last PO Qty` (latest transaction per product, summed if multiple on that date).
6. **Special Order** — `skiprows=1`. Drives the yellow highlight and `Customer Allocations` column.

If any of these export formats change (column names, banner row presence, monthly column position), the load section near `app.py:168` breaks silently — values flow through `fillna` and end up as zeros rather than raising.

## Core calculation (single pass in `app.py`)

After merging all sources on `Product Code`:

```
12m Sales          = Total Sales + Allocated
Average Weekly Sales = 12m Sales / 52
Available Stock    = On Hand + On Purchase Order - Allocated
Target Stock Qty   = coverage_weeks * Average Weekly Sales   # coverage_weeks is user input (1-52, default 20)
Need To Order      = max(Target Stock Qty - Available Stock, 0)
Searay Order qty   = ceil(Need) if Base Unit == "each", round(Need, 3) if "weight", else ceil(Need)
Purchaseable       = "NO" if Obsolete == "YES" or Average Weekly Sales == 0 else "YES"
```

Note the inconsistency: `Average Weekly Sales (3m)` is `3m Sales / 12` in code (`app.py:194`) but the `Calculation_Logic` sheet documents it as `/13`. Don't "fix" one without confirming intent.

Rows are then re-sorted so that products sharing a `Bin Location` stay together, with bins ordered by the highest `Need To Order` they contain (see the `seen_bins` / `seen_products` loop around `app.py:319`). Don't replace this with a plain `sort_values` — bin grouping is intentional for picking workflow.

## Excel output structure

- `UNL_Order_Sheet` — full sheet; Product Code cells are hyperlinks to `searay.net.au/search?q=<code>`; rows whose Product Code appears in the Special Order export are yellow-highlighted; rows where `UNL Purchasable` is `NO` (Unleashed's Purchasable toggle off, from the Product List export) are red-highlighted, and red wins over yellow. `UNL Purchasable` is deliberately the LAST column (`Z`) so the VBA letters below stay valid — don't insert columns before it. Row writing for the main sheet and `_ALL` tabs is shared in `_write_highlighted_rows`. Note `Purchaseable` (column `Y`) is a different, derived flag (Obsolete/zero sales) — it is not the Unleashed toggle.
- `<SUPPLIER>` — per-supplier order form with a green "PROCESS ORDER" button and 1000 placeholder rows. The button is wired to a VBA macro (`ProcessOrderMacro`) that pulls from the `<SUPPLIER>_ALL` tab using **hard-coded column letters** (`A`=Product Code, `B`=Supplier Code (also the row filter), `C`=Supplier Product Code, `D`=Supplier Description, `E`=Weight, `V`=Searay Order (the order qty → "Qty (pcs)", and the `>0` include filter), `W`=Comments → "Notes"). These letters are positions in `final_cols`/`_ALL`, NOT the column names: `T`=Need To Order, `U`=Base Unit, `V`=Searay Order, `W`=Comments. `Searay Order` is a **blank column the user fills in by hand** — the macro only pulls rows where it holds a positive number, so on a freshly generated sheet (all blank) the macro finds nothing until quantities are entered. If you reorder `final_cols` (in `views/stock_ordering_sheet.py`, ~line 220) you MUST update the VBA letter references in the same file (~line 380) to match — there is no programmatic link between them. (History: before 2026-06-04 the macro read `U`/`V`, so "Qty (pcs)" showed "Each"; the 2026-06-04 fix over-corrected both to `T`=Need To Order; fixed 2026-07-16 to read `V`=Searay Order per Ron's request — the order form must reflect the hand-entered quantities, not the system suggestion.)
- `VBA_Code_<SUPPLIER>` — paste-in instructions for the macro (Excel doesn't let xlsxwriter inject real VBA, so the user copies it manually).
- `<SUPPLIER>_ALL` — full row set for that supplier with the same highlight/hyperlink treatment as the main sheet.
- `Calculation_Logic` — human-readable formula reference.

## Editing notes

- `app copy.py` is a backup snapshot, not an import target. Leave it alone unless asked.
- Streamlit reruns the whole script on every widget change; the `if all([...])` gate at `app.py:168` is what prevents work until all six files are uploaded.
- The version string is embedded twice (page title and H1, both `v4.9`); bump both together if releasing a new version.
- Column names, YES/NO values and highlight colours for the order sheet live in `constants/order_sheet.py`; don't add new raw string literals for these in the view.
