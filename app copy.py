import pandas as pd
import math
import io
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="UNL Order Sheet Generator v4.8", layout="wide")
st.title("UNL Order Sheet Generator v4.8")

# Add Download Sample Sheet button at the top
st.header("Download Sample Data Format")

# Create sample data for each required format
sample_output = io.BytesIO()

with pd.ExcelWriter(sample_output, engine='xlsxwriter') as writer:
    workbook = writer.book
 
    # Tab 1: Product Data (PO Product Data)
    product_data = pd.DataFrame({
        'Product Code':[],
        'Supplier Code':[],
        'Supplier Product Code':[],
        'Bin Location':[],
        'Supplier Product Description':[],
        'Product Group':[],
        'On Purchase':[],
        'Allocated':[],
        'Obsolete':[],
        'Base Unit':[]
    })
    product_data.to_excel(writer, sheet_name='1_Product_Data', index=False, startrow=1)
    
    # Add headers with formatting
    worksheet = writer.sheets['1_Product_Data']
    worksheet.write(0, 0, 'Product List as of dd/mm/yyyy')
    for col_num, value in enumerate(product_data.columns.values):
        worksheet.write(1, col_num, value)
    
    # Tab 2: Sales Data (Unit Sales Enquiry)
    sales_data = pd.DataFrame({
        'Product Code':[],
        'Product Description':[],
        'Pref Supplier':[],
        'Warehouse':[],
        'MM1':[], 
        'MM2':[],
        'MM3':[],
        'MM4':[], 
        'MM5':[], 
        'MM6':[], 
        'MM7':[], 
        'MM8':[], 
        'MM9':[], 
        'MM10':[], 
        'MM11':[], 
        'MM12':[], 
        'Total':[],
        'Stock On Hand':[]
    })
    sales_data.to_excel(writer, sheet_name='2_Sales_Data', index=False, startrow=1)
    
    worksheet = writer.sheets['2_Sales_Data']
    worksheet.write(0, 0, 'Unit Sales Enquiry as of dd/mm/yyyy')
    for col_num, value in enumerate(sales_data.columns.values):
        worksheet.write(1, col_num, value)
    
    # Tab 3: Warehouse Data (Stock on Hand)
    warehouse_data = pd.DataFrame({
        '*Product Code':[],
        'Warehouse Code':[],
        '*SOH':[],  # Stock on Hand
        '*AverageLandCost':[],
        'LastCost':[]
    })
    warehouse_data.to_excel(writer, sheet_name='3_Warehouse_Data', index=False)
    
    worksheet = writer.sheets['3_Warehouse_Data']
    for col_num, value in enumerate(warehouse_data.columns.values):
        worksheet.write(0, col_num, value)
    
    # Tab 4: Product List (for Weight info)
    product_list_data = pd.DataFrame({
        '*Product Code':[],
        '*Product Description':[],
        'Barcode':[],
        '...':[],
        'Default Purchasing Unit of Measure':[],
        'Is Batch Tracked':[]
    })
    product_list_data.to_excel(writer, sheet_name='4_Product_List', index=False)
    
    worksheet = writer.sheets['4_Product_List']
    for col_num, value in enumerate(product_list_data.columns.values):
        worksheet.write(0, col_num, value)
    
    # Tab 5: Transaction Detail (Purchase Order history)
    transaction_data = pd.DataFrame({
        'Transaction Date':[],
        'Transaction Ref':[],
        'Warehouse':[],
        'Transaction Type':[],
        'Product Code':[],
        'Product Description':[],
        'Value':[],
        'Quantity':[],
        'Running Total':[]



    })
    transaction_data.to_excel(writer, sheet_name='5_Transaction_Detail', index=False, startrow=1)
    
    worksheet = writer.sheets['5_Transaction_Detail']
    worksheet.write(0, 0, 'Transaction Enquiry as of dd/mm/yyyy')
    for col_num, value in enumerate(transaction_data.columns.values):
        worksheet.write(1, col_num, value)
    
    # Tab 6: Special Order (bonus tab)
    special_data = pd.DataFrame({
        'Order No.':[],
        'Order Date':[],
        'Required Date':[],
        'Completed Date':[],
        'Warehouse':[],
        'Customer':[],
        'Customer Type':[],
        'Product':[],
        'Product Code':[],
        'Product Group':[],
        'Status':[],
        'Quantity':[],
        'Sub Total':[]
    })
    special_data.to_excel(writer, sheet_name='6_Special_Order', index=False, startrow=1)
    
    worksheet = writer.sheets['6_Special_Order']
    worksheet.write(0, 0, 'Sales Enquiry as of dd/mm/yyyy')
    for col_num, value in enumerate(special_data.columns.values):
        worksheet.write(1, col_num, value)

st.download_button(
    label="📥 Download Sample Data Format (6 Tabs)",
    data=sample_output.getvalue(),
    file_name="Sample_Data_Format.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.header("Data Source Instructions")
st.text(" 1. Product Data: Unleased PO Product Data (Inventory>View Products>Grid Layout: PO Product Data>Export to CSV")
st.text(" 2. Sales Data: Unleased Sales Data (Reports>Sales>Unit Sales Enquiry>Grid Layout: Searay>Export to CSV) Make sure \"Date To\" is set to your latest month. Date From will automatically populate (12m)")
st.text(" 3. Warehouse Data: Warehouse Data (Inventory > Products > Import/Export > Stock on Hand)")
st.text(" 4. Product List: Product List (Inventory > Products > Import/Export > Products)")
st.text(" 5. Transaction Detail: Reports > Inventory > Transaction Enquiry > Transaction Type: Purchase Order > Grid Layout: Searay > Run > Export to CSV")
st.text(" 6. Special Order: Reports > Sales > Sales Enquiry > last 12m > Sales Order Status (Special Order) > Grid Layout: Searay > Hidden Columns(Drag Product Code in) > Run > Export to CSV")

coverage_weeks = st.number_input("Enter Purchase Coverage Weeks", min_value=1, max_value=52, value=20)
supplier_input = st.text_input("Enter Supplier Code(s) (comma separated, case insensitive):", value="", help="Enter one or multiple supplier codes separated by commas, e.g., ABC123, DEF456, GHI789")

st.header("Upload Data Files")
po_prod_file   = st.file_uploader("1) PO Product Data", type=["csv"])
po_sales_file  = st.file_uploader("2) PO Sales Data", type=["csv"])
warehouse_file = st.file_uploader("3) Warehouse Data", type=["csv"])
plist_file     = st.file_uploader("4) Product List (for Weight info)", type=["csv"])
original_file  = st.file_uploader("5) Transaction Detail (Original) File", type=["csv"])
special_file   = st.file_uploader("6) Special Order File", type=["csv"])

if all([po_prod_file, po_sales_file, warehouse_file, original_file, plist_file, special_file]):
    # --- Load Files ---
    product_df   = pd.read_csv(po_prod_file, skiprows=1)
    sales_df     = pd.read_csv(po_sales_file, skiprows=1)
    warehouse_df = pd.read_csv(warehouse_file)
    plist_df     = pd.read_csv(plist_file)
    original_df  = pd.read_csv(original_file, skiprows=1)
    special_df   = pd.read_csv(special_file, skiprows=1)

    # --- On Hand ---
    warehouse_df = warehouse_df.rename(columns={"*Product Code": "Product Code", "*SOH": "SOH"})
    on_hand_df = warehouse_df.groupby("Product Code", as_index=False)["SOH"].sum().rename(columns={"SOH": "On Hand"})

    # --- Sales ---
    sales_df["Total Sales"] = pd.to_numeric(sales_df["Total"], errors="coerce").fillna(0)
    
    # get the last 3 months sales
    month_cols = sales_df.columns[4:7]
    # ensure the columns are numeric
    for col in month_cols:
        sales_df[col] = pd.to_numeric(sales_df[col], errors='coerce').fillna(0)

    # calculate the sum of the last 3 months sales
    sales_df["3m Sales"] = sales_df[month_cols].sum(axis=1)

    # calculate the average weekly sales (13 weeks)
    sales_df["Average Weekly Sales (3m)"] = (sales_df["3m Sales"] / 12).round(3)
    sales_df = sales_df[["Product Code", "Total Sales", "3m Sales", "Average Weekly Sales (3m)"]]

    # --- On Purchase Order ---
    po_df = product_df.copy()
    po_df["On Purchase Order"] = pd.to_numeric(po_df["On Purchase"], errors="coerce").fillna(0)
    on_po_df = po_df[["Product Code", "On Purchase Order"]].copy()

    # --- Allocated ---
    alloc_df = product_df[["Product Code", "Allocated"]].copy()
    alloc_df["Allocated"] = pd.to_numeric(alloc_df["Allocated"], errors="coerce").fillna(0)

    # --- Static Info ---
    static_df = product_df.rename(columns={"Supplier Product Description": "Supplier Description"})[[
        "Product Code", "Supplier Code", "Supplier Product Code", "Supplier Description",
        "Product Group", "Bin Location", "Base Unit", "Obsolete"
    ]]

    # --- Transaction Records (Latest PO) ---
    original_df["Transaction Date"] = pd.to_datetime(original_df["Transaction Date"], errors="coerce")
    filtered_df = original_df.dropna(subset=["Product Code", "Transaction Date"])
    latest_dates = filtered_df.groupby("Product Code")["Transaction Date"].max().reset_index()
    latest_orders = pd.merge(filtered_df, latest_dates, on=["Product Code", "Transaction Date"], how="inner")
    
    # Group by Product Code and Transaction Date, sum the quantities for same product on same latest date
    latest_orders = latest_orders.groupby(["Product Code", "Transaction Date"], as_index=False)["Quantity"].sum()
    
    latest_orders["Transaction Date"] = latest_orders["Transaction Date"].dt.strftime("%d/%m/%Y")
    latest_orders = latest_orders.rename(columns={
        "Transaction Date": "Last PO Date",
        "Quantity": "Last PO Qty"
    })

    # --- Weight Info ---
    weight_df = plist_df.rename(columns={"*Product Code": "Product Code"})[["Product Code", "Weight"]]

    # --- Special Order ---
    special_ProductCode = special_df["Product Code"].unique()

    # --- Merge All ---
    df = (static_df
        .merge(weight_df, on="Product Code", how="left")
        .merge(on_hand_df, on="Product Code", how="left")
        .merge(on_po_df, on="Product Code", how="left")
        .merge(alloc_df, on="Product Code", how="left")
        .merge(sales_df, on="Product Code", how="left")
        .merge(latest_orders, on="Product Code", how="left"))

    df.fillna({
        "On Hand": 0,
        "On Purchase Order": 0,
        "Allocated": 0,
        "Total Sales": 0,
        "3m Sales": 0,
        "Average Weekly Sales (3m)": 0,
        "Last PO Qty": 0,
        "Last PO Date": "N/A",
        "Weight": "N/A",
        "Obsolete": "NO"
    }, inplace=True)

    # --- Inventory Logic ---
    df["12m Sales"] = df["Total Sales"] + df["Allocated"]
    df["Average Weekly Sales"] = (df["12m Sales"] / 52).round(3)
    df["Available Stock"] = df["On Hand"] + df["On Purchase Order"] - df["Allocated"]
    df["Target Stock Qty"] = coverage_weeks * df["Average Weekly Sales"]
    df["Need To Order"] = (df["Target Stock Qty"] - df["Available Stock"]).clip(lower=0)

    def order_qty(row):
        need = row["Need To Order"]
        bu = str(row["Base Unit"]).strip().lower()
        if not bu or bu in ["blank", "na", "none"]:
            bu = "each"
        if bu == "each":
            return math.ceil(need)
        if bu == "weight":
            return round(need, 3)
        return math.ceil(need)

    df["Need To Order"] = df.apply(order_qty, axis=1)
    df["Searay Order"] = ""
    df["Comments"] = ""

    # Purchaseable column
    df["Purchaseable"] = df.apply(lambda r: "NO" if r["Obsolete"] == "YES" or r["Average Weekly Sales"] == 0 else "YES", axis=1)


    # Final column order
    final_cols = [
        "Product Code", "Supplier Code", "Supplier Product Code", "Supplier Description", "Weight",
        "Product Group", "Bin Location", "Last PO Date", "Last PO Qty",
        "On Hand", "On Purchase Order", "Allocated", "Available Stock",
        "3m Sales", "Average Weekly Sales (3m)", "12m Sales", "Average Weekly Sales",
        "Target Stock Qty", "Need To Order",
        "Base Unit", "Searay Order", "Comments",
        "Obsolete", "Purchaseable"
    ]
    df = df[final_cols]

    # Step 1: Sort by Need To Order
    df["Bin Location"] = df["Bin Location"].fillna("UNKNOWN").astype(str)
   
    df_sorted = df.sort_values(by="Need To Order", ascending=False).copy()

    # Step 2: Initialize the list and set of processed bins and products
    final_rows = []
    seen_bins = set()
    seen_products = set()

    # Step 3: Iterate through the sorted product rows, organize bin groups as needed
    for _, row in df_sorted.iterrows():
        bin_loc = row["Bin Location"]
        
        # If the bin has been processed, skip
        if bin_loc in seen_bins:
            continue

        # Find all products in this bin, excluding already processed ones
        bin_df = df[(df["Bin Location"] == bin_loc) & (~df["Product Code"].isin(seen_products))].copy()
        bin_df = bin_df.sort_values(by="Need To Order", ascending=False)

        # Add to the result
        final_rows.append(bin_df)

        # Mark the processed bin and product
        seen_bins.add(bin_loc)
        seen_products.update(bin_df["Product Code"].tolist())

    # Step 4: Concatenate the final sorted DataFrame
    df = pd.concat(final_rows, ignore_index=True)

    df.fillna("N/A", inplace=True)


    # Output
    st.success("✅ UNL Order Sheet Generated Successfully")
    st.dataframe(df)

    # --- Supplier Filter ---
    st.header("Supplier Filter")
    
    if supplier_input.strip():
        # case insensitive, handle multiple suppliers
        supplier_codes = [code.strip().upper() for code in supplier_input.split(',')]
        supplier_df = df[df["Supplier Code"].str.upper().isin(supplier_codes)].copy()
        
        if not supplier_df.empty:
            st.success(f"✅ Found {len(supplier_df)} products for suppliers: {', '.join(supplier_codes)}")
            st.subheader(f"Suppliers: {', '.join(supplier_codes)}")
            st.dataframe(supplier_df)
        else:
            st.warning(f"❌ No products found for suppliers: {', '.join(supplier_codes)}")
            # show available supplier codes
            available_suppliers = sorted(df["Supplier Code"].dropna().unique())
            st.info(f"Available Supplier Codes: {', '.join(available_suppliers)}")
    else:
        st.info("There is no supplier filter applied.")

    # Excel writing
    logic_text = [
        ["Field Name", "Description"],
        ["On Hand", "From Warehouse Data *SOH aggregated by Product Code"],
        ["On Purchase Order", "From PO Product Data On Purchase field"],
        ["Allocated", "From PO Product Data Allocated field"],
        ["Total Sales", "From PO Sales Data Total field"],
        ["3m Sales", "Last 3 months of sales from PO Sales"],
        ["Average Weekly Sales (3m)", "3m Sales / 13"],
        ["12m Sales", "Total Sales + Allocated"],
        ["Average Weekly Sales", "12m Sales / 52"],
        ["Available Stock", "On Hand + On Purchase Order - Allocated"],
        ["Target Stock Qty", "Average Weekly Sales (12m) * coverage_weeks (user input)"],
        ["Need To Order", "Target - Available, if less than 0 then 0"],
        ["Searay Order", "Based on Base Unit, if each then round up, if weight then round to 3 decimal places"],
        ["Last PO Date / Qty", "From Transaction Detail, get the latest purchase date and quantity"],
        ["Weight", "From Product List"],
        ["Obsolete", "From PO Product Data"],
        ["Purchaseable", "NO if Obsolete is YES or Average Weekly Sales == 0, else YES"],
        ["Product Data", "Unleased PO Product Data (Inventory>View Products>Grid Layout: PO Product Data>Export to CSV"],
        ["Sales Data", "Unleased Sales Data (Reports>Sales>Unit Sales Enquiry>Grid Layout: Searay>Export to CSV) Make sure \"Date To\" is set to your latest month. Date From will automatically populate (12m)"],
        ["Warehouse Data", "Warehouse Data (Inventory > Products > Import/Export > Stock on Hand)"],
        ["Product List", "Product List (Inventory > Products > Import/Export > Products)"],
        ["Transaction Detail", "Reports > Inventory > Transaction Enquiry > Transaction Type: Purchase Order > Grid Layout: Searay > Run > Export to CSV"],
        ["Special Order", "Reports > Sales > Sales Enquiry > last 12m > Sales Order Status (Special Order) > Grid Layout: Searay > Hidden Columns(Drag Product Code in) > Run > Export to CSV"]
    ]
    logic_df = pd.DataFrame(logic_text[1:], columns=logic_text[0])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        
        # write full UNL order sheet
        df.to_excel(writer, sheet_name='UNL_Order_Sheet', index=False, startrow=1, header=False)
        
        workbook = writer.book
        url_format = workbook.add_format({'font_color': 'blue', 'underline': 1})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
        yellow_highlight = workbook.add_format({'bg_color': '#FFFF00'})  # yellow background
        
        # handle main table
        main_worksheet = writer.sheets['UNL_Order_Sheet']
        for col_num, value in enumerate(df.columns.values):
            main_worksheet.write(0, col_num, value, header_format)
        
        # Add autofilter and freeze top row for main sheet
        main_worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
        main_worksheet.freeze_panes(1, 0)  # Freeze the top row
        
        main_code_col = df.columns.get_loc("Product Code")
        for row_num, (index, row) in enumerate(df.iterrows(), start=1):
            product_code = row["Product Code"]
            
            # check if the product code is in the special order list
            is_special = product_code in special_ProductCode
            
            # write each cell
            for col_num, value in enumerate(row):
                if col_num == main_code_col:
                    # Product Code column write URL
                    url = f"https://searay.net.au/search?q={product_code}&options%5Bprefix%5D=last"
                    if is_special:
                        # if it is a special order, use the yellow background URL format
                        special_url_format = workbook.add_format({
                            'font_color': 'blue', 
                            'underline': 1, 
                            'bg_color': '#FFFF00'
                        })
                        main_worksheet.write_url(row_num, col_num, url, special_url_format, string=product_code)
                    else:
                        main_worksheet.write_url(row_num, col_num, url, url_format, string=product_code)
                else:
                    # other columns
                    if is_special:
                        main_worksheet.write(row_num, col_num, value, yellow_highlight)
                    else:
                        main_worksheet.write(row_num, col_num, value)
        
        # --- Create tabs for the inputted supplier ---
        if supplier_input.strip():
            supplier_codes = [code.strip().upper() for code in supplier_input.split(',')]
            for supplier_code in supplier_codes:
                supplier_df_full = df[df["Supplier Code"].str.upper() == supplier_code].copy()
                
                if not supplier_df_full.empty:
                    # clean supplier code as sheet name
                    sheet_name = str(supplier_code).replace('/', '_').replace('\\', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_').replace('?', '_')[:31]
                    
                    # create worksheet
                    worksheet = workbook.add_worksheet(sheet_name)
                    
                    # write supplier info header
                    worksheet.write('A1', 'To:', header_format)
                    worksheet.write('B1', supplier_code, header_format)
                    worksheet.write('A2', 'From:', header_format)
                    worksheet.write('B2', 'Searay Pty Ltd')
                    worksheet.write('A3', 'PO:', header_format)
                    worksheet.write('B3', 'PO-00000084')
                    worksheet.write('A4', 'PO Date:', header_format)
                    worksheet.write('B4', datetime.now().strftime('%d/%m/%Y'), header_format)
                    
                    # Add a button-like cell for Excel automation
                    button_format = workbook.add_format({
                        'bold': True,
                        'font_color': 'white',
                        'bg_color': '#4CAF50',  # Green background
                        'border': 2,
                        'border_color': '#2E7D32',
                        'align': 'center',
                        'valign': 'vcenter',
                        'font_size': 12
                    })
                    
                    # Create button-like cells
                    worksheet.merge_range('F1:F2', 'PROCESS ORDER', button_format)
                    
                    # Add instruction for making it clickable
                    worksheet.write('F3', 'Click button to update calculations', header_format)
                    
                    # write header (from the 6th row)
                    start_row = 5
                    headers = ['Line no.', 'Product Code', 'Supplier Product Code', 'Product Description', 'Weight', 'Qty (pcs)', 'Total Weight', 'Notes']
                    for col_num, header in enumerate(headers):
                        worksheet.write(start_row, col_num, header, header_format)
                    
                    # Instead of dynamic formulas, create static rows with placeholders
                    # We'll populate these with VBA when button is clicked
                    for i in range(1, 1001):  # Create 1000 empty rows for data
                        excel_row = start_row + i
                        
                        # Initialize with empty values - VBA will populate these
                        worksheet.write(excel_row, 0, "")  # Line no.
                        worksheet.write(excel_row, 1, "")  # Product Code
                        worksheet.write(excel_row, 2, "")  # Supplier Product Code
                        worksheet.write(excel_row, 3, "")  # Product Description
                        worksheet.write(excel_row, 4, "")  # Weight
                        worksheet.write(excel_row, 5, "")  # Qty (pcs)
                        worksheet.write(excel_row, 6, "")  # Total Weight
                        worksheet.write(excel_row, 7, "")  # Notes
                    
                    # Add VBA code to the workbook
                    vba_code = f'''
                        Sub ProcessOrderMacro()
                            Dim ws As Worksheet
                            Dim mainWs As Worksheet
                            Dim supplierCode As String
                            Dim lastRow As Long
                            Dim i As Long, j As Long
                            Dim dataRow As Long
                            Dim lineNo As Long
                            
                            ' Set references
                            Set ws = ActiveSheet
                            Set mainWs = ThisWorkbook.Sheets("{supplier_code}_ALL")
                            supplierCode = "{supplier_code}"
                            
                            ' Clear existing data (rows 7-1006)
                            ws.Range("A7:H1006").ClearContents
                            
                            ' Find data in main sheet and populate supplier sheet
                            lastRow = mainWs.Cells(mainWs.Rows.Count, "B").End(xlUp).Row
                            dataRow = 7  ' Start from row 7 (after headers)
                            lineNo = 1
                            
                            For i = 2 To lastRow  ' Start from row 2 (skip header)
                                ' Check if this row matches our supplier and has quantity > 0
                                If UCase(mainWs.Cells(i, "B").Value) = UCase(supplierCode) And _
                                mainWs.Cells(i, "U").Value > 0 Then
                                    
                                    ' Line no.
                                    ws.Cells(dataRow, 1).Value = lineNo
                                    
                                    ' Product Code
                                    ws.Cells(dataRow, 2).Value = mainWs.Cells(i, "A").Value
                                    
                                    ' Supplier Product Code
                                    ws.Cells(dataRow, 3).Value = mainWs.Cells(i, "C").Value
                                    
                                    ' Product Description
                                    ws.Cells(dataRow, 4).Value = mainWs.Cells(i, "D").Value
                                    
                                    ' Weight (check if it's "N/A")
                                    If mainWs.Cells(i, "E").Value = "N/A" Then
                                        ws.Cells(dataRow, 5).Value = ""
                                    Else
                                        ws.Cells(dataRow, 5).Value = mainWs.Cells(i, "E").Value
                                    End If
                                    
                                    ' Qty (pcs) - Get from Searay Order column (U)
                                    ws.Cells(dataRow, 6).Value = mainWs.Cells(i, "U").Value
                                    
                                    ' Total Weight (Weight * Qty if both are numeric)
                                    If IsNumeric(mainWs.Cells(i, "E").Value) And IsNumeric(mainWs.Cells(i, "U").Value) Then
                                        ws.Cells(dataRow, 7).Value = mainWs.Cells(i, "E").Value * mainWs.Cells(i, "U").Value
                                    Else
                                        ws.Cells(dataRow, 7).Value = ""
                                    End If
                                    
                                    ' Notes (empty for user input)
                                    ws.Cells(dataRow, 8).Value = mainWs.Cells(i, "V").Value
                                    
                                    dataRow = dataRow + 1
                                    lineNo = lineNo + 1
                                    
                                    ' Stop if we reach 1000 items
                                    If lineNo > 1000 Then Exit For
                                End If
                            Next i
                            
                            MsgBox "Order data updated successfully! Found " & (lineNo - 1) & " items for supplier " & supplierCode
                        End Sub
                        '''
                    
                    # Create a button assignment instruction
                    worksheet.write('F4', 'Button will run ProcessOrderMacro', header_format)
                    
                    # adjust column width
                    worksheet.set_column('A:A', 8)   # Line no.
                    worksheet.set_column('B:B', 12)  # Product Code
                    worksheet.set_column('C:C', 15)  # Supplier Product Code
                    worksheet.set_column('D:D', 50)  # Product Description
                    worksheet.set_column('E:E', 10)  # Weight
                    worksheet.set_column('F:F', 15)  # Qty (pcs)
                    worksheet.set_column('G:G', 15)  # Total Weight
                    worksheet.set_column('H:H', 25)  # Notes
                    
                    # Create a separate worksheet for VBA code
                    vba_worksheet = workbook.add_worksheet(f'VBA_Code_{supplier_code}'[:31])
                    vba_worksheet.write('A1', f'VBA Code for {supplier_code} ProcessOrderMacro', header_format)
                    vba_worksheet.write('A2', 'Instructions:', header_format)
                    vba_worksheet.write('A3', '1. Press Alt+F11 to open VBA Editor')
                    vba_worksheet.write('A4', '2. Insert > Module')
                    vba_worksheet.write('A5', '3. Copy the code below and paste it in the module')
                    vba_worksheet.write('A6', '4. Close VBA Editor')
                    vba_worksheet.write('A7', '5. Right-click the green button in {sheet_name} tab')
                    vba_worksheet.write('A8', '6. Select "Assign Macro..." and choose "ProcessOrderMacro"')
                    vba_worksheet.write('A10', 'VBA Code:', header_format)
                    
                    vba_lines = vba_code.split('\n')
                    for idx, line in enumerate(vba_lines):
                        vba_worksheet.write(10 + idx, 0, line)
                    
                    # Set column width for VBA sheet
                    vba_worksheet.set_column('A:A', 80)
                    
                    # --- Create ALL PRODUCTS tab for the supplier ---
                    all_products_sheet_name = f"{sheet_name}_ALL"[:31]
                    all_products_worksheet = workbook.add_worksheet(all_products_sheet_name)
                    
                    # Write all products for this supplier to the new tab
                    supplier_df_full.to_excel(writer, sheet_name=all_products_sheet_name, index=False, startrow=1, header=False)
                    
                    # Add headers with formatting
                    for col_num, value in enumerate(supplier_df_full.columns.values):
                        all_products_worksheet.write(0, col_num, value, header_format)
                    
                    # Add autofilter and freeze top row for supplier ALL sheet
                    all_products_worksheet.autofilter(0, 0, len(supplier_df_full), len(supplier_df_full.columns) - 1)
                    all_products_worksheet.freeze_panes(1, 0)  # Freeze the top row
                    
                    # Format the data rows with special order highlighting and product links
                    code_col = supplier_df_full.columns.get_loc("Product Code")
                    for row_num, (index, row) in enumerate(supplier_df_full.iterrows(), start=1):
                        product_code = row["Product Code"]
                        is_special = product_code in special_ProductCode
                        
                        for col_num, value in enumerate(row):
                            if col_num == code_col:
                                # Product Code column with URL
                                url = f"https://searay.net.au/search?q={product_code}&options%5Bprefix%5D=last"
                                if is_special:
                                    special_url_format = workbook.add_format({
                                        'font_color': 'blue', 
                                        'underline': 1, 
                                        'bg_color': '#FFFF00'
                                    })
                                    all_products_worksheet.write_url(row_num, col_num, url, special_url_format, string=product_code)
                                else:
                                    all_products_worksheet.write_url(row_num, col_num, url, url_format, string=product_code)
                            else:
                                # Other columns
                                if is_special:
                                    all_products_worksheet.write(row_num, col_num, value, yellow_highlight)
                                else:
                                    all_products_worksheet.write(row_num, col_num, value)

        # write calculation logic table
        logic_df.to_excel(writer, sheet_name='Calculation_Logic', index=False)

    # --- Download Button ---
    if supplier_input.strip():
        supplier_codes = [code.strip().upper() for code in supplier_input.split(',')]
        if len(supplier_codes) == 1:
            download_label = f"Download Order Sheet with {supplier_codes[0]} Tabs"
            file_name = f"UNL_Order_Sheet_with_{supplier_codes[0]}.xlsx"
        else:
            download_label = f"Download Order Sheet with {len(supplier_codes)} Supplier Tabs"
            file_name = f"UNL_Order_Sheet_with_{len(supplier_codes)}_Suppliers.xlsx"
    else:
        download_label = "Download UNL Order Sheet"
        file_name = "UNL_Order_Sheet.xlsx"

    st.download_button(
        label=download_label,
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )










