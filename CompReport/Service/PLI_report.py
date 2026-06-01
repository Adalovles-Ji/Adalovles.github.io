def PLI_report(df, latest_year):
    from io import BytesIO
    import pandas as pd
    import streamlit as st

    # Define target years and available PLI methods
    years = [latest_year - 2, latest_year - 1, latest_year]
    PLI_item = ["Operating Margin (OM)", "Net Cost Plus (NCP)", "Full Cost Markup (FCMP)", "Berry Ratio"]

    # User selection for PLI method
    PLI_method = st.sidebar.selectbox("Select Your Method", options=PLI_item, index=0, key="selected_item")

    # Initialize in-memory buffer for Excel output
    output = BytesIO()

    # Map P&L line items to their corresponding column prefixes
    pl_items = [
        ("Net Sales", "Net sales\nth USD\n"),
        ("Gross Profit", "Gross Profit\nth USD\n"),
        ("Operating Profit", "Operating P/L\nth USD\n"),
    ]

    # Extract and average values across selected years
    data = pd.DataFrame()
    for item_key, item_name in pl_items:
        # Build expected column names and filter for existing ones
        col_names = [f"{item_name}{year}" for year in years]
        valid_cols = [c for c in col_names if c in df.columns]
        
        if not valid_cols:
            st.error(f"Missing columns for {item_key}!\nExpected: {col_names}")
            st.stop()
            
        # Convert to numeric (coercing invalid strings to NaN) and compute row-wise mean
        temp_df = df[valid_cols].apply(pd.to_numeric, errors='coerce')
        data[item_key] = temp_df.mean(axis=1)

    # Assign unique company names as index
    company_names = df['Company name'].unique()
    data.index = company_names

    # --- PLI Calculation Functions ---
    def OM_calculation(d): return d['Operating Profit'] / d['Net Sales']
    def NCP_calculation(d): return d['Operating Profit'] / d['Gross Profit']
    def FCMP_calculation(d): return d['Gross Profit'] / d['Net Sales']
    def Berry_Ratio_calculation(d): return d['Gross Profit'] / (d['Gross Profit'] - d['Operating Profit'])

    # Execute calculation based on user selection
    calc_map = {
        "Operating Margin (OM)": OM_calculation,
        "Net Cost Plus (NCP)": NCP_calculation,
        "Full Cost Markup (FCMP)": FCMP_calculation,
        "Berry Ratio": Berry_Ratio_calculation
    }
    pl_data = calc_map[PLI_method](data).squeeze()  # Squeeze ensures a 1D Series

    # --- Generate Excel File ---
    def get_formats(workbook):
        """Define reusable cell formats."""
        return {
            'title': workbook.add_format({'size': 11, 'align': 'center', 'bold': True, 'bg_color': '#FFC000', 'font_color': '#002060', 'border': 1, 'valign': 'vcenter'}),
            'header_center': workbook.add_format({'size': 11, 'bold': True, 'bg_color': '#002060', 'font_color': '#FFC000', 'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'header_left': workbook.add_format({'size': 11, 'bold': True, 'bg_color': '#002060', 'font_color': '#FFC000', 'align': 'left', 'valign': 'vcenter', 'border': 1}),
            'center': workbook.add_format({'size': 10, 'bg_color': "#F3F7FF", 'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'left': workbook.add_format({'size': 10, 'bg_color': '#F3F7FF', 'border': 1, 'align': 'left', 'valign': 'vcenter'}),
            'PLI_data': workbook.add_format({'size': 10, 'num_format': '#,##0.00%;[Red](#,##0.00%)', 'bg_color': '#F3F7FF', 'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'quartile': workbook.add_format({'size': 10, 'num_format': '#,##0.00%;[Red](#,##0.00%)', 'bg_color': '#F3F7FF', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
            'whitespace': workbook.add_format({'border': 1, 'border_color': '#FFFFFF'}),
            'white_left': workbook.add_format({'size': 9, 'align': 'left'}),
            'white_right': workbook.add_format({'size': 9, 'align': 'right'})
        }

    sheet_name = 'PLI Report'
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(sheet_name)
        fmt = get_formats(workbook)

        # Apply whitespace framing
        for r in range(40):
            for c in range(15):
                worksheet.write(r, c, "", fmt['whitespace'])

        # Write report headers and metadata
        worksheet.merge_range('B2:J2', "PROFIT LEVEL INDICATOR BENCHMARK REPORT", fmt['title'])
        worksheet.write('B3', f"Set = {len(company_names)} Companies", fmt['white_left'])
        worksheet.write('E3', f"{latest_year-2}-{latest_year} Weighted Average", fmt['white_left'])
        worksheet.write('J3', f"Method = {PLI_method}", fmt['white_right'])
        worksheet.merge_range('G5:I5', "Benchmark Metrics", fmt['header_center'])

        # Write benchmark labels and table headers
        for idx, label in enumerate(["Count", "Count with PLI", "Upper Quartile", "Median", "Lower Quartile"], start=6):
            worksheet.write(f'G{idx}', label, fmt['center'])
        worksheet.write('C12', "No.", fmt['header_left'])
        worksheet.merge_range('D12:H12', "Company Name", fmt['header_left'])
        worksheet.write('I12', "", fmt['header_left'])

        # Loop through companies and write PLI values
        for i, company in enumerate(company_names):
            worksheet.write(i + 12, 2, i + 1, fmt['center']) # No.
            worksheet.merge_range(f"D{i + 13}:H{i + 13}", company, fmt['left'])

            value = pl_data.loc[company] if company in pl_data.index else 0
            worksheet.write(i + 12, 8, value, fmt['PLI_data']) 
        # Calculate and write benchmark statistics
        count = len(pl_data)
        count_with_pli = len(pl_data[pl_data > 0])
        upper_q = pl_data.quantile(0.75)
        med = pl_data.median()
        lower_q = pl_data.quantile(0.25)

        worksheet.merge_range('H6:I6', count, fmt['center'])
        worksheet.merge_range('H7:I7', count_with_pli, fmt['center'])
        worksheet.merge_range('H8:I8', upper_q, fmt['quartile'])
        worksheet.merge_range('H9:I9', med, fmt['quartile'])
        worksheet.merge_range('H10:I10', lower_q, fmt['quartile'])

        # Insert quartile comparison chart
        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'categories': f"'{sheet_name}'!$G$8:$G$10",
            'values':f"'{sheet_name}'!$H$8:$H$10",
            'fill': {'color': '#002060'},
            'data_labels': {'value': True, 'num_format': '#,##0.00%', 'font': {'size': 8}}
        })
        chart.set_chartarea({'fill': {'color': '#F3F7FF'}})
        chart.set_plotarea({'fill': {'color': "#FFFFFF"}})
        chart.set_x_axis({'line': {'color': '#000000'}, 'num_font': {'size': 9, 'color': "#000000"}, 'major_gridlines': {'visible': True}})
        chart.set_y_axis({'major_gridlines': {'visible': False}, 'label_position': 'none'})
        chart.set_legend({'position': 'none'})
        chart.set_size({'width': 300, 'height': 160})
        worksheet.insert_chart('C5', chart)

        # Set column widths and row heights
        worksheet.set_column('A:B', 2)
        worksheet.set_column('C:C', 5)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:F', 10)
        worksheet.set_column('G:G', 20)
        worksheet.set_column('H:H', 10)
        worksheet.set_column('J:J', 2)

        # height
        for r in range(40):
            worksheet.set_row(r, 20)

    # Reset buffer pointer before returning bytes
    output.seek(0)
    return output