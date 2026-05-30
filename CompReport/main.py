import streamlit as st
import pandas as pd
from io import BytesIO
import turtle

# 设置页面配置
st.set_page_config(page_title="Financial Dashboard Generator", layout="wide")
st.title("📊 Automated Financial Dashboard Generator")

# 侧边栏输入
latest_year = st.sidebar.text_input("Enter Latest Year", value="2024")

# 主页面上传器
uploaded_file = st.file_uploader("Upload your financial database (Excel)", type=["xlsx", "xls"])

def get_formats(workbook):
    """定义Excel样式格式"""
    return {
        'comp_name_text': workbook.add_format({'size': 12, 'align': 'center','bold': True, 'bg_color': '#002060', 'font_color': '#FFC000', 'border': 1, 'valign': 'vcenter'}),
        'header': workbook.add_format({'size': 10, 'bold': True, 'bg_color': '#FFFFFF', 'font_color': '#002060', 'align': 'center', 'valign': 'vcenter', 'bottom': 6,'right': 1, 'left': 1}),
        'year_header': workbook.add_format({'size': 10, 'bold': True, 'bg_color': '#002060', 'font_color': '#FFC000', 'align': 'center', 'valign': 'vcenter', 'border': 1}),
        'cell_text': workbook.add_format({'size': 10, 'border': 1, 'bg_color': "#F3F7FF", 'align': 'left'}),
        'cell_num': workbook.add_format({'size': 10, 'num_format': '#,##0', 'border': 1, 'bg_color': '#F3F7FF'}),
        'cell_num_neg': workbook.add_format({'size': 10, 'num_format': '#,##0;[Red](#,##0)', 'border': 1, 'bg_color': '#F3F7FF'}),
        'Item': workbook.add_format({'size': 10, 'border': 1, 'bg_color': '#F3F7FF', 'align': 'center'}),
        'IPO_date': workbook.add_format({'num_format': 'yyyy-mm-dd', 'size': 10, 'border': 1, 'bg_color': '#F3F7FF', 'align': 'center'}),
        'website': workbook.add_format({'size': 10, 'border': 1, 'bg_color': "#FFFFFF",'font_color': "#FF0000", 'underline': 1}),
        'whitespace': workbook.add_format({'border': 1, 'border_color': '#FFFFFF'})
    }

if uploaded_file:
    try:
        # 读取上传的Excel文件,第一列删除，第二列作为索引
        df = pd.read_excel(uploaded_file, sheet_name="Raw Data" or "result")
        
        # 定义财务项目映射
        SI_items = [
            ("Primary SIC", "Primary US SIC code"),
            ("Ticker Symbol", "Ticker symbol"),
            ("Stock Exchange", "Main exchange"),
            ("IPO date", "IPO date"),
            ("Market Cap(Mil US$)", "Current market capitalisation\nth USD"),
            ("Number of Employees", "No of\nemployees\nLast Year\nLast avail. yr")
        ]
        
        KR_items = [
            ("EBIT%", "Earnings Before Interest & Tax\nth USD\n"),
            ("Gross Margin%", "Gross Margin%"),
            ("ROCE%", "ROCE%"),
            ("Receivable Days", "Receivable Days"),
            ("Inventory Days", "Inventory Days"),
            ("Sales-Employee (th US$)", "Sales-Employee (th US$)")
        ]
        
        pl_items = [
            ("Net Sales", "Net sales\nth USD\n"),
            ("Cost of Goods Sold", "COGS\nth USD\n"),
            ("Gross Profit", "Gross Profit\nth USD\n"),
            ("OPEX", "OPEX\nth USD\n"),
            ("Operating Profit", "Operating P/L\nth USD\n"),
            ("EBIT Reported", "Earnings Before Interest & Tax\nth USD\n"),
            ("Financial Expenses", "Other non Oper./Financial Inc./Exp.\nth USD\n"),
            ("EBT Reported", "Earnings before tax\nth USD\n"),
            ("Taxes Paid", "Taxes Paid\nth USD\n")
        ]
        
        bl_items = [
            ("Accounts Receivable", "Accounts receivable\nth USD\n"),
            ("Inventory", "Net Stated Inventory\nth USD\n"),
            ("Net PP&E", "Net property, plant & equipment\nth USD\n"),
            ("Other Assets", "Other Assets\nth USD\n"),
            ("Total Assets", "Total Assets\nth USD\n"),
            ("Accounts Payable", "Accounts payable\nth USD\n"),
            ("LT Debt", "Long Term Debt\nth USD\n"),
            ("Equity", "Total shareholders' equity\nth USD\n"),
            ("Other Liabilities", "Other Liabilities\nth USD\n"),
            ("Total Liabilities", "Total Liabilities\nth USD\n")
        ]

        if st.button("Generate Formatted Reports"):
            output = BytesIO()
            
            # 使用 pd.ExcelWriter 管理工作簿生命周期
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = None # 初始化工作表变量
                
                # 获取样式
                fmt = get_formats(workbook)
                
                # 计算年份列表
                try:
                    current_year = int(latest_year)
                    years = [str(current_year), str(current_year-1), str(current_year-2), '3-Year Avg']
                except ValueError:
                    st.error("Please enter a valid year.")
                    st.stop()
                
                for _, row in df.iterrows():
                    comp_name = str(row.iloc[0])[:31].replace("/", "_")
                    
                    # 添加新的工作表
                    worksheet = workbook.add_worksheet(comp_name)

                    # 定义白色填充和无边框样式
                    for row_num in range(40):
                        for col in range(12):
                            worksheet.write(row_num, col, "", fmt['whitespace'])

                    ### Section 1: 表头与框架 ###
                    # 合并单元格写入标题
                    worksheet.merge_range('B2:J2', comp_name, fmt['comp_name_text'])
                    worksheet.merge_range('B3:C3', 'Item', fmt['header'])
                    worksheet.merge_range('D3:H3', 'Business Description', fmt['header'])
                    worksheet.write('I3', 'Key Ratio', fmt['header'])
                    worksheet.write('J3', latest_year, fmt['header'])
                    
                    # P&L 表头
                    worksheet.write('B12', 'P&L (th US$)', fmt['year_header'])
                    for i, y in enumerate(years):
                        worksheet.write(11, i + 2, y, fmt['year_header'])
                    worksheet.write('F12', '3-Year Avg', fmt['year_header'])
                    worksheet.merge_range('H12:J12', 'Net Sales (th US$)', fmt['year_header'])
                    
                    # Balance Sheet 表头
                    worksheet.write('B23', 'Balance Sheet (th US$)', fmt['year_header'])
                    for i, y in enumerate(years):
                        worksheet.write(22, i + 2, y, fmt['year_header'])
                    worksheet.write('F23', '3-Year Avg', fmt['year_header'])
                    worksheet.merge_range('H23:J23', 'Operating Profit (th US$)', fmt['year_header'])
                    
                    ### Section 2: 写入数据行 ###
                    # 写入 SI (公司信息)
                    for row_idx, (label, base_key) in enumerate(SI_items, start=1):
                        worksheet.write(row_idx + 2, 1, label, fmt['cell_text'])
                        val = row.get(base_key, "-")
                        
                        # 处理日期类型 (NaT)
                        if pd.isna(val):
                            val = "-"
                        elif isinstance(val, pd.Timestamp):
                            date_val = val.date()
                            worksheet.write(row_idx + 2, 2, date_val, fmt['IPO_date'])
                        elif isinstance(val, (int, float)): #
                            if type(val) == float:
                                worksheet.write(row_idx + 2, 2, round(val, 2), fmt['Item'])
                            else:
                                worksheet.write(row_idx + 2, 2, val, fmt['Item'])
                        else:
                            worksheet.write(row_idx + 2, 2, str(val), fmt['Item'])
                    
                    # 写入网站地址
                    website_address = row.get("Website address", "")
                    if pd.isna(website_address):
                        website_address = "http://0"
                    worksheet.merge_range('B10:J10', website_address, fmt['website'])
                    
                    worksheet.merge_range('D4:H9',"") # 合并单元格用于业务描述
                    
                    # 写入 KR (关键比率)
                    for row_idx, (label, base_key) in enumerate(KR_items, start=1):
                        worksheet.write(row_idx + 2, 8, label, fmt['cell_text'])
                        val = row.get(base_key, 0)
                        
                        # 安全转换为浮点数
                        def safe_float(val):
                            try:
                                return float(val)
                            except (ValueError, TypeError):
                                return 0.0
                        
                        num_val = safe_float(val)
                        if num_val < 0:
                            worksheet.write(row_idx + 2, 9, num_val, fmt['cell_num_neg'])
                        else:
                            worksheet.write(row_idx + 2, 9, num_val, fmt['cell_num'])
                    
                    # 写入 PL (损益表) 和 计算 3年平均
                    for row_idx, (label, base_key) in enumerate(pl_items, start=1):
                        worksheet.write(row_idx + 11, 1, label, fmt['cell_text'])
                        values = []
                        
                        # 写入各年份数据
                        for col_idx, y in enumerate(years):
                            if y == '3-Year Avg':
                                # 计算平均值
                                if values: # 防止除以0
                                    # 修正：计算所有可用年份的平均值
                                    avg_val = sum(values) / len(values)
                                else:
                                    avg_val = 0
                                # 写入平均值
                                if avg_val < 0:
                                    worksheet.write(row_idx + 11, 5, avg_val, fmt['cell_num_neg']) # F列 (索引5) 是平均值列
                                else:
                                    worksheet.write(row_idx + 11, 5, avg_val, fmt['cell_num'])
                            else:
                                col_name = f"{base_key}{y}"
                                val = row.get(col_name, 0)
                                num_val = safe_float(val)
                                values.append(num_val)
                                
                                # 写入单元格
                                if num_val < 0:
                                    worksheet.write(row_idx + 11, col_idx + 2, num_val, fmt['cell_num_neg'])
                                else:
                                    worksheet.write(row_idx + 11, col_idx + 2, num_val, fmt['cell_num'])
                    
                    # 写入 BL (资产负债表) 和 计算 3年平均
                    for row_idx, (label, base_key) in enumerate(bl_items, start=1):
                        worksheet.write(row_idx + 22, 1, label, fmt['cell_text'])
                        values = []
                        
                        for col_idx, y in enumerate(years):
                            if y == '3-Year Avg':
                                if values:
                                    avg_val = sum(values) / len(values)
                                else:
                                    avg_val = 0
                                if avg_val < 0:
                                    worksheet.write(row_idx + 22, 5, avg_val, fmt['cell_num_neg'])
                                else:
                                    worksheet.write(row_idx + 22, 5, avg_val, fmt['cell_num'])
                            else:
                                col_name = f"{base_key}{y}"
                                val = row.get(col_name, 0)
                                num_val = safe_float(val)
                                values.append(num_val)
                                
                                if num_val < 0:
                                    worksheet.write(row_idx + 22, col_idx + 2, num_val, fmt['cell_num_neg'])
                                else:
                                    worksheet.write(row_idx + 22, col_idx + 2, num_val, fmt['cell_num'])
                    
                    ### Section 3 & 4: 创建图表 ###
                    # 1. 净销售额柱状图
                    chart_NS = workbook.add_chart({'type': 'column'})
                    # 动态获取数据范围：第12行是标题，第13行是数据 (Net Sales 是 pl_items 的第一个)
                    # categories: C12:E12 (年份)
                    # values: C13:E13 (数据)
                    chart_NS.add_series({
                        'categories': [comp_name, 11, 2, 11, 5], # C12:F12
                        'values': [comp_name, 12, 2, 12, 5],     # C13:F13
                        'fill': {'color': '#002060'},
                        'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'size': 8}} # 显示数值标签，格式为千分位整数
                    })
                    chart_NS.set_chartarea({'fill': {'color': '#F3F7FF'}})
                    chart_NS.set_plotarea({'fill': {'color': "#FFFFFF"}})
                    chart_NS.set_x_axis({'line': {'color': 'Black'}, 'num_font': {'size': 9, 'color': 'Black'}, 'major_gridlines': {'visible': True}})
                    chart_NS.set_y_axis({'major_gridlines': {'visible': False}, 'label_position': 'none'})
                    chart_NS.set_legend({'position': 'none'})
                    worksheet.insert_chart('H13', chart_NS)
                    chart_NS.set_size({'width': 253, 'height': 170})
                    
                    # 2. 运营利润柱状图
                    chart_IS = workbook.add_chart({'type': 'column'})
                    # 运营利润是 pl_items 的第5个 (索引4)，所以行号是 12 + 4 = 16 (即第17行)
                    # 但根据你的模板，图表引用的是 C24:F24，这里保持模板结构，引用第23行(标题)和第24行(数据)
                    # 假设运营利润是 BL 或 PL 中的第一个数值行，这里手动指定为第23行(标题)和24行(数据)
                    chart_IS.add_series({
                        'categories': [comp_name, 22, 2, 22, 5], # C23:F23
                        'values': [comp_name, 23, 2, 23, 5],     # C24:F24
                        'fill': {'color': '#002060'},
                        'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'size': 8}} # 显示数值标签，格式为千分位整数，标签字体颜色为黑色
                    })
                    chart_IS.set_chartarea({'fill': {'color': '#F3F7FF'}})
                    chart_IS.set_plotarea({'fill': {'color': "#FFFFFF"}})
                    chart_IS.set_x_axis({'line': {'color': '#000000'}, 'num_font': {'size': 9, 'color': "#000000"}, 'major_gridlines': {'visible': True}})
                    chart_IS.set_y_axis({'major_gridlines': {'visible': False}, 'label_position': 'none'})
                    chart_IS.set_legend({'position': 'none'})
                    worksheet.insert_chart('H24', chart_IS)
                    chart_IS.set_size({'width': 253, 'height': 170})
                    
                    ### Section 5: 调整列宽和行高 ###
                    worksheet.set_column('A:A', 2)
                    worksheet.set_column('B:B', 20)
                    worksheet.set_column('C:E', 11) # 修正：只设置到E列，F列是平均值
                    worksheet.set_column('F:F', 11) # 单独设置平均值列
                    worksheet.set_column('G:H', 4)
                    worksheet.set_column('I:I', 20)
                    worksheet.set_column('J:J', 10)

                    ### 计算KR_items （2024） ###
                    # EBIT% (EBIT Reported / Net Sales) * 100
                    try:
                        ebit_reported = row.get(f"Earnings Before Interest & Tax\nth USD\n{latest_year}", 0)
                        net_sales = row.get(f"Net sales\nth USD\n{latest_year}", 0)
                        if net_sales != 0:
                            ebit_percent = (ebit_reported / net_sales) * 100
                        else:
                            ebit_percent = "-"
                    except Exception as e:
                        ebit_percent = "-"


                    # Gross Margin% (Gross Profit / Net Sales) * 100
                    try:
                        gross_profit = row.get(f"Gross Profit\nth USD\n{latest_year}", 0)
                        if net_sales != 0:
                            gross_margin_percent = (gross_profit / net_sales) * 100
                        else:
                            gross_margin_percent = "-"
                    except Exception as e:
                        gross_margin_percent = "-"
                    
                    # ROCE% (Operating Profit / Total Assets) * 100
                    try:
                        operating_profit = row.get(f"Operating P/L\nth USD\n{latest_year}", 0)
                        total_assets = row.get(f"Total Assets\nth USD\n{latest_year}", 0)
                        if total_assets != 0:
                            roce_percent = (operating_profit / total_assets) * 100
                        else:
                            roce_percent = "-"
                    except Exception as e:
                        roce_percent = "-"

                    # Receivable Days (Accounts Receivable / Net Sales) * 365
                    try:
                        accounts_receivable = row.get(f"Accounts receivable\nth USD\n{latest_year}", 0)
                        if net_sales != 0:
                            receivable_days = (accounts_receivable / net_sales) * 365
                        else:
                            receivable_days = "-"
                    except Exception as e:
                        receivable_days = "-"

                    # Inventory Days (Inventory / Net Sales) * 365
                    try:
                        inventory = row.get(f"Net Stated Inventory\nth USD\n{latest_year}", 0)
                        if net_sales != 0:
                            inventory_days = (inventory / net_sales) * 365
                        else:
                            inventory_days = "-"
                    except Exception as e:
                        inventory_days = "-"

                    # Sales-Employee (Net Sales / Number of Employees)
                    try:
                        number_of_employees = row.get(f"No of\nemployees\nLast Year\nLast avail. yr", 0)
                        if number_of_employees != 0:
                            sales_employee = net_sales / number_of_employees
                        else:
                            sales_employee = "-"
                    except Exception as e:
                        sales_employee = "-"

                    # 将计算的关键比率写入对应单元格
                    kr_values = [ebit_percent, gross_margin_percent, roce_percent, receivable_days, inventory_days, sales_employee]
                    for idx, val in enumerate(kr_values):
                        if isinstance(val, (int, float)):
                            if val < 0:
                                worksheet.write(idx + 3, 9, val, fmt['cell_num_neg'])
                            else:
                                worksheet.write(idx + 3, 9, val, fmt['cell_num'])
                        else:
                            worksheet.write(idx + 3, 9, val, fmt['cell_text'])
                        if idx == 0 or idx == 1 or idx == 2: # 保留两位小数
                            worksheet.write(idx + 3, 9, val, workbook.add_format({'num_format': '0.00', 'border': 1, 'bg_color': '#F3F7FF'}))

                    # 设置行高
                    for row in range(35):
                        worksheet.set_row(row, 20)

            # 将生成的文件提供下载
            st.download_button(
                label="📥 Download Reports",
                data=output.getvalue(),
                # 文件重新命名为Company_Report_{year}_{公司数量}.xlsx
                file_name=f"Company_Report_{latest_year}_{df.shape[0]}_Comps.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
        st.write("Please check if the uploaded file has a sheet named 'Raw Data' and the column names match the expected format.")
