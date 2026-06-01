import re
import streamlit as st
import pandas as pd
from Service.company_report import company_report
from Service.PLI_report import PLI_report

st.set_page_config(page_title="Financial Dashboard Generator", layout="wide")
st.title("📊 Financial Dashboard Generator")

report_year = st.sidebar.text_input("Enter Report Year", value=str(pd.Timestamp.now().year - 1))
latest_year = int(report_year) - 1

uploaded_file = st.file_uploader("Upload your financial database (Excel)", type=["xlsx", "xls"])

if 'report_output' not in st.session_state:
    st.session_state.report_output = None

if 'report_name' not in st.session_state:
    st.session_state.report_name = "Default_Report" 

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Raw Data")
    except Exception as e:
        st.error(f"An error occurred while reading the file: {e}")
        st.stop()

    if st.button("Generate Company Report"):
        output = company_report(df, latest_year)
        st.session_state.report_output = output
        st.session_state.report_name = "Company_Report" 
        st.success("Company Report generated successfully!")

    if st.button("Generate PLI Report"):
        output = PLI_report(df, latest_year)
        st.session_state.report_output = output
        st.session_state.report_name = "PLI_Report"
        st.success("PLI Report generated successfully!")

    if st.session_state.report_output is not None:
        st.download_button(
            label="📥 Download Reports",
            data=st.session_state.report_output.getvalue(),
            file_name=f"{st.session_state.report_name}_{report_year}_{df.shape[0] + 1}_Comps.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )