import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 1. PAGE SETUP
st.set_page_config(page_title="Lead Report Automator", page_icon="🎯", layout="wide")
st.title("🎯 Lead Report Automator")
st.write("Upload your lead data file to generate automated reports.")

# 2. FILE UPLOADER
uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Load the file - check for header row location
        df_test = pd.read_excel(uploaded_file, engine='openpyxl', nrows=2)
        
        # Check if first row looks like it has valid headers
        if df_test.columns[0] == 'Unnamed: 0' or pd.isna(df_test.columns[0]):
            # Headers are likely in row 2
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, engine='openpyxl', header=1)
        else:
            # Headers are in row 1
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, engine='openpyxl', header=0)
        
        # Remove any unnamed columns
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        st.success(f"✅ File uploaded successfully! Loaded {len(df)} records.")
        
        # Display column names for debugging
        with st.expander("📋 View Column Names & Data Preview"):
            st.write(f"Total columns: {len(df.columns)}")
            st.write(df.columns.tolist())
            st.write("### First 5 Rows:")
            st.dataframe(df.head(), use_container_width=True)
        
        st.info("📊 Please tell me what reports and metrics you'd like to generate from this data.")
        
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.write("Please ensure you're uploading a valid Excel file.")
