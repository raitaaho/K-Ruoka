import streamlit as st
import pandas as pd
import glob
import os

# Find the latest Excel file matching the pattern
excel_files = glob.glob("product_prices_kruoka_*.xlsx")
if not excel_files:
    st.error("No Excel files found matching 'product_prices_kruoka_*.xlsx'")
else:
    latest_file = max(excel_files, key=os.path.getctime)

    # Load the Excel file
    df = pd.read_excel(latest_file, engine='openpyxl')

    # Streamlit app
    st.title("K-Ruoka Product Explorer")

    # Filters
    if "Category" in df.columns:
        category = st.selectbox("Select Category", sorted(df["Category"].dropna().unique()))
        df = df[df["Category"] == category]

    if "Store" in df.columns:
        store = st.selectbox("Select Store", sorted(df["Store"].dropna().unique()))
        df = df[df["Store"] == store]

    # Display filtered data
    st.subheader("Filtered Products")
    st.dataframe(df)

    # Download button
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Filtered Data as CSV",
        data=csv,
        file_name="filtered_kruoka_data.csv",
        mime="text/csv"
    )
