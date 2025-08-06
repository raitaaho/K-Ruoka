from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.mouse_button import MouseButton
import undetected_chromedriver as uc
import time
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import date
import json
import os
import re
import random
import streamlit as st
import glob

st.set_page_config(page_title="K-Ruoka Product Explorer", page_icon="📈")

st.markdown("# K-Ruoka Product Explorer")
st.write(
    """This is a product explorer for K-Ruoka that allows you to filter and download product data from the latest scraped Excel files. The data includes product names, categories, stores, units, prices, nutritional information, and discount validity dates.
You can select product categories to filter the data, and download the filtered results as a CSV file. The data is sourced from the latest Excel files matching the patterns 'discounted_product_prices_kruoka_.xlsx' and 'product_prices_kruoka_.xlsx'."""
)

st.header("Discounted Product Explorer")

# Find the latest Excel file matching the pattern
excel_files = glob.glob("discounted_product_prices_kruoka_*.xlsx")

if not excel_files:
    st.error("No Excel files found matching 'discounted_product_prices_kruoka_*.xlsx'")

else:
    latest_file = max(excel_files, key=os.path.getctime)

    # Load the Excel file
    df = pd.read_excel(latest_file, engine='openpyxl')

    columns = df.columns.tolist()
    column_names = st.multiselect("Select Columns to Display", columns, default=columns)
    if column_names:
        df = df.loc[:, column_names]

    if "Store" in df.columns:
        stores = st.multiselect("Select Store(s)", sorted(df["Store"].dropna().unique()))
        if stores:
            df = df[df["Store"].isin(stores)]

    if "Category" in df.columns:
        categories = st.multiselect("Select Categories", sorted(df["Category"].dropna().unique()))
        if categories:
            df = df[df["Category"].isin(categories)]
    
    # Display filtered data
    st.subheader("Filtered Discounted Products")
    st.data_editor(df)

    # Download button
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Filtered Discount Product Data as CSV",
        data=csv,
        file_name="filtered_kruoka_discount_product_data.csv",
        mime="text/csv"
    )

st.header("Product Explorer")

# Find the latest Excel file matching the pattern
excel_files2 = glob.glob("product_prices_kruoka_*.xlsx")

if not excel_files2:
    st.error("No Excel files found matching 'product_prices_kruoka_*.xlsx'")

else:
    latest_file2 = max(excel_files2, key=os.path.getctime)

    # Load the Excel file
    df2 = pd.read_excel(latest_file2, engine='openpyxl')

    columns2 = df2.columns.tolist()
    column_names2 = st.multiselect("Select Columns to Display", columns2, default=columns2)
    if column_names2:
        df2 = df2.loc[:, column_names2]

    if "Store" in df2.columns:
        stores2 = st.multiselect("Select Store(s)", sorted(df2["Store"].dropna().unique()))
        if stores2:
            df2 = df2[df2["Store"].isin(stores2)]

    if "Category" in df2.columns:
        categories2 = st.multiselect("Select Categories", sorted(df2["Category"].dropna().unique()))
        if categories2:
            df2 = df2[df2["Category"].isin(categories2)]
    
    # Display filtered data
    st.subheader("Filtered Products")
    st.data_editor(df2)

    # Download button
    csv2 = df2.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Filtered Product Data as CSV",
        data=csv2,
        file_name="filtered_kruoka_product_data.csv",
        mime="text/csv"
    )