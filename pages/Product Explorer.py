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

st.markdown("# Product Explorer")
st.write(
    """This is a product explorer for K-Ruoka that allows you to filter and download product data from the latest scraped Excel file. The data includes product names, categories, stores, units, prices, nutritional information, and discount validity dates.
You can select product categories to filter the data, and download the filtered results as a CSV file. The data is sourced from the latest Excel file matching the pattern 'discounted_product_prices_kruoka_*.xlsx'."""
)

st.header("K-Ruoka Product Explorer")

# Find the latest Excel file matching the pattern
excel_files = glob.glob("discounted_product_prices_kruoka_*.xlsx")

if not excel_files:
    st.error("No Excel files found matching 'discounted_product_prices_kruoka_*.xlsx'")

else:
    latest_file = max(excel_files, key=os.path.getctime)

    # Load the Excel file
    df = pd.read_excel(latest_file, engine='openpyxl')

    # Filters
    if "Category" in df.columns:
        categories = st.multiselect("Select Category", sorted(df["Category"].dropna().unique()))
        
        if categories:
            df = df[df["Category"].isin(categories)]

    # Display filtered data
    st.subheader("Filtered Products")
    st.data_editor(df)

    # Download button
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Filtered Data as CSV",
        data=csv,
        file_name="filtered_kruoka_data.csv",
        mime="text/csv"
    )