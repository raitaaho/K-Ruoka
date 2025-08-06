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

milligram_pattern = re.compile(r'(\d+(?:[.,]\d+)?)\s*mg')
gram_pattern = re.compile(r'(\d+(?:[.,]\d+)?)\s*g')
kilogram_pattern = re.compile(r'(\d+(?:[.,]\d+)?)\s*kg')
milliliter_pattern = re.compile(r'(\d+(?:[.,]\d+)?)\s*ml')
deciliter_pattern = re.compile(r'(\d+(?:[.,]\d+)?)\s*dl')
liter_pattern = re.compile(r'(\d+(?:[.,]\d+)?)\s*l')

def extract_size_in_g(size_string):
    string = size_string.replace(',', '.')

    for pattern, multiplier in [
        (milligram_pattern, float(1/1000)),
        (gram_pattern, 1),
        (kilogram_pattern, 1000),
        (deciliter_pattern, 100),
        (milliliter_pattern, 1),
        (liter_pattern, 1000),
    ]:
        matches = pattern.findall(string)
        if len(matches) >= 2:
            size = float(matches[1]) * multiplier
            if isinstance(size, float):
                return round(size, 2)
        elif len(matches) == 1:
            size = float(matches[0]) * multiplier
            if isinstance(size, float):
                return round(size, 2)

    return 'Unknown'


def extract_size_in_kg(size_string):
    string = size_string.replace(',', '.')

    for pattern, multiplier in [
        (milligram_pattern, float(1/1000000)),
        (gram_pattern, float(1/1000)),
        (kilogram_pattern, 1),
        (deciliter_pattern, float(1/10)),
        (milliliter_pattern, float(1/1000)),
        (liter_pattern, 1),
    ]:
        matches = pattern.findall(string)
        if len(matches) >= 2:
            size = float(matches[1]) * multiplier
            if isinstance(size, float):
                return round(size, 2)
        elif len(matches) == 1:
            size = float(matches[0]) * multiplier
            if isinstance(size, float):
                return round(size, 2)

    return 'Unknown'

def extract_portion_size(nutritional_header_string):
    string = nutritional_header_string.replace(',', '.')

    if milligram_match := milligram_pattern.search(string):
        return f"{milligram_match.group(1)} mg"
    if gram_match := gram_pattern.search(string):
        return f"{gram_match.group(1)} g"
    if kilogram_match := kilogram_pattern.search(string):
        return f"{kilogram_match.group(1)} kg"
    if milliliter_match := milliliter_pattern.search(string):
        return f"{milliliter_match.group(1)} ml"
    if deciliter_match := deciliter_pattern.search(string):
        return f"{deciliter_match.group(1)} dl"
    if liter_match := liter_pattern.search(string):
        return f"{liter_match.group(1)} l"
    
    return 'Unknown'

st.set_page_config(page_title="K-Ruoka Product Explorer", page_icon="📈")

st.markdown("# K-Ruoka Product Explorer")
st.write(
    """This is a product explorer for K-Ruoka that allows you to filter and download product data from the latest scraped Excel files. The data includes product names, categories, stores, units, prices, nutritional information, and discount validity dates.
You can select product categories to filter the data, and download the filtered results as a CSV file. The data is sourced from the latest Excel files matching the patterns 'discounted_product_prices_kruoka_.xlsx' and 'product_prices_kruoka_.xlsx'."""
)

st.header("Discounted Product Explorer")

# Find the latest Excel file matching the pattern
excel_files = glob.glob("discounted_product_prices_kruoka_*.xlsx")

if excel_files:
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
        label="Download Filtered Discounted Product Data as CSV",
        data=csv,
        file_name="filtered_kruoka_discounted_product_data.csv",
        mime="text/csv"
    )

else:
    st.error("No Excel files found matching 'discounted_product_prices_kruoka_*.xlsx', using JSON data instead.")
    try:
        with open('discounted_product_prices_data.json', 'r') as file:
            discounted_product_price_dict = json.load(file)
    except IOError:
        print("Could not open product price data file. Using empty dictionary.")
        discounted_product_price_dict = {}

    for ean, product_data in discounted_product_price_dict.items():
        portion_size_string = product_data.get('Nutritional Value per', 'Unknown')
        portion_size_in_grams = extract_size_in_g(portion_size_string)

        if portion_size_in_grams != 0 and portion_size_in_grams != 'Unknown':
            if product_data.get('Proteiini', 0) + product_data.get('Rasva', 0) + product_data.get('Hiilihydraatit', 0) <= portion_size_in_grams:
                product_data['Proteiinia per 100g'] = product_data.get('Proteiini', 0) * (100 / portion_size_in_grams)
            else:
                product_data['Proteiinia per 100g'] = product_data.get('Proteiini', 0)
        else:
            product_data['Proteiinia per 100g'] = product_data.get('Proteiini', 0)

    df = pd.DataFrame.from_dict(discounted_product_price_dict, orient='index')
    df['Size (kg)'] = df['Size (kg)'].apply(lambda x: x if isinstance(x, (float, int)) else 0)

    # Calculate 'Euroa per 100g Proteiinia' with zero-check for 'Proteiinia per 100g' and 'Size (kg)' columns
    df['Euroa per 100g Proteiinia'] = np.where(
        (df['Proteiinia per 100g'] > 0),
        (df['Price per kg'] / 10.00) * (100 / df['Proteiinia per 100g']),
        9999.999
    )

    df.index.name = 'EAN-code'

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
            df2 = df[df["Category"].isin(categories)]
    
    # Display filtered data
    st.subheader("Filtered Discounted Products")
    st.data_editor(df)

    # Download button
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Filtered Discounted Product Data as CSV",
        data=csv,
        file_name="filtered_kruoka_discounted_product_data.csv",
        mime="text/csv"
    )

st.header("Product Explorer")

# Find the latest Excel file matching the pattern
excel_files2 = glob.glob("product_prices_kruoka_*.xlsx")

if excel_files2:
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

else:
    st.error("No Excel files found matching 'product_prices_kruoka_*.xlsx', using JSON data instead.")
    try:
        with open('product_prices_data.json', 'r') as file2:
            product_price_dict = json.load(file2)
    except IOError:
        print("Could not open product price data file. Using empty dictionary.")
        product_price_dict = {}

    for ean, product_data in product_price_dict.items():
        portion_size_string = product_data.get('Nutritional Value per', 'Unknown')
        portion_size_in_grams = extract_size_in_g(portion_size_string)

        if portion_size_in_grams != 0 and portion_size_in_grams != 'Unknown':
            if product_data.get('Proteiini', 0) + product_data.get('Rasva', 0) + product_data.get('Hiilihydraatit', 0) <= portion_size_in_grams:
                product_data['Proteiinia per 100g'] = product_data.get('Proteiini', 0) * (100 / portion_size_in_grams)
            else:
                product_data['Proteiinia per 100g'] = product_data.get('Proteiini', 0)
        else:
            product_data['Proteiinia per 100g'] = product_data.get('Proteiini', 0)

    df2 = pd.DataFrame.from_dict(product_price_dict, orient='index')
    df2['Size (kg)'] = df2['Size (kg)'].apply(lambda x: x if isinstance(x, (float, int)) else 0)

    # Calculate 'Euroa per 100g Proteiinia' with zero-check for 'Proteiinia per 100g' and 'Size (kg)' columns
    df2['Euroa per 100g Proteiinia'] = np.where(
        (df2['Proteiinia per 100g'] > 0),
        (df2['Price per kg'] / 10.00) * (100 / df2['Proteiinia per 100g']),
        9999.999
    )

    df2.index.name = 'EAN-code'

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
