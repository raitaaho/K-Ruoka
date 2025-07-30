import streamlit as st
import time
import numpy as np

st.set_page_config(
    page_title="Hello",
    page_icon="👋",
)

st.write("# Welcome! 👋")

st.sidebar.success("Select functionality above")

st.markdown(
    """
    This is a web scraper for K-Ruoka.fi, a Finnish grocery store website.
    It scrapes product data, including discounts, and saves it to an Excel file.

    **👈 Select to scrape product data or view it from the sidebar**
"""
)