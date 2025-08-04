# K-Ruoka Product Scraper

This Python script is a web scraper designed to collect product information, nutritional content, and pricing data from the K-Ruoka website. It provides a Streamlit-based user interface for selecting store locations and product categories, and outputs structured data for further analysis.

## Features

- Interactive Streamlit interface for user input
- Search and select K-Ruoka store locations
- Choose product categories to scrape
- Extract product details including:
  - Name, size, price, unit
  - Nutritional attributes (e.g., protein, fat, carbohydrates)
  - Dietary labels (e.g., vegan, gluten-free, organic)
  - Discount information and validity dates
- Save scraped data to JSON and Excel files
- Progress indicators for scraping status

## Installation

1. Clone the Repository
    ```bash
    git clone https://github.com/raitaaho/K-ruoka.git
    cd K-ruoka
    ```

3. Install Dependencies
    ```
    pip install -r requirements.txt
    ```

## Usage

1. Run the script using Streamlit:
    ```bash
    streamlit run Home.py
    ```
