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
from IPython.display import display
from datetime import datetime
import json
import os
import re
from collections import OrderedDict

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
            return float(matches[1]) * multiplier
        elif len(matches) == 1:
            return float(matches[0]) * multiplier

    return None


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
            return float(matches[1]) * multiplier
        elif len(matches) == 1:
            return float(matches[0]) * multiplier

    return None

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

nutritional_content_keys_dict = {
    'Energiaa': 'Energia',
    'Rasvaa': 'Rasva',
    'Rasvaa, josta tyydyttyneitä rasvoja': 'josta tyydyttyneitä',
    'Hiilihydraattia': 'Hiilihydraatit',
    'Hiilihydraattia, joista sokereita': 'josta sokereita',
    'Proteiinia': 'Proteiini',
    'Ravintokuitua': 'Ravintokuitu',
    'Suolaa': 'Suola'
}

product_categories = [
    'liha-ja-kasviproteiinit-1/kinkut-ja-leikkeleet',
    'liha-ja-kasviproteiinit-1/kana-broileri-ja-kalkkuna',
    'liha-ja-kasviproteiinit-1/porsas',
    'liha-ja-kasviproteiinit-1/jauheliha',
    'liha-ja-kasviproteiinit-1/kasviproteiinit',
    'juustot-tofut-ja-kasvipohjaiset/pala-ja-viipalejuustot',
    'juustot-tofut-ja-kasvipohjaiset/ruoka-ja-herkuttelujuustot',
    'juustot-tofut-ja-kasvipohjaiset/tofut-ja-kasvipohjaiset',
    'pakasteet-1/liha-ja-kalapakasteet',
    'pakasteet-1/pakasteateriat',
    'hillot-ja-sailykkeet/kasvissailykkeet/pavut',
    'maito-munat-ja-rasvat-0/rahkat-vanukkaat-ja-jalkiruoka',
    'kuivatuotteet-ja-leivonta-1/leseet-alkiot-rouheet',
    'kuivatuotteet-ja-leivonta-1/siemenet'
]

try:
    with open('nutritional_content_data.json', 'r') as file:
        nutritional_content_dict = json.load(file)
except IOError:
    print("Could not open nutritional content data file. Using empty dictionary.")
    nutritional_content_dict = {}

driver = uc.Chrome()
driver.get(f"https://www.s-kaupat.fi/tuotteet/")
time.sleep(5)
wait = WebDriverWait(driver, 3)
try:

    # Hae shadow root
    shadow_host = driver.find_element(By.CSS_SELECTOR, "#usercentrics-root")
    shadow_root = driver.execute_script("return arguments[0].shadowRoot", shadow_host)

    # Hae ja klikkaa "Hyväksy kaikki" -nappi
    accept_button = shadow_root.find_element(By.CSS_SELECTOR, "button[data-testid='uc-accept-all-button']")
    accept_button.click()
    time.sleep(1)

except TimeoutException:
    print('Prompt to accept cookies did not pop up')

product_dict = {}
product_urls = {}

for category in product_categories:
    driver.get(f"https://www.s-kaupat.fi/tuotteet/{category}")
    time.sleep(5)
    wait = WebDriverWait(driver, 3)
    try:
        products_list = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@data-test-id='product-list']")))
        product_cards = products_list.find_elements(By.XPATH, ".//article[@data-test-id='product-card']")
        last_list_len = len(product_cards)
        while True:
            driver.execute_script("arguments[0].scrollIntoView()", product_cards[-1])
            time.sleep(5)
            product_cards = driver.find_elements(By.XPATH, "//div[@data-test-id='product-list']//article[@data-test-id='product-card']")
            new_list_len = len(product_cards)
            if new_list_len == last_list_len:
                break
            last_list_len = new_list_len
    except TimeoutException:
        print("No products found in")
        continue

    for card in product_cards:
        ean_code = card.get_attribute("data-product-id")
        url = card.find_element(By.XPATH, ".//a[@href]").get_attribute("href")
        product_name = card.find_element(By.XPATH, ".//a[@href]").text
        size = extract_size_in_kg(product_name)

        sydanmerkki = 'No'
        hyvaa_suomesta = 'No'
        luomu = 'No'

        attribute_elements = card.find_elements(By.XPATH, ".//img[@title]")
        for attribute_element in attribute_elements:
            attribute_text = attribute_element.get_attribute("title")
            if attribute_text == 'Sydänmerkki':
                sydanmerkki = 'Yes'
            if attribute_text == 'Hyvää Suomesta, Sininen Joutsen':
                hyvaa_suomesta = 'Yes'
            if attribute_text == 'EU:n luomutunnus, lehti':
                luomu = 'Yes'

        if nutritional_content_dict.get(ean_code, 'Unknown') != 'Unknown':
            nutritional_content_dict[ean_code]['Sydänmerkki'] = sydanmerkki if nutritional_content_dict[ean_code].get('Sydänmerkki', 'Unknown') == 'Unknown' or sydanmerkki == 'Yes' else nutritional_content_dict[ean_code]['Sydänmerkki']
            nutritional_content_dict[ean_code]['Hyvää Suomesta'] = hyvaa_suomesta if nutritional_content_dict[ean_code].get('Hyvää Suomesta', 'Unknown') == 'Unknown' or hyvaa_suomesta == 'Yes' else nutritional_content_dict[ean_code]['Hyvää Suomesta']
            nutritional_content_dict[ean_code]['Luomu'] = luomu if nutritional_content_dict[ean_code].get('Luomu', 'Unknown') == 'Unknown' or luomu == 'Yes' else nutritional_content_dict[ean_code]['Luomu']
            try:
                unit_price_string = card.find_element(By.XPATH, ".//span[@data-test-id='product-price__comparisonPrice']//span").text
                backslash_index = unit_price_string.find("/")
                if backslash_index != -1:
                    unit_type = unit_price_string[backslash_index+1:]
                    unit_price = unit_price_string[:backslash_index-1].strip()
                    unit_price = float(unit_price.replace(',', '.')) if len(unit_price) != 0 else 999.999
                else:
                    search_res = re.search(r'(\d+(?:[.,]\d+)?)', unit_price_string)
                    unit_price = float(search_res.group().replace(',', '.')) if search_res else 999.999
                    unit_type = 'kg' if 'kg' in unit_price_string else 'Unknown'

            except Exception as e:
                print("Could not get product price for", e)
                unit_price = 999.999
                unit_type = 'Unknown'

            if product_dict.get(ean_code, "None") != "None":
                if product_dict[ean_code].get('Price per Unit', 'Unknown') != 'Unknown':
                    if product_dict[ean_code]['Price per Unit'] > unit_price:
                        product_dict[ean_code]['Price per Unit'] = unit_price
                        product_dict[ean_code]['Unit'] = unit_type
                        product_dict[ean_code]['Size (kg)'] = size
                        product_dict[ean_code]['Category'] = category

                else:
                    product_dict[ean_code]['Price per Unit'] = unit_price
                    product_dict[ean_code]['Unit'] = unit_type
                    product_dict[ean_code]['Size (kg)'] = size
                    product_dict[ean_code]['Category'] = category

            else:
                product_dict[ean_code] = {}
                product_dict[ean_code]['Price per Unit'] = unit_price
                product_dict[ean_code]['Unit'] = unit_type
                product_dict[ean_code]['Size (kg)'] = size
                product_dict[ean_code]['Category'] = category

            product_dict[ean_code].update(nutritional_content_dict[ean_code])

        else:
            product_urls.update({url: ean_code})
            product_dict[ean_code] = {}
            product_dict[ean_code]['Name'] = product_name
            product_dict[ean_code]['Size (kg)'] = size
            product_dict[ean_code]['Category'] = category

for url, ean in product_urls.items():
    if url is not None:
        driver.get(url)
        time.sleep(1)
    else:
        print("Could not open link")
        continue
    
    product_name = driver.find_element(By.XPATH, "//h1[@data-test-id='product-name']").text
    sydanmerkki = 'No'
    hyvaa_suomesta = 'No'
    luomu = 'No'

    attribute_elements = driver.find_elements(By.XPATH, "//div[@data-test-id='product-page-top-area']//ul[@aria-label='Tuotemerkinnät']//img[@title]")
    for attribute_element in attribute_elements:
        attribute_text = attribute_element.get_attribute("title")
        if attribute_text == 'Sydänmerkki':
            sydanmerkki = 'Yes'
        if attribute_text == 'Hyvää Suomesta, Sininen Joutsen':
            hyvaa_suomesta = 'Yes'
        if attribute_text == 'EU:n luomutunnus, lehti':
            luomu = 'Yes'

    keys_list = []
    values_list = []
    try:
        wait = WebDriverWait(driver, 3)
        table = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@data-test-id='nutrients-info-content']//table")))

        unit_size_string = table.find_element(By.XPATH, ".//thead//tr//th[2]").text
        backslash_index = unit_size_string.find("/")
        if backslash_index != -1:
            unit_size = extract_portion_size(unit_size_string[:backslash_index])
        else:
            unit_size = extract_portion_size(unit_size_string[:backslash_index])

        keys = table.find_elements(By.XPATH, ".//tbody//tr//th")
        values = table.find_elements(By.XPATH, ".//tbody//tr//td[1]")

        for key_element in keys:
            key = key_element.text
            keys_list.append(nutritional_content_keys_dict.get(key, key))
        
        for value_element in values:
            value = value_element.text
            kcal_index = value.find("kcal")
            kj_index = value.find("kJ")
            backslash_index = value.find("/")

            if kcal_index != -1:
                if backslash_index != -1:
                    value_string = value[backslash_index+1:kcal_index].replace(',', '.').strip()
                    values_list.append(float(value_string) if len(value_string) != 0 else 0.0)
                else:
                    if kj_index != -1:
                        value_string = value[:kj_index].replace(',', '.').strip()
                        values_list.append(float(value_string) * 0.2390057 if len(value_string) != 0 else 0.0)
                    else:
                        value_string = value[:kcal_index].replace(',', '.').strip()
                        values_list.append(float(value_string) if len(value_string) != 0 else 0.0)

            else:
                if kj_index != -1:
                    value_string = value[:kj_index].replace(',', '.').replace(' ', '').strip()
                    values_list.append(float(value_string) * 0.2390057 if len(value_string) != 0 else 0.0)
                else:
                    values_list.append(extract_size_in_g(value)) if value != None else 0.0

        kv_pairs = dict(zip(keys_list, values_list))

    except TimeoutException:
        print("Could not scrape nutritional info for", product_name)
        continue

    nutritional_content_dict[ean] = {}
    nutritional_content_dict[ean]['Name'] = product_name
    nutritional_content_dict[ean]['Nutritional Value per'] = unit_size
    nutritional_content_dict[ean].update(kv_pairs)
    nutritional_content_dict[ean]['Sydänmerkki'] = sydanmerkki
    nutritional_content_dict[ean]['Hyvää Suomesta'] = hyvaa_suomesta
    nutritional_content_dict[ean]['Luomu'] = luomu
    product_dict[ean].update(nutritional_content_dict[ean])

try:
    # Serializing json
    nutritional_content_json = json.dumps(nutritional_content_dict, indent=4)
    # Writing to json
    with open("nutritional_content_data.json", "w") as outfile:
        outfile.write(nutritional_content_json)
except IOError:
    print("JSON file is possibly open, close the file before continuing")
    time.sleep(30)
    try:
        nutritional_content_json = json.dumps(nutritional_content_dict, indent=4)
        with open("nutritional_content_data.json", "w") as outfile:
            outfile.write(nutritional_content_json)
    except Exception as e:
        print("Could not write nutritional content data to JSON file", e)