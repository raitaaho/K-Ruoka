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
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import shutil
import undetected_chromedriver as uc
from selenium_stealth import stealth

today = date.today()

# Category options for user selection

category_options = {
    "Lihat ja kasviproteiinit": "liha-ja-kasviproteiinit",
    "Kala ja merenelävät": "kala-ja-merenelavat",
    "Valmisruoka": "valmisruoka",
    "Maito, juusto, munat ja rasvat": "maito-juusto-munat-ja-rasvat",
    "Murot ja myslit": "kuivat-elintarvikkeet-ja-leivonta/murot-ja-myslit",
    "Leseet, rouheet, alkiot, soijavalmisteet ja viljajyvät": "kuivat-elintarvikkeet-ja-leivonta/leseet-rouheet-alkiot-soijavalmisteet-ja-muut-viljatuotteet/leseet-rouheet-alkiot-soijavalmisteet-ja-viljajyvat",
    "Kuivatut herneet, pavut ja linssit": "kuivat-elintarvikkeet-ja-leivonta/kuivatut-herneet-pavut-ja-linssit",
    "Siemenet, pähkinät ja kuivatut hedelmät": "kuivat-elintarvikkeet-ja-leivonta/siemenet-pahkinat-ja-kuivatut-hedelmat",
    "Säilykkeet, keitot ja ainesosat": "sailykkeet-keitot-ja-ateria-ainekset",
    "Pakasteet": "pakasteet",
    "Makeiset ja naposteltavat": "makeiset-ja-naposteltavat",
    "Virvoitusjuomat": "juomat/virvoitusjuomat",
    "Energia- ja urheilujuomat": "juomat/energia--ja-urheilujuomat",
    "Kivennäis- ja lähdevedet": "juomat/kivennais--ja-lahdevedet",
    "Urheiluvalmisteet": "kosmetiikka-terveys-ja-hygienia/terveysvalmisteet/urheiluvalmisteet"
}

inverted_category_options = {
    "liha-ja-kasviproteiinit": "Lihat ja kasviproteiinit",
    "kala-ja-merenelavat": "Kala ja merenelävät",
    "valmisruoka": "Valmisruoka",
    "maito-juusto-munat-ja-rasvat": "Maito, juusto, munat ja rasvat",
    "kuivat-elintarvikkeet-ja-leivonta/murot-ja-myslit": "Murot ja myslit",
    "kuivat-elintarvikkeet-ja-leivonta/leseet-rouheet-alkiot-soijavalmisteet-ja-muut-viljatuotteet/leseet-rouheet-alkiot-soijavalmisteet-ja-viljajyvat": "Leseet, rouheet, alkiot, soijavalmisteet ja viljajyvät",
    "kuivat-elintarvikkeet-ja-leivonta/kuivatut-herneet-pavut-ja-linssit": "Kuivatut herneet, pavut ja linssit",
    "kuivat-elintarvikkeet-ja-leivonta/siemenet-pahkinat-ja-kuivatut-hedelmat": "Siemenet, pähkinät ja kuivatut hedelmät",
    "sailykkeet-keitot-ja-ateria-ainekset": "Säilykkeet, keitot ja ainesosat",
    "pakasteet": "Pakasteet",
    "makeiset-ja-naposteltavat": "Makeiset ja naposteltavat",
    "juomat/virvoitusjuomat": "Virvoitusjuomat",
    "juomat/energia--ja-urheilujuomat": "Energia- ja urheilujuomat",
    "juomat/kivennais--ja-lahdevedet": "Kivennäis- ja lähdevedet",
    "kosmetiikka-terveys-ja-hygienia/terveysvalmisteet/urheiluvalmisteet": "Urheiluvalmisteet"
}


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

def get_caffeine_amount(driver):
    percent_pattern = re.compile(r'kofeiin(ipitoisuus)?[ia]*\s*\(\s*(\d+,\d+|\d+)\s*%\s*\)', re.IGNORECASE)
    percent_pattern_2 = re.compile(r'(\d+,\d+|\d+)\s*%\s*\)?[^a-zA-Z0-9]{0,10}kofeiin(ipitoisuus)?[ia]*', re.IGNORECASE)
    percent_pattern_3 = re.compile(r'\(([^)]*kofeiin[ia]*[^)]*?)(\d+,\d+|\d+)\s*%\)', re.IGNORECASE)
    mg_100ml_pattern = re.compile(r'kofeiin(ipitoisuus)?[ia]*\s*\(\s*(\d+)\s*mg\s*/\s*100\s*ml\s*\)', re.IGNORECASE)
    mg_100ml_pattern_2 = re.compile(r'(\d+)\s*mg\s*/\s*100\s*ml[^a-zA-Z0-9]{0,10}kofeiin(ipitoisuus)?[ia]*', re.IGNORECASE)
    mg_100ml_pattern_3 = re.compile(r'\(([^)]*kofeiin[ia]*[^)]*?)(\d+)\s*mg\s*/\s*100\s*ml\)', re.IGNORECASE)
    mg_100ml_pattern_4 = re.compile(r'kofeiin(ipitoisuus)?[ia]*\s*(\d+)\s*mg\s*/\s*100\s*ml', re.IGNORECASE)
    mg_l_pattern = re.compile(r'kofeiin(ipitoisuus)?[ia]*\s*\(\s*(\d+)\s*mg\s*/\s*l\s*\)', re.IGNORECASE)
    mg_l_pattern_2 = re.compile(r'(\d+)\s*mg\s*/\s*l[^a-zA-Z0-9]{0,10}kofeiin(ipitoisuus)?[ia]*', re.IGNORECASE)
    mg_l_pattern_3 = re.compile(r'\(([^)]*kofeiin[ia]*[^)]*?)(\d+)\s*mg\s*/\s*l\)', re.IGNORECASE)
    amount_pattern = re.compile(r'(\d+(?:[.,]\d+)?)\s*mg\s*kofeiin[ia]*', re.IGNORECASE)

    caffeine_amount = 0.0
    caffeine_content = 0.0
    
    try:
        wait = WebDriverWait(driver, 3)
        product_description = wait.until(EC.presence_of_element_located((By.XPATH, "//p[starts-with(@class, 'ProductDetailsstyle__Description')]"))).text
        if mg_100ml_match := mg_100ml_pattern.search(product_description):
            caffeine_content = float(mg_100ml_match.group(2)) 
        elif mg_100ml_match_2 := mg_100ml_pattern_2.search(product_description):
            caffeine_content = float(mg_100ml_match_2.group(1)) 
        elif mg_100ml_match_3 := mg_100ml_pattern_3.search(product_description):
            caffeine_content = float(mg_100ml_match_3.group(2))
        elif mg_100ml_match_4 := mg_100ml_pattern_4.search(product_description):
            caffeine_content = float(mg_100ml_match_4.group(2))
        elif mg_l_match := mg_l_pattern.search(product_description):
            caffeine_content = float(mg_l_match.group(2)) / 10
        elif mg_l_match_2 := mg_l_pattern_2.search(product_description):
            caffeine_content = float(mg_l_match_2.group(1)) / 10
        elif mg_l_match_3 := mg_l_pattern_3.search(product_description):
            caffeine_content = float(mg_l_match_3.group(2)) / 10
        elif percent_match := percent_pattern.search(product_description):
            percent_string = percent_match.group(2).replace(',', '.')
            caffeine_content = 1000 * float(percent_string)
        elif percent_match_2 := percent_pattern_2.search(product_description):
            percent_string_2 = percent_match_2.group(1).replace(',', '.')
            caffeine_content = 1000 * float(percent_string_2)
        elif percent_match_3 := percent_pattern_3.search(product_description):
            percent_string_3 = percent_match_3.group(2).replace(',', '.')
            caffeine_content = 1000 * float(percent_string_3)
        if amount_match := amount_pattern.search(product_description):
            caffeine_amount = float(amount_match.group(1).replace(',', '.'))

        if caffeine_content > 0:
            return caffeine_content, caffeine_amount
        
    except TimeoutException:
        product_description = ''

    try:
        wait = WebDriverWait(driver, 3)
        product_info_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//h2[text()='Tuotetiedot']")))
        product_info_header.click()
        time.sleep(random.uniform(1, 2))
        try:
            wait = WebDriverWait(driver, 3)
            product_details = wait.until(EC.presence_of_element_located((By.XPATH, "//h3[text()='Ainesosat']//following-sibling::p"))).text

            if mg_100ml_match := mg_100ml_pattern.search(product_details):
                caffeine_content = float(mg_100ml_match.group(2)) 
            elif mg_100ml_match_2 := mg_100ml_pattern_2.search(product_details):
                caffeine_content = float(mg_100ml_match_2.group(1)) 
            elif mg_100ml_match_3 := mg_100ml_pattern_3.search(product_details):
                caffeine_content = float(mg_100ml_match_3.group(2))
            elif mg_100ml_match_4 := mg_100ml_pattern_4.search(product_details):
                caffeine_content = float(mg_100ml_match_4.group(2))
            elif mg_l_match := mg_l_pattern.search(product_details):
                caffeine_content = float(mg_l_match.group(2)) / 10
            elif mg_l_match_2 := mg_l_pattern_2.search(product_details):
                caffeine_content = float(mg_l_match_2.group(1)) / 10
            elif mg_l_match_3 := mg_l_pattern_3.search(product_details):
                caffeine_content = float(mg_l_match_3.group(2)) / 10
            elif percent_match := percent_pattern.search(product_details):
                percent_string = percent_match.group(2).replace(',', '.')
                caffeine_content = 1000 * float(percent_string)
            elif percent_match_2 := percent_pattern_2.search(product_details):
                percent_string_2 = percent_match_2.group(1).replace(',', '.')
                caffeine_content = 1000 * float(percent_string_2)
            elif percent_match_3 := percent_pattern_3.search(product_details):
                percent_string_3 = percent_match_3.group(2).replace(',', '.')
                caffeine_content = 1000 * float(percent_string_3)
            if amount_match := amount_pattern.search(product_details):
                caffeine_amount = float(amount_match.group(1).replace(',', '.'))
            
            if caffeine_content > 0:
                product_info_header.click()
                time.sleep(random.uniform(1, 2))
                return caffeine_content, caffeine_amount
            
        except TimeoutException:
            product_details = ''
        try:
            wait = WebDriverWait(driver, 3)
            product_instructions = wait.until(EC.presence_of_element_located((By.XPATH, "//h3[text()='Säilytys- ja käyttöohjeet']//following-sibling::div"))).text

            if mg_100ml_match := mg_100ml_pattern.search(product_instructions):
                caffeine_content = float(mg_100ml_match.group(2)) 
            elif mg_100ml_match_2 := mg_100ml_pattern_2.search(product_instructions):
                caffeine_content = float(mg_100ml_match_2.group(1)) 
            elif mg_100ml_match_3 := mg_100ml_pattern_3.search(product_instructions):
                caffeine_content = float(mg_100ml_match_3.group(2))
            elif mg_100ml_match_4 := mg_100ml_pattern_4.search(product_instructions):
                caffeine_content = float(mg_100ml_match_4.group(2))
            elif mg_l_match := mg_l_pattern.search(product_instructions):
                caffeine_content = float(mg_l_match.group(2)) / 10
            elif mg_l_match_2 := mg_l_pattern_2.search(product_instructions):
                caffeine_content = float(mg_l_match_2.group(1)) / 10
            elif mg_l_match_3 := mg_l_pattern_3.search(product_instructions):
                caffeine_content = float(mg_l_match_3.group(2)) / 10
            elif percent_match := percent_pattern.search(product_instructions):
                percent_string = percent_match.group(2).replace(',', '.')
                caffeine_content = 1000 * float(percent_string)
            elif percent_match_2 := percent_pattern_2.search(product_instructions):
                percent_string_2 = percent_match_2.group(1).replace(',', '.')
                caffeine_content = 1000 * float(percent_string_2)
            elif percent_match_3 := percent_pattern_3.search(product_instructions):
                percent_string_3 = percent_match_3.group(2).replace(',', '.')
                caffeine_content = 1000 * float(percent_string_3)
            if amount_match := amount_pattern.search(product_instructions):
                caffeine_amount = float(amount_match.group(1).replace(',', '.'))

            if caffeine_content > 0:
                product_info_header.click()
                time.sleep(random.uniform(1, 2))
                return caffeine_content, caffeine_amount
            
        except TimeoutException:
            product_instructions = ''
        product_info_header.click()
        time.sleep(random.uniform(1, 2))
                                                                        
    except TimeoutException:
        product_details = ''
        product_instructions = ''

    return caffeine_content, caffeine_amount

def preprocess_and_save():
    try:
        with open('nutritional_content_data.json', 'r') as file:
            nutritional_content_dict = json.load(file)
    except IOError:
        print("Could not open nutritional content data file. Using empty dictionary.")
        nutritional_content_dict = {}

    try:
        with open('product_prices_data.json', 'r') as file2:
            product_price_dict = json.load(file2)
    except IOError:
        print("Could not open product price data file. Using empty dictionary.")
        product_price_dict = {}

    try:
        with open('discounted_product_prices_data.json', 'r') as file3:
            discounted_product_price_dict = json.load(file3)
    except IOError:
        print("Could not open product price data file. Using empty dictionary.")
        discounted_product_price_dict = {}

    for ean in list(product_price_dict.keys()):
        details = product_price_dict[ean]
        discount_until = details.get('Discount valid until', 'Unknown')
        if discount_until != "Unknown":
            #if product_price_dict[ean]['Discount valid until'].split('.')[2] == '':
                #product_price_dict.pop(ean)
                #continue
            if date(today.year, today.month, today.day) <= date(int(product_price_dict[ean]['Discount valid until'].split('.')[2]), int(product_price_dict[ean]['Discount valid until'].split('.')[1]), int(product_price_dict[ean]['Discount valid until'].split('.')[0])):   
                continue
            else:
                product_price_dict.pop(ean)
                print('Deleted', ean, "from products JSON due to expired discount")          

    product_price_json = json.dumps(product_price_dict, indent=4)
    with open("product_prices_data.json", "w") as outfile:
        outfile.write(product_price_json)

    for ean in list(discounted_product_price_dict.keys()):
        details = discounted_product_price_dict[ean]
        discount_until = details.get('Discount valid until', "Unknown")
        if discount_until != "Unknown":
            #if discounted_product_price_dict[ean]['Discount valid until'].split('.')[2] == '':
                #discounted_product_price_dict.pop(ean)
                #continue
            if date(today.year, today.month, today.day) <= date(int(discounted_product_price_dict[ean]['Discount valid until'].split('.')[2]), int(discounted_product_price_dict[ean]['Discount valid until'].split('.')[1]), int(discounted_product_price_dict[ean]['Discount valid until'].split('.')[0])):   
                continue
            else:
                discounted_product_price_dict.pop(ean)
                print('Deleted', ean, "from discounted products JSON due to expired discount")
        else:
            discounted_product_price_dict.pop(ean)
            print('Deleted', ean, "from discounted products JSON due to missing discount date")

    discounted_product_price_json = json.dumps(discounted_product_price_dict, indent=4)
    with open("discounted_product_prices_data.json", "w") as outfile2:
        outfile2.write(discounted_product_price_json)

    return nutritional_content_dict, product_price_dict, discounted_product_price_dict

def get_stores_list(search_string):
    logpath=get_logpath()
    user_agents = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (Linux; Android 14; SM-G991B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Mobile Safari/537.36",
    "Mozilla/5.0 (iPad; CPU OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/124.0.2478.80",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Linux; Android 13; Pixel 6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; Android 14; Samsung Galaxy S22) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Mobile Safari/537.36",
    "Mozilla/5.0 (Windows NT 11.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Linux; Android 13; Xiaomi Mi 11) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Mobile Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36"
    ]

    user_agent = random.choice(user_agents)

    service = get_webdriver_service(logpath=logpath)

    try:
        options = uc.ChromeOptions()
        options.binary_location = "/usr/bin/chromium"
        options.add_argument(f'--user-agent={user_agent}')
        options.add_argument("--headless=new")
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--remote-debugging-port=9222")

        driver = uc.Chrome(options=options, service=service)
    except Exception as e:
        options = uc.ChromeOptions()
        options.binary_location = "/usr/bin/chromium"
        options.add_argument(f'--user-agent={user_agent}')
        options.add_argument("--headless=new")
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--remote-debugging-port=9222")

        main_version_string = re.search(r"Current browser version is (\d+\.\d+\.\d+)", str(e)).group(1)
        main_version = int(main_version_string.split(".")[0])

        driver = uc.Chrome(options=options, service=service, version_main=main_version)
    
    driver.get(f"https://www.k-ruoka.fi/?kaupat&kauppahaku={search_string}")
    time.sleep(random.uniform(8, 10))

    wait = WebDriverWait(driver, 10)
    try:
        accept_cookies = wait.until(EC.element_to_be_clickable((By.XPATH, f"//button[@id='onetrust-accept-btn-handler']")))
        # Click accept cookies button
        accept_cookies.click()
        time.sleep(random.uniform(1, 2))
    except TimeoutException:
        print('Prompt to accept cookies did not pop up')
    

    wait = WebDriverWait(driver, 20)
    try:
        search_summary_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@data-component='search-summary']")))
        search_summary_string = search_summary_element.text if search_summary_element.text else "0"
        if search_summary_string == "0":
            print("No stores found in the specified locations")
            return[]
        number_of_stores = int(''.join(filter(str.isdigit, search_summary_string)))
    except TimeoutException:
        print("Store search summary element not found or not visible")
        driver.save_screenshot('screenshot.png')
        st.image("screenshot.png", caption="Screen")
        return []

    store_list_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-component='store-list']")))
    stores = store_list_element.find_elements(By.XPATH, ".//li[@data-component='store-list-item']")
    new_number_of_stores = len(stores)

    while True:
        if new_number_of_stores == number_of_stores:
            break
        store_list_container = driver.find_element(By.XPATH, "//div[starts-with(@class, 'StoreSelector__StyledVerticalScrollAwareContainer')]")
        store_list_container.click()
        time.sleep(random.uniform(0, 1))
        store_list_container.send_keys(Keys.PAGE_DOWN)

        time.sleep(random.uniform(2, 3))

        stores = driver.find_elements(By.XPATH, "//ul[@data-component='store-list']//li[@data-component='store-list-item']")
        new_number_of_stores = len(stores)

    store_options = []
    for store in stores:
        try:
            store_title = store.find_element(By.XPATH, ".//h3[@data-testid='store-title']").text
            store_options.append(store_title)
        except NoSuchElementException:
            print("Store title element not found for one of the stores, skipping this store")
            continue

    return store_options, driver

def run_scraper(selected_categories, selected_stores, nutritional_content_dict, product_price_dict, discounted_product_price_dict, search_string, driver, only_discounted_products):

    start0 = time.perf_counter()
    elapsed_time_text = st.empty()
    
    store_progress_text = st.empty()
    store_progress_bar = st.progress(0)
    category_progress_text = st.empty()
    category_progress_bar = st.progress(0)
    product_progress_text = st.empty()
    product_progress_bar = st.progress(0)
    url_progress_text = st.empty()
    url_progress_bar = st.progress(0)

    counter = 0
    store_counter = 0
    total_stores = len(selected_stores)

    wait = WebDriverWait(driver, 10)
    try:
        search_summary_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@data-component='search-summary']")))
        search_summary_string = search_summary_element.text if search_summary_element.text else "0"
        if search_summary_string == "0":
            print("No stores found in the specified locations")
            driver.quit()
            exit()
        number_of_stores = int(''.join(filter(str.isdigit, search_summary_string)))
    except TimeoutException:
        print("Store search summary element not found or not visible")
        driver.quit()
        exit()

    while counter < number_of_stores:
        if store_counter >= total_stores:
            break
        elapsed = time.perf_counter() - start0
        elapsed_time_text.text(f"Total time elapsed: {round(elapsed/60, 2)} minutes")
        start = time.perf_counter()
        wait = WebDriverWait(driver, 30)
        try:
            store_list_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-component='store-list']")))
            stores = store_list_element.find_elements(By.XPATH, ".//li[@data-component='store-list-item']")
            new_number_of_stores = len(stores)

            while True:
                if new_number_of_stores - 1 >= counter:
                    break
                store_list_container = driver.find_element(By.XPATH, "//div[starts-with(@class, 'StoreSelector__StyledVerticalScrollAwareContainer')]")
                store_list_container.click()
                time.sleep(random.uniform(0, 1))
                store_list_container.send_keys(Keys.PAGE_DOWN)

                time.sleep(random.uniform(2, 3))

                stores = driver.find_elements(By.XPATH, "//ul[@data-component='store-list']//li[@data-component='store-list-item']")
                new_number_of_stores = len(stores)

            store_location = stores[counter].find_element(By.XPATH, ".//div[@data-testid='store-location']").text
            store_name = stores[counter].get_attribute("data-store")
            store_title = stores[counter].find_element(By.XPATH, ".//h3[@data-testid='store-title']").text

            if store_title not in selected_stores: 
                print(f"Skipping store {counter + 1} - {store_title} - {store_location} as it is not in the selected stores list")
                counter += 1
                continue

            store = stores[counter].find_element(By.XPATH, f".//button[@data-select-store='{store_name}']")
            driver.execute_script("arguments[0].scrollIntoView()", store)

            time.sleep(random.uniform(1, 2))
            store.click()
            time.sleep(random.uniform(4, 6))

            counter += 1

        except TimeoutException:
            print("Store list element not found or not visible")
            counter += 1
            continue

        store_progress_text.text(f"Scraping store {store_counter+1} of {total_stores} - {store_title} ({store_location})")
        
        product_urls = {}
        category_progress_bar.progress(0)
        total_categories = len(selected_categories)
        category_counter = 0

        for product_category in selected_categories:
            elapsed = time.perf_counter() - start0
            elapsed_time_text.text(f"Total time elapsed: {round(elapsed/60, 2)} minutes")

            category_progress_text.text(f"Scraping category {category_counter+1} of {total_categories} - {inverted_category_options[product_category]}")
            driver.get(f"https://www.k-ruoka.fi/kauppa/tuotehaku/{product_category}")
            time.sleep(random.uniform(1, 2))
            
            wait = WebDriverWait(driver, 20)
            try:
                products_list = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-testid='product-search-results']")))
                product_cards = products_list.find_elements(By.XPATH, ".//li[@data-testid='product-card']")
                last_list_len = len(product_cards)
                while True:
                    elapsed = time.perf_counter() - start0
                    elapsed_time_text.text(f"Total time elapsed: {round(elapsed/60, 2)} minutes")
                    driver.execute_script("arguments[0].scrollIntoView()", product_cards[-1])
                    time.sleep(random.uniform(2, 3))
                    product_cards = driver.find_elements(By.XPATH, "//ul[@data-testid='product-search-results']//li[@data-testid='product-card']")
                    new_list_len = len(product_cards)
                    if new_list_len != last_list_len:
                        last_list_len = new_list_len
                    else:
                        time.sleep(random.uniform(1, 2))
                        product_cards = driver.find_elements(By.XPATH, "//ul[@data-testid='product-search-results']//li[@data-testid='product-card']")
                        new_list_len = len(product_cards)
                        if new_list_len == last_list_len:
                            break
                        last_list_len = new_list_len
            except TimeoutException:
                print("No products found in", store_title, "for category", product_category)
                category_counter += 1
                category_progress_bar.progress(int((category_counter / total_categories) * 100))
                continue
            
            wait = WebDriverWait(driver, 3)
            try:
                products_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-testid='product-search-results']")))
                product_cards = products_element.find_elements(By.XPATH, ".//div[starts-with(@class, 'ProductCardDiscount__Badge') or @data-testid='product-normal-price']//ancestor::li") if only_discounted_products else products_element.find_elements(By.XPATH, ".//li[@data-testid='product-card']")
            except TimeoutException:
                print("No products found in", store_title, "for category", product_category, "after scrolling")
                category_counter += 1
                category_progress_bar.progress(int((category_counter / total_categories) * 100))
                continue
            
            total_products = len(product_cards)
            product_progress_bar.progress(0)
            product_counter = 0
            scrape_start = time.perf_counter()
            scrape_end = time.perf_counter()

            elapsed = time.perf_counter() - start0
            elapsed_time_text.text(f"Total time elapsed: {round(elapsed/60, 2)} minutes")
            
            for card in product_cards:
                product_progress_text.text(f"Scraping product card {product_counter+1} of {total_products}")
                url_elements = card.find_elements(By.XPATH, ".//a[@data-testid='product-link']")
                if len(url_elements) > 0:
                    try:
                        url = url_elements[0].get_attribute("href")
                        product_name = url_elements[0].text
                        size = extract_size_in_kg(product_name)
                        ean_code_string = card.get_attribute("data-product-id")
                        if ean_code_string != None:
                            hyphen_index = ean_code_string.find("-")
                        else:
                            print("EAN code is an empty string for", product_name)
                            product_urls.update({url: 'Unknown'})
                            product_counter += 1
                            product_progress_bar.progress(int((product_counter / total_products) * 100))
                            continue
                        ean_code = ean_code_string[:hyphen_index] if hyphen_index != -1 else ean_code_string

                        if only_discounted_products:
                            discount = 'Yes'
                        else:
                            discount_badge_elements = card.find_elements(By.XPATH, ".//div[starts-with(@class, 'ProductCardDiscount__Text')]")
                            normal_price_elements = card.find_elements(By.XPATH, ".//div[@data-testid='product-normal-price']")
                            if len(discount_badge_elements) > 0 or len(normal_price_elements) > 0:
                                discount = 'Yes'
                            else:
                                discount = 'No'
                        
                    except Exception as e:
                        print("Could not get product url or EAN code for", card.text, e)
                        product_counter += 1
                        product_progress_bar.progress(int((product_counter / total_products) * 100))
                        continue

                    if nutritional_content_dict.get(ean_code, 'Unknown') != 'Unknown':
                        try:
                            unit_price_string = card.find_element(By.XPATH, ".//div[@data-testid='product-unit-price']").text
                            backslash_index = unit_price_string.find("/")
                            if backslash_index != -1:
                                unit_type = unit_price_string[backslash_index+1:]
                                unit_price = unit_price_string[:backslash_index]
                                unit_price = float(unit_price.replace(',', '.')) if len(unit_price) != 0 else 999.999
                            else:
                                search_res = re.search(r'(\d+(?:[.,]\d+)?)', unit_price_string)
                                unit_price = float(search_res.group().replace(',', '.')) if search_res else 999.999
                                unit_type = 'kg' if 'kg' in unit_price_string else 'Unknown'

                            # Check if unit_price is suspiciously low
                            if unit_price <= 0.2:
                                raise ValueError("Unit price seems too low, fallback to alternative method")

                        except (NoSuchElementException, ValueError):
                            try:
                                price_element_integer = card.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__IntegerPart')]")
                                unit_price_integer = price_element_integer.text

                                price_element_decimal = card.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__DecimalPart')]")
                                unit_price_decimal = price_element_decimal.text

                                unit_element = card.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__Extra')]")
                                unit_type = unit_element.text.replace('/', '')

                                unit_price = float(unit_price_integer + '.' + unit_price_decimal) if len(unit_price_integer) > 0 and len(unit_price_decimal) > 0 else 999.999

                            except Exception as e:
                                print("Could not get product price for", product_name, e)
                                unit_price = 999.999
                                unit_type = 'Unknown'
                                product_urls.update({url: ean_code})

                        if unit_type == 'kpl':
                            kg_price = unit_price / size if size != 'Unknown' and size != 0 else 999.999
                        elif unit_type == 'kg' or unit_type == 'l':
                            kg_price = unit_price
                        else:
                            kg_price = 999.999

                        if product_price_dict.get(ean_code, "Unknown") != "Unknown":
                            if discount == 'Yes':
                                if product_price_dict[ean_code].get('Discount valid until', 'Unknown') == 'Unknown':
                                    product_price_dict[ean_code]['Price per Unit'] = unit_price
                                    product_price_dict[ean_code]['Price per kg'] = kg_price
                                    product_price_dict[ean_code]['Unit'] = unit_type
                                    product_price_dict[ean_code]['Size (kg)'] = size
                                    product_price_dict[ean_code]['Store'] = store_name
                                    product_urls.update({url: ean_code})
                                else:
                                    if product_price_dict[ean_code].get('Price per Unit', 'Unknown') != 'Unknown':
                                        if product_price_dict[ean_code]['Price per Unit'] > unit_price:
                                            product_price_dict[ean_code]['Price per Unit'] = unit_price
                                            product_price_dict[ean_code]['Price per kg'] = kg_price
                                            product_price_dict[ean_code]['Unit'] = unit_type
                                            product_price_dict[ean_code]['Size (kg)'] = size
                                            product_price_dict[ean_code]['Store'] = store_name
                                            product_urls.update({url: ean_code})
                                        elif product_price_dict[ean_code]['Price per Unit'] < unit_price:
                                            if product_price_dict[ean_code].get('Store') == store_name:
                                                product_price_dict[ean_code]['Price per Unit'] = unit_price
                                                product_price_dict[ean_code]['Price per kg'] = kg_price
                                                product_urls.update({url: ean_code})
                                    else:
                                        product_price_dict[ean_code]['Price per Unit'] = unit_price
                                        product_price_dict[ean_code]['Price per kg'] = kg_price
                                        product_price_dict[ean_code]['Unit'] = unit_type
                                        product_price_dict[ean_code]['Size (kg)'] = size
                                        product_price_dict[ean_code]['Store'] = store_name
                                        product_urls.update({url: ean_code})
                            else:
                                if product_price_dict[ean_code].get('Price per Unit', 'Unknown') != 'Unknown':
                                    if product_price_dict[ean_code]['Price per Unit'] > unit_price:
                                        product_price_dict[ean_code]['Price per Unit'] = unit_price
                                        product_price_dict[ean_code]['Price per kg'] = kg_price
                                        product_price_dict[ean_code]['Unit'] = unit_type
                                        product_price_dict[ean_code]['Size (kg)'] = size
                                        product_price_dict[ean_code]['Store'] = store_name
                                    else:
                                        if product_price_dict[ean_code].get('Store') == store_name:
                                            product_price_dict[ean_code]['Price per Unit'] = unit_price
                                            product_price_dict[ean_code]['Price per kg'] = kg_price
                                else:
                                    product_price_dict[ean_code]['Price per Unit'] = unit_price
                                    product_price_dict[ean_code]['Price per kg'] = kg_price
                                    product_price_dict[ean_code]['Unit'] = unit_type
                                    product_price_dict[ean_code]['Size (kg)'] = size
                                    product_price_dict[ean_code]['Store'] = store_name
                                    product_urls.update({url: ean_code})

                        else:
                            product_price_dict[ean_code] = {}
                            product_price_dict[ean_code]['Price per Unit'] = unit_price
                            product_price_dict[ean_code]['Price per kg'] = kg_price
                            product_price_dict[ean_code]['Unit'] = unit_type
                            product_price_dict[ean_code]['Size (kg)'] = size
                            product_price_dict[ean_code]['Store'] = store_name

                            if discount == 'Yes':
                                if product_price_dict[ean_code].get('Discount valid until', 'Unknown') == 'Unknown':
                                    product_urls.update({url: ean_code})

                        product_price_dict[ean_code].update(nutritional_content_dict[ean_code])

                    else:
                        product_urls.update({url: ean_code})
                        product_price_dict[ean_code] = {}
                        product_price_dict[ean_code]['Name'] = product_name
                        product_price_dict[ean_code]['Size (kg)'] = size
                        product_price_dict[ean_code]['Store'] = store_name

                    if discount == 'Yes':
                        discounted_product_price_dict[ean_code] = product_price_dict[ean_code]
                    
                else:
                    try:
                        driver.execute_script("arguments[0].scrollIntoView()", card)
                        nayta_tuotteet_button = card.find_element(By.XPATH, ".//button[@aria-label='Näytä tuotteet']")
                        nayta_tuotteet_button.click()
                        time.sleep(random.uniform(1, 2))
                    except Exception as e:
                        try:
                            time.sleep(random.uniform(1, 2))
                            nayta_tuotteet_button = card.find_element(By.XPATH, ".//button[@aria-label='Näytä tuotteet']")
                            driver.execute_script("arguments[0].scrollIntoView()", nayta_tuotteet_button)
                            time.sleep(random.uniform(1, 2))
                            nayta_tuotteet_button.click()
                            time.sleep(random.uniform(1, 2))
                        except Exception as e:
                            try:
                                driver.find_element(By.XPATH, "//button[@title='Sulje']").click()
                                time.sleep(random.uniform(2, 3))

                                nayta_tuotteet_button = card.find_element(By.XPATH, ".//button[@aria-label='Näytä tuotteet']")
                                driver.execute_script("arguments[0].scrollIntoView()", nayta_tuotteet_button)
                                time.sleep(random.uniform(2, 3))
                                nayta_tuotteet_button.click()
                                
                            except Exception as e:
                                print("Could not click 'Näytä tuotteet' button for", card.text, e)
                                product_counter += 1
                                product_progress_bar.progress(int((product_counter / total_products) * 100))
                                continue

                    wait = WebDriverWait(driver, 5)
                    try:
                        products_list_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-testid='offer-products']")))
                        product_elements = products_list_element.find_elements(By.XPATH, ".//li[@data-testid='product-card']")

                    except TimeoutException:
                        print("Product elements not found or not visible for", card.text)
                        product_counter += 1
                        product_progress_bar.progress(int((product_counter / total_products) * 100))
                        continue

                    for product in product_elements:
                        try:
                            url = product.find_element(By.XPATH, ".//a[@data-testid='product-link']").get_attribute("href")
                            product_name = product.find_element(By.XPATH, ".//a[@data-testid='product-link']").text
                            size = extract_size_in_kg(product_name)
                            ean_code_string = product.get_attribute("data-product-id")
                            if ean_code_string != None:
                                hyphen_index = ean_code_string.find("-")
                            else:
                                print("EAN code is an empty string for", product_name)
                                product_urls.update({url: 'Unknown'})
                                continue
                            ean_code = ean_code_string[:hyphen_index] if hyphen_index != -1 else ean_code_string

                            if only_discounted_products:
                                discount = 'Yes'
                            else:
                                discount_badge_elements = product.find_elements(By.XPATH, ".//div[starts-with(@class, 'ProductCardDiscount__Text')]")
                                normal_price_elements = product.find_elements(By.XPATH, ".//div[@data-testid='product-normal-price']")
                                if len(discount_badge_elements) > 0 or len(normal_price_elements) > 0:
                                    discount = 'Yes'
                                else:
                                    discount = 'No'

                        except Exception as e:
                            print("Could not get product url or EAN code for", product.text, e)
                            continue

                        if nutritional_content_dict.get(ean_code, 'Unknown') != 'Unknown':
                            try:
                                unit_price_string = product.find_element(By.XPATH, ".//div[@data-testid='product-unit-price']").text
                                backslash_index = unit_price_string.find("/")
                                if backslash_index != -1:
                                    unit_type = unit_price_string[backslash_index+1:]
                                    unit_price = unit_price_string[:backslash_index]
                                    unit_price = float(unit_price.replace(',', '.')) if len(unit_price) != 0 else 999.999
                                else:
                                    search_res = re.search(r'(\d+(?:[.,]\d+)?)', unit_price_string)
                                    unit_price = float(search_res.group().replace(',', '.')) if search_res else 999.999
                                    unit_type = 'kg' if 'kg' in unit_price_string else 'Unknown'

                                # Check if unit_price is suspiciously low
                                if unit_price <= 0.2:
                                    raise ValueError("Unit price seems too low, fallback to alternative method")

                            except (NoSuchElementException, ValueError):
                                try:
                                    price_element_integer = product.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__IntegerPart')]")
                                    unit_price_integer = price_element_integer.text

                                    price_element_decimal = product.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__DecimalPart')]")
                                    unit_price_decimal = price_element_decimal.text

                                    unit_element = product.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__Extra')]")
                                    unit_type = unit_element.text.replace('/', '')

                                    unit_price = float(unit_price_integer + '.' + unit_price_decimal) if len(unit_price_integer) > 0 and len(unit_price_decimal) > 0 else 999.999

                                except Exception as e:
                                    print("Could not get product price for", ean_code, product_name, e)
                                    unit_price = 999.999
                                    unit_type = 'Unknown'
                                    product_urls.update({url: ean_code})

                            if unit_type == 'kpl':
                                kg_price = unit_price / size if size != 'Unknown' and size != 0 else 999.999
                            elif unit_type == 'kg' or unit_type == 'l':
                                kg_price = unit_price
                            else:
                                kg_price = 999.999

                            if product_price_dict.get(ean_code, "Unknown") != "Unknown":
                                if product_price_dict[ean_code].get('Store') == store_name:
                                    product_price_dict[ean_code]['Price per Unit'] = unit_price
                                    product_price_dict[ean_code]['Price per kg'] = kg_price
                                if discount == 'Yes':
                                    if product_price_dict[ean_code].get('Discount valid until', 'Unknown') == 'Unknown':
                                        product_urls.update({url: ean_code})
                                    else:
                                        if product_price_dict[ean_code].get('Price per Unit', 'Unknown') != 'Unknown':
                                            if product_price_dict[ean_code]['Price per Unit'] > unit_price:
                                                product_price_dict[ean_code]['Price per Unit'] = unit_price
                                                product_price_dict[ean_code]['Price per kg'] = kg_price
                                                product_price_dict[ean_code]['Unit'] = unit_type
                                                product_price_dict[ean_code]['Size (kg)'] = size
                                                product_price_dict[ean_code]['Store'] = store_name

                                        else:
                                            product_price_dict[ean_code]['Price per Unit'] = unit_price
                                            product_price_dict[ean_code]['Price per kg'] = kg_price
                                            product_price_dict[ean_code]['Unit'] = unit_type
                                            product_price_dict[ean_code]['Size (kg)'] = size
                                            product_price_dict[ean_code]['Store'] = store_name
                                            product_urls.update({url: ean_code})
                                else:
                                    if product_price_dict[ean_code].get('Price per Unit', 'Unknown') != 'Unknown':
                                        if product_price_dict[ean_code]['Price per Unit'] > unit_price:
                                            product_price_dict[ean_code]['Price per Unit'] = unit_price
                                            product_price_dict[ean_code]['Price per kg'] = kg_price
                                            product_price_dict[ean_code]['Unit'] = unit_type
                                            product_price_dict[ean_code]['Size (kg)'] = size
                                            product_price_dict[ean_code]['Store'] = store_name
                                    else:
                                        product_price_dict[ean_code]['Price per Unit'] = unit_price
                                        product_price_dict[ean_code]['Price per kg'] = kg_price
                                        product_price_dict[ean_code]['Unit'] = unit_type
                                        product_price_dict[ean_code]['Size (kg)'] = size
                                        product_price_dict[ean_code]['Store'] = store_name
                                        product_urls.update({url: ean_code})

                            else:
                                product_price_dict[ean_code] = {}
                                product_price_dict[ean_code]['Price per Unit'] = unit_price
                                product_price_dict[ean_code]['Price per kg'] = kg_price
                                product_price_dict[ean_code]['Unit'] = unit_type
                                product_price_dict[ean_code]['Size (kg)'] = size
                                product_price_dict[ean_code]['Store'] = store_name

                            product_price_dict[ean_code].update(nutritional_content_dict[ean_code])

                        else:
                            product_urls.update({url: ean_code})
                            product_price_dict[ean_code] = {}
                            product_price_dict[ean_code]['Name'] = product_name
                            product_price_dict[ean_code]['Size (kg)'] = size
                            product_price_dict[ean_code]['Store'] = store_name

                        if discount == 'Yes':
                            discounted_product_price_dict[ean_code] = product_price_dict[ean_code]

                    driver.find_element(By.XPATH, "//button[@title='Sulje']").click()
                    time.sleep(random.uniform(2, 3))

                product_counter += 1
                product_progress_bar.progress(int((product_counter / total_products) * 100))
                scrape_end = time.perf_counter()
                
            product_progress_text.text(f"Finished scraping {total_products} product cards in {round((scrape_end - scrape_start), 2)} seconds")
            category_counter += 1
            category_progress_text.text(f"Finished scraping category {category_counter}/{total_categories} - {inverted_category_options[product_category]}")
            category_progress_bar.progress(int((category_counter / total_categories) * 100))

        url_progress_bar.progress(0)
        total_urls = len(product_urls)
        url_counter = 0
        url_start = time.perf_counter()

        for url, ean in product_urls.items():
            elapsed = time.perf_counter() - start0
            elapsed_time_text.text(f"Total time elapsed: {round(elapsed/60, 2)} minutes")
            url_progress_text.text(f"Scraping URL {url_counter + 1} of {total_urls}")
            if url:
                try:
                    driver.get(url)
                    time.sleep(random.uniform(1, 2))
                except Exception as e:
                    print("Could not open link")
                    url_counter += 1
                    url_progress_bar.progress(int((url_counter / total_urls) * 100))
                    continue
            else:
                print("Invalid URL, skipping")
                url_counter += 1
                url_progress_bar.progress(int((url_counter / total_urls) * 100))
                time.sleep(random.uniform(1, 2))
                continue

            try:
                wait = WebDriverWait(driver, 5)
                header = wait.until(EC.visibility_of_element_located((By.XPATH, "//h1[@data-testid='product-name']")))
                product_name = header.text
                size = extract_size_in_kg(product_name)
                category_elements = driver.find_elements(By.XPATH, "//li[starts-with(@class, 'Breadcrumbs__BreadcrumbsItem')]")
                if len(category_elements) > 1:
                    category = category_elements[len(category_elements) - 2].text
                elif len(category_elements) == 1:
                    category = category_elements[0].text
                else:
                    print('Unknown category for', ean, product_name)
                    category = 'Unknown'
            except TimeoutException:
                print("Product name not found for", url)
                url_counter += 1
                url_progress_bar.progress(int((url_counter / total_urls) * 100))
                continue

            if ean == 'Unknown' or ean == '':
                try:
                    wait = WebDriverWait(driver, 5)
                    product_info_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//h2[text()='Tuotetiedot']")))
                    product_info_header.click()
                    time.sleep(random.uniform(1, 2))
                except TimeoutException:
                    print("Product info header not found for", product_name)
                    url_counter += 1
                    url_progress_bar.progress(int((url_counter / total_urls) * 100))
                    continue

                try:
                    wait = WebDriverWait(driver, 5)
                    ean_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//h3[text()='EAN-koodi']//following-sibling::p")))
                    ean_code = ean_element.text
                    product_info_header.click()
                    time.sleep(random.uniform(1, 2))
                except TimeoutException:
                    print("EAN code not found for", product_name)
                    url_counter += 1
                    url_progress_bar.progress(int((url_counter / total_urls) * 100))
                    continue
            else:
                ean_code = ean

            vegan = 'No'
            gluten_free = 'No'
            lactose_free = 'No'
            sydanmerkki = 'No'
            hyvaa_suomesta = 'No'
            luomu = 'No'

            try:
                attribute_elements = driver.find_elements(By.XPATH, "//div[starts-with(@class, 'NutritionalAttributeHighlights__Symbol')]")
                for attribute in attribute_elements:
                    if attribute.text == 'V':
                        vegan = 'Yes'
                    elif attribute.text == 'G':
                        gluten_free = 'Yes'
                    elif attribute.text == 'L':
                        lactose_free = 'Yes'
                    elif attribute.text == 'LU':
                        luomu = 'Yes'
            except Exception as e:
                print("Could not find nutritional attributes for", product_name, e)
            
            try:
                responsibility_elements = driver.find_elements(By.XPATH, "//div[starts-with(@class, 'ResponsibilityHighlights__Container')]//img")
                for responsibility in responsibility_elements:
                    if responsibility.get_attribute("alt") == 'Sydänmerkki':
                        sydanmerkki = 'Yes'
                    elif responsibility.get_attribute("alt") == 'Hyvää Suomesta':
                        hyvaa_suomesta = 'Yes'
            except Exception as e:
                print("Could not find responsibility attributes for", product_name, e)
            
            try:
                wait = WebDriverWait(driver, 3)
                price_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//h1[@data-testid='product-name']//following-sibling::div//div[@data-testid='product-unit-price']")))
                unit_price_string = price_element.text
                backslash_index = unit_price_string.find("/")

                if backslash_index != -1:
                    unit_type = unit_price_string[backslash_index+1:]
                    unit_price = unit_price_string[:backslash_index]
                    unit_price = float(unit_price.replace(',', '.')) if len(unit_price) > 0 else 999.999
                else:
                    search_res = re.search(r'(\d+(?:[.,]\d+)?)', unit_price_string)
                    unit_price = float(search_res.group().replace(',', '.')) if search_res else 999.999
                    unit_type = 'kg' if 'kg' in unit_price_string else 'Unknown'
                
                # Check if unit_price is suspiciously low
                if unit_price <= 0.1:
                    raise ValueError("Unit price seems too low, fallback to alternative method")

            except (TimeoutException, ValueError):
                try:
                    price_element_integer = driver.find_element(By.XPATH, "//div[@data-testid='product-details-sidebar']//div[starts-with(@class, 'ProductPrice__IntegerPart')]")
                    unit_price_integer = price_element_integer.text
                    price_element_decimal = driver.find_element(By.XPATH, "//div[@data-testid='product-details-sidebar']//div[starts-with(@class, 'ProductPrice__DecimalPart')]")
                    unit_price_decimal = price_element_decimal.text
                    unit_element = driver.find_element(By.XPATH, "//div[@data-testid='product-details-sidebar']//div[starts-with(@class, 'ProductPrice__Extra')]")
                    unit_type = unit_element.text.replace('/', '')
                    unit_price = float(unit_price_integer + '.' + unit_price_decimal) if len(unit_price_integer) > 0 and len(unit_price_decimal) > 0 else 999.999
                except Exception as e:
                    print("Could not get product price for", product_name, e)
                    unit_price = 999.999
                    unit_type = 'Unknown'
                
            if unit_type == 'kpl':
                kg_price = unit_price / size if size != 'Unknown' and size != 0 else 999.999
            elif unit_type == 'kg' or unit_type == 'l':
                kg_price = unit_price
            else:
                kg_price = 999.999

            valid_during_elements = driver.find_elements(By.XPATH, "//h1[@data-testid='product-name']//following::div[starts-with(@class, 'ProductSidebarContent__Info') and contains(translate(text(), 'VOIMASSA', 'voimassa'), 'voimassa')]")
            if len(valid_during_elements) > 0:
                discount = 'Yes'
                valid_during_string = valid_during_elements[0].text
                
                search_res = re.findall(r'(\d{1,2})\.(\d{1,2})(?:\.(\d{4}))?', valid_during_string)
                valid_starting_string = '.'.join([part for part in search_res[0] if part])
                valid_until_string = '.'.join([part for part in search_res[1] if part])

                valid_starting = valid_starting_string if len(valid_starting_string.split('.')) == 3 else valid_starting_string + '.' + str(today.year) if len(valid_starting_string.split('.')) == 2 else 'Unknown'

                valid_until = valid_until_string if len(valid_until_string.split('.')) == 3 else valid_until_string + '.' + str(today.year) if len(valid_until_string.split('.')) == 2 else 'Unknown'
            
            else:
                discount = 'No'
                valid_starting = None
                valid_until = None

            if product_price_dict.get(ean_code, 'Unknown') != 'Unknown':
                if product_price_dict[ean_code].get('Store', 'Unknown') == store_name:
                    product_price_dict[ean_code]['Price per Unit'] = unit_price
                    product_price_dict[ean_code]['Price per kg'] = kg_price
                    product_price_dict[ean_code]['Unit'] = unit_type
                
                else:
                    if product_price_dict[ean_code].get('Price per Unit', 'Unknown') != 'Unknown':
                        if product_price_dict[ean_code]['Price per Unit'] > unit_price:
                            product_price_dict[ean_code]['Price per Unit'] = unit_price
                            product_price_dict[ean_code]['Price per kg'] = kg_price
                            product_price_dict[ean_code]['Unit'] = unit_type
                            product_price_dict[ean_code]['Size (kg)'] = size
                            product_price_dict[ean_code]['Store'] = store_name

                    else:
                        product_price_dict[ean_code]['Price per Unit'] = unit_price
                        product_price_dict[ean_code]['Price per kg'] = kg_price
                        product_price_dict[ean_code]['Unit'] = unit_type
                        product_price_dict[ean_code]['Size (kg)'] = size
                        product_price_dict[ean_code]['Store'] = store_name

            else:
                product_price_dict[ean_code] = {}
                product_price_dict[ean_code]['Name'] = product_name
                product_price_dict[ean_code]['Price per Unit'] = unit_price
                product_price_dict[ean_code]['Price per kg'] = kg_price
                product_price_dict[ean_code]['Unit'] = unit_type
                product_price_dict[ean_code]['Size (kg)'] = size
                product_price_dict[ean_code]['Store'] = store_name

            if discount == 'Yes':
                product_price_dict[ean_code]['Discount valid starting'] = valid_starting
                product_price_dict[ean_code]['Discount valid until'] = valid_until

            if nutritional_content_dict.get(ean_code, 'Unknown') != 'Unknown':
                nutritional_content_dict[ean_code]['Name'] = product_name
                nutritional_content_dict[ean_code]['Category'] = category
                nutritional_content_dict[ean_code]['Vegan'] = vegan if nutritional_content_dict[ean_code].get('Vegan', 'Unknown') == 'Unknown' or vegan == 'Yes' else nutritional_content_dict[ean_code]['Vegan']
                nutritional_content_dict[ean_code]['Gluten Free'] = gluten_free if nutritional_content_dict[ean_code].get('Gluten Free', 'Unknown') == 'Unknown' or gluten_free == 'Yes' else nutritional_content_dict[ean_code]['Gluten Free']
                nutritional_content_dict[ean_code]['Lactose Free'] = lactose_free if nutritional_content_dict[ean_code].get('Lactose Free', 'Unknown') == 'Unknown' or lactose_free == 'Yes' else nutritional_content_dict[ean_code]['Lactose Free']
                nutritional_content_dict[ean_code]['Organic'] = luomu if nutritional_content_dict[ean_code].get('Organic', 'Unknown') == 'Unknown' or luomu == 'Yes' else nutritional_content_dict[ean_code]['Organic']
                nutritional_content_dict[ean_code]['Sydänmerkki'] = sydanmerkki if nutritional_content_dict[ean_code].get('Sydänmerkki', 'Unknown') == 'Unknown' or sydanmerkki == 'Yes' else nutritional_content_dict[ean_code]['Sydänmerkki']
                nutritional_content_dict[ean_code]['Hyvää Suomesta'] = hyvaa_suomesta if nutritional_content_dict[ean_code].get('Hyvää Suomesta', 'Unknown') == 'Unknown' or hyvaa_suomesta == 'Yes' else nutritional_content_dict[ean_code]['Hyvää Suomesta']
                product_price_dict[ean_code].update(nutritional_content_dict[ean_code])
                time.sleep(random.uniform(1, 2))
            else:
                nutritional_content_dict[ean_code] = {}
                nutritional_content_dict[ean_code]['Name'] = product_name
                nutritional_content_dict[ean_code]['Category'] = category
                nutritional_content_dict[ean_code]['Vegan'] = vegan
                nutritional_content_dict[ean_code]['Gluten Free'] = gluten_free
                nutritional_content_dict[ean_code]['Lactose Free'] = lactose_free
                nutritional_content_dict[ean_code]['Organic'] = luomu
                nutritional_content_dict[ean_code]['Sydänmerkki'] = sydanmerkki
                nutritional_content_dict[ean_code]['Hyvää Suomesta'] = hyvaa_suomesta

                if category == 'Energiajuomat' or category == 'Urheiluvalmisteet' or category == 'Energia- ja urheilujuomat':
                    caffeine_content, caffeine_amount = get_caffeine_amount(driver)
                    nutritional_content_dict[ean_code]['Kofeiini (per 100ml)'] = caffeine_content
                    nutritional_content_dict[ean_code]['Kofeiini (per tuote)'] = caffeine_amount

                try:
                    wait = WebDriverWait(driver, 3)
                    nutritional_content_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//h2[text()='Ravintosisältö']")))
                    nutritional_content_header.click()
                    time.sleep(random.uniform(1, 2))
                except ElementClickInterceptedException:
                    nutritional_content_header = driver.find_element(By.XPATH, "//h2[text()='Ravintosisältö']")
                    driver.execute_script("arguments[0].scrollIntoView()", nutritional_content_header)
                    time.sleep(random.uniform(1, 2))
                    nutritional_content_header.click()
                    time.sleep(random.uniform(0, 1))
                
                except TimeoutException:
                    print("Nutritional content header not found for", product_name)
                    product_price_dict[ean_code].update(nutritional_content_dict[ean_code])
                    if discount == 'Yes':
                        discounted_product_price_dict[ean_code] = product_price_dict[ean_code]
                    url_counter += 1
                    url_progress_bar.progress(int((url_counter / total_urls) * 100))
                    continue

                keys_list = []
                values_list = []
                try:
                    wait = WebDriverWait(driver, 3)
                    table = wait.until(EC.visibility_of_element_located((By.XPATH, "//h2[(text()='Ravintosisältö')]//parent::button//following::table[starts-with(@class, 'NewNutritionalDetails__Table')]")))
                    unit_size = table.find_element(By.XPATH, ".//th[starts-with(@class, 'NewNutritionalDetails__NutritionContentTableColumnHeading')]//div").text
                    product_price_dict[ean_code]['Nutritional Value per'] = unit_size
                    keys = table.find_elements(By.XPATH, ".//tbody//th[starts-with(@class, 'NewNutritionalDetails')]")
                    values = table.find_elements(By.XPATH, ".//tbody//td[starts-with(@class, 'NewNutritionalDetails')][1]")

                    for key in keys:
                        keys_list.append(key.text.strip())

                    for value in values:
                        token = value.text.strip()
                        kcal_index = token.find("kcal")
                        kj_index = token.find("kJ")
                        backslash_index = token.find("/")

                        if kcal_index != -1:
                            if backslash_index != -1:
                                value_string = token[backslash_index+1:kcal_index].replace(',', '.').strip()
                                values_list.append(float(value_string) if len(value_string) != 0 else 0.0)
                            else:
                                if kj_index != -1:
                                    value_string = token[:kj_index].replace(',', '.').strip()
                                    values_list.append(float(value_string) * 0.2390057 if len(value_string) != 0 else 0.0)
                                else:
                                    value_string = token[:kcal_index].replace(',', '.').strip()
                                    values_list.append(float(value_string) if len(value_string) != 0 else 0.0)

                        else:
                            if kj_index != -1:
                                value_string = token[:kj_index].replace(',', '.').replace(' ', '').strip()
                                values_list.append(float(value_string) * 0.2390057 if len(value_string) != 0 else 0.0)
                            else:
                                value = extract_size_in_g(token.strip())
                                values_list.append(value if value else 0.0)
                    if len(keys_list) != len(values_list):
                        print("Keys and values lists do not match for", product_name, "keys:", keys_list, "values:", values_list)
                        continue
                    else:
                        kv_pairs = dict(zip(keys_list, values_list))
                    
                except TimeoutException:
                    try:
                        unit_size_element = driver.find_element(By.XPATH, "//h2[text()='Ravintosisältö']//parent::button//following::h3").text
                        unit_size = extract_portion_size(unit_size_element)
                        keys = driver.find_elements(By.XPATH, "//h2[text()='Ravintosisältö']//parent::button//following::dt[starts-with(@id, 'product-nutritional-detail')]")
                        values = driver.find_elements(By.XPATH, "//h2[text()='Ravintosisältö']//parent::button//following::dd[starts-with(@class, 'NewNutritionalDetails')][1]")

                        for key in keys:
                            keys_list.append(key.text.strip())

                        for value in values:
                            value_string = value.text.strip()
                            value_in_grams = extract_size_in_g(value_string)
                            values_list.append(value_in_grams if value_in_grams else 0.0)

                        if len(keys_list) != len(values_list):
                            print("Keys and values lists do not match for", product_name, "keys:", keys_list, "values:", values_list)
                            continue
                        else:
                            kv_pairs = dict(zip(keys_list, values_list))

                    except Exception as e:
                        print("Could not get nutritional content for", product_name, e)
                        product_price_dict[ean_code].update(nutritional_content_dict[ean_code])
                        if discount == 'Yes':
                            discounted_product_price_dict[ean_code] = product_price_dict[ean_code]
                        url_counter += 1
                        url_progress_bar.progress(int((url_counter / total_urls) * 100))
                        continue
                

                nutritional_content_dict[ean_code]['Nutritional Value per'] = unit_size
                nutritional_content_dict[ean_code].update(kv_pairs)
                product_price_dict[ean_code].update(nutritional_content_dict[ean_code])

            if discount == 'Yes':
                discounted_product_price_dict[ean_code] = product_price_dict[ean_code]

            url_counter += 1
            url_progress_bar.progress(int((url_counter / total_urls) * 100))

        elapsed = time.perf_counter() - start0
        elapsed_time_text.text(f"Total time elapsed: {round(elapsed/60, 2)} minutes")

        store_counter += 1
        end = time.perf_counter()
        url_progress_text.text(f"Finished scraping {url_counter} new product URL(s) in {round((end - url_start) / 60, 2)} minutes ")
        store_progress_text.text(f"Finished scraping store {store_counter} of {total_stores} - {store_title} in {round((end - start) / 60, 2)} minutes")
        store_progress_bar.progress(int((store_counter / total_stores) * 100))

        print(f"Finished scraping store {counter} of {number_of_stores} - {store_title} in {round((end - start) / 60, 2)} minutes")

        try:
            # Serializing json
            nutritional_content_json = json.dumps(nutritional_content_dict, indent=4)
            # Writing to json
            with open("nutritional_content_data.json", "w") as outfile:
                outfile.write(nutritional_content_json)
        except Exception as e:
            print("Could not write nutritional content data to JSON file", e)
        
        try:
            product_price_json = json.dumps(product_price_dict, indent=4)
            with open("product_prices_data.json", "w") as outfile2:
                outfile2.write(product_price_json)
        except Exception as e:
            print("Could not write product price data to JSON file", e)

        driver.get(f"https://www.k-ruoka.fi/?kaupat&kauppahaku={search_string}")

    elapsed = time.perf_counter() - start0
    elapsed_time_text.text(f"Total time elapsed: {round(elapsed/60, 2)} minutes")
    
    end2 = time.perf_counter()
    print("Finished scraping all of", number_of_stores, "in", round((end2 - start0) / 60, 2), "minutes")
    driver.quit()
    
    store_progress_bar.empty()
    category_progress_bar.empty()
    product_progress_bar.empty()
    url_progress_bar.empty()

    updated_discounted_product_price_dict = discounted_product_price_dict.copy()
    return updated_discounted_product_price_dict

def postprocess_and_save(discounted_product_price_dict):
    try:
        discounted_product_price_json = json.dumps(discounted_product_price_dict, indent=4)
        with open("discounted_product_prices_data.json", "w") as outfile:
            outfile.write(discounted_product_price_json)
    except Exception as e:
        print("Could not write discounted products price data to JSON file", e)

    try:
        with open('product_prices_data.json', 'r') as file2:
            product_price_dict = json.load(file2)
    except IOError:
        print("Could not open product price data file.")
        product_price_dict = {}
    
    current_time = datetime.now()
    file_name = f"product_prices_kruoka_{current_time.strftime('%d')}_{current_time.strftime('%b')}_{current_time.strftime('%H')}_{current_time.strftime('%M')}.xlsx"

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

    discounted_product_data_df = pd.DataFrame.from_dict(discounted_product_price_dict, orient='index')
    discounted_product_data_df['Size (kg)'] = discounted_product_data_df['Size (kg)'].apply(lambda x: x if isinstance(x, (float, int)) else 0)

    # Calculate 'Euroa per 100g Proteiinia' with zero-check for 'Proteiinia per 100g' and 'Size (kg)' columns
    discounted_product_data_df['Euroa per 100g Proteiinia'] = np.where(
        (discounted_product_data_df['Proteiinia per 100g'] > 0),
        (discounted_product_data_df['Price per kg'] / 10.00) * (100 / discounted_product_data_df['Proteiinia per 100g']),
        9999.999
    )

    discounted_product_data_df.index.name = 'EAN-code'

    with pd.ExcelWriter(f"discounted_{file_name}") as writer:
        discounted_product_data_df.to_excel(writer, sheet_name='Products')

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

    product_data_df = pd.DataFrame.from_dict(product_price_dict, orient='index')
    product_data_df['Size (kg)'] = product_data_df['Size (kg)'].apply(lambda x: x if isinstance(x, (float, int)) else 0)

    # Calculate 'Euroa per 100g Proteiinia' with zero-check for 'Proteiinia per 100g' and 'Size (kg)' columns
    product_data_df['Euroa per 100g Proteiinia'] = np.where(
        (product_data_df['Proteiinia per 100g'] > 0),
        (product_data_df['Price per kg'] / 10.00) * (100 / product_data_df['Proteiinia per 100g']),
        9999.999
    )

    product_data_df.index.name = 'EAN-code'

    with pd.ExcelWriter(f"{file_name}") as writer:
        product_data_df.to_excel(writer, sheet_name='Products')

def get_logpath() -> str:
    return os.path.join(os.getcwd(), 'selenium.log')

def get_chromedriver_path() -> str:
    return shutil.which('chromedriver')

def get_webdriver_service(logpath) -> Service:
    service = Service(
    executable_path=get_chromedriver_path(),
    log_output=logpath,
    )
    return service

st.set_page_config(page_title="K-Ruoka Web Scraper", page_icon="📈")

st.markdown("# Product scraper")
st.write(
    """This is a web scraper for K-Ruoka that collects product information, nutritional content, and prices from K-Ruoka's website. The scraped data is merged into existing data and can be processed further."""
)

st.header("1. Enter search phrase for store locations")
store_locations_input = st.text_input(
    "Write the search phrase for locations (e.g. Tampere)",
    value=""
)

if "store_options" not in st.session_state:
    st.session_state.store_options = []
if "driver" not in st.session_state:
    st.session_state.driver = None
if "send_clicked" not in st.session_state:
    st.session_state.send_clicked = False

if st.button("Confirm"):
    st.session_state.store_options, st.session_state.driver = get_stores_list(store_locations_input)
    st.session_state.send_clicked = True

if st.session_state.send_clicked:
    st.header("2. Select stores to be scraped")
    selected_stores = st.multiselect(
        "Select stores ",
        options=st.session_state.store_options,
        default=[]
    )

    st.header("3. Choose the product categories to scrape")
    selected_categories = st.multiselect(
        "Choose product categories",
        options=category_options.keys(),
        default=[]
    )

    st.header("4. Select if you want to scrape only discounted products")
    st.info('Scraping all of the products instead of only the discounted takes significantly longer', icon="ℹ️")
    only_discounted_products_button = st.toggle("Scrape only discounted products", value=True)
    if only_discounted_products_button:
        only_discounted_products = True
    else:
        only_discounted_products = False

    if st.button("Start scraping"):
        if not selected_categories:
            st.warning("Choose at least one product category.")
        elif not selected_stores:
            st.warning("Select at least one store.")
        elif not store_locations_input:
            st.warning("Enter a search phrase for store locations.")
        else:
            nutritional_content_dict, product_price_dict, discounted_product_price_dict = preprocess_and_save()
            selected_categories = [category_options[cat] for cat in selected_categories]

            st.success(f"Chosen product categories: {', '.join([inverted_category_options[cat] for cat in selected_categories])}")
            st.success(f"Chosen stores: {', '.join(selected_stores)}")

            updated_discounted_product_price_dict = run_scraper(selected_categories, selected_stores, nutritional_content_dict, product_price_dict, discounted_product_price_dict, store_locations_input, st.session_state.driver, only_discounted_products)
            _ = postprocess_and_save(updated_discounted_product_price_dict)



