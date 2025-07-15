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
from datetime import date
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

def get_caffeine_amount(driver):
    percent_pattern = re.compile(r'kofeiin(ipitoisuus)?[ia]*\s*\(\s*(\d+,\d+|\d+)\s*%\s*\)', re.IGNORECASE)
    percent_pattern_2 = re.compile(r'(\d+,\d+|\d+)\s*%\s*\)?[^a-zA-Z0-9]{0,10}kofeiin(ipitoisuus)?[ia]*', re.IGNORECASE)
    percent_pattern_3 = re.compile(r'\(([^)]*kofeiin[ia]*[^)]*?)(\d+,\d+|\d+)\s*%\)', re.IGNORECASE)
    mg_100ml_pattern = re.compile(r'kofeiin(ipitoisuus)?[ia]*\s*\(\s*(\d+)\s*mg\s*/\s*100\s*ml\s*\)', re.IGNORECASE)
    mg_100ml_pattern_2 = re.compile(r'(\d+)\s*mg\s*/\s*100\s*ml[^a-zA-Z0-9]{0,10}kofeiin(ipitoisuus)?[ia]*', re.IGNORECASE)
    mg_100ml_pattern_3 = re.compile(r'\(([^)]*kofeiin[ia]*[^)]*?)(\d+)\s*mg\s*/\s*100\s*ml\)', re.IGNORECASE)
    mg_l_pattern = re.compile(r'kofeiin(ipitoisuus)?[ia]*\s*\(\s*(\d+)\s*mg\s*/\s*l\s*\)', re.IGNORECASE)
    mg_l_pattern_2 = re.compile(r'(\d+)\s*mg\s*/\s*l[^a-zA-Z0-9]{0,10}kofeiin(ipitoisuus)?[ia]*', re.IGNORECASE)
    mg_l_pattern_3 = re.compile(r'\(([^)]*kofeiin[ia]*[^)]*?)(\d+)\s*mg\s*/\s*l\)', re.IGNORECASE)
    amount_pattern = re.compile(r'(\d+(?:[.,]\d+)?)\s*mg\s*kofeiin[ia]*', re.IGNORECASE)

    caffeine_amount = 0
    caffeine_content = 0
    
    try:
        wait = WebDriverWait(driver, 3)
        product_description = wait.until(EC.presence_of_element_located((By.XPATH, "//p[starts-with(@class, 'ProductDetailsstyle__Description')]"))).text
        if mg_100ml_match := mg_100ml_pattern.search(product_description):
            caffeine_content = int(mg_100ml_match.group(2)) 
        elif mg_100ml_match_2 := mg_100ml_pattern_2.search(product_description):
            caffeine_content = int(mg_100ml_match_2.group(1)) 
        elif mg_100ml_match_3 := mg_100ml_pattern_3.search(product_description):
            caffeine_content = int(mg_100ml_match_3.group(2)) 
        elif mg_l_match := mg_l_pattern.search(product_description):
            caffeine_content = int(mg_l_match.group(2)) / 10
        elif mg_l_match_2 := mg_l_pattern_2.search(product_description):
            caffeine_content = int(mg_l_match_2.group(1)) / 10
        elif mg_l_match_3 := mg_l_pattern_3.search(product_description):
            caffeine_content = int(mg_l_match_3.group(2)) / 10
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
        
        if caffeine_content != 0:
            return caffeine_content, caffeine_amount
    except TimeoutException:
        product_description = None

    try:
        wait = WebDriverWait(driver, 3)
        product_info_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//h2[text()='Tuotetiedot']")))
        product_info_header.click()
        time.sleep(1)
        try:
            wait = WebDriverWait(driver, 3)
            product_details = wait.until(EC.presence_of_element_located((By.XPATH, "//h3[text()='Ainesosat']//following-sibling::p"))).text

            if mg_100ml_match := mg_100ml_pattern.search(product_details):
                caffeine_content = int(mg_100ml_match.group(2)) 
            elif mg_100ml_match_2 := mg_100ml_pattern_2.search(product_details):
                caffeine_content = int(mg_100ml_match_2.group(1)) 
            elif mg_100ml_match_3 := mg_100ml_pattern_3.search(product_details):
                caffeine_content = int(mg_100ml_match_3.group(2)) 
            elif mg_l_match := mg_l_pattern.search(product_details):
                caffeine_content = int(mg_l_match.group(2)) / 10
            elif mg_l_match_2 := mg_l_pattern_2.search(product_details):
                caffeine_content = int(mg_l_match_2.group(1)) / 10
            elif mg_l_match_3 := mg_l_pattern_3.search(product_details):
                caffeine_content = int(mg_l_match_3.group(2)) / 10
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
            
            if caffeine_content != 0:
                product_info_header.click()
                time.sleep(1)
                return caffeine_content, caffeine_amount
        except TimeoutException:
            product_details = ''
        try:
            wait = WebDriverWait(driver, 3)
            product_instructions = wait.until(EC.presence_of_element_located((By.XPATH, "//h3[text()='Säilytys- ja käyttöohjeet']//following-sibling::div"))).text

            if mg_100ml_match := mg_100ml_pattern.search(product_instructions):
                caffeine_content = int(mg_100ml_match.group(2)) 
            elif mg_100ml_match_2 := mg_100ml_pattern_2.search(product_instructions):
                caffeine_content = int(mg_100ml_match_2.group(1)) 
            elif mg_100ml_match_3 := mg_100ml_pattern_3.search(product_instructions):
                caffeine_content = int(mg_100ml_match_3.group(2)) 
            elif mg_l_match := mg_l_pattern.search(product_instructions):
                caffeine_content = int(mg_l_match.group(2)) / 10
            elif mg_l_match_2 := mg_l_pattern_2.search(product_instructions):
                caffeine_content = int(mg_l_match_2.group(1)) / 10
            elif mg_l_match_3 := mg_l_pattern_3.search(product_instructions):
                caffeine_content = int(mg_l_match_3.group(2)) / 10
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
            
            if caffeine_content != 0:
                product_info_header.click()
                time.sleep(1)
                return caffeine_content, caffeine_amount
        except TimeoutException:
            product_instructions = ''
        product_info_header.click()
        time.sleep(1)
                                                                        
    except TimeoutException:
        product_details = None
        product_instructions = None

    return caffeine_content, caffeine_amount

product_categories = ['pakasteet/liha--kala--ja-kasvispakasteet',
                      'pakasteet/pakasteateriat',
                      'pakasteet/pizzat-ja-pizzapohjat',
                      'liha-ja-kasviproteiinit',
                      'kala-ja-merenelavat',
                      'valmisruoka/valmisruoat-ja--keitot',
                      'valmisruoka/laatikot-pastat-ja-lasagnet',
                      'valmisruoka/pyorykat-pihvit-ja-ohukaiset',
                      'maito-juusto-munat-ja-rasvat/maitotuotteet/rahkat',
                      'maito-juusto-munat-ja-rasvat/jogurtit',
                      'maito-juusto-munat-ja-rasvat/munat',
                      'maito-juusto-munat-ja-rasvat/rahkat-vanukkaat-ja-valipalat',
                      'maito-juusto-munat-ja-rasvat/juustot-leivan-paalle',
                      'maito-juusto-munat-ja-rasvat/ruoka--ja-herkuttelujuustot',
                      'kuivat-elintarvikkeet-ja-leivonta/leseet-rouheet-alkiot-soijavalmisteet-ja-muut-viljatuotteet/leseet-rouheet-alkiot-soijavalmisteet-ja-viljajyvat',
                      'kuivat-elintarvikkeet-ja-leivonta/siemenet-pahkinat-ja-kuivatut-hedelmat/pahkinat',
                      'kuivat-elintarvikkeet-ja-leivonta/siemenet-pahkinat-ja-kuivatut-hedelmat/siemenet',
                      'kuivat-elintarvikkeet-ja-leivonta/kuivatut-herneet-pavut-ja-linssit',
                      'sailykkeet-keitot-ja-ateria-ainekset/sailykkeet/vihannessailykkeet',
                      'sailykkeet-keitot-ja-ateria-ainekset/sailykkeet/valmisruokasailykkeet',
                      'sailykkeet-keitot-ja-ateria-ainekset/sailykkeet/kala--ja-ayriaissailykkeet',
                      'sailykkeet-keitot-ja-ateria-ainekset/sailykkeet/liha--ja-riistasailykkeet',
                      'kosmetiikka-terveys-ja-hygienia/terveysvalmisteet/urheiluvalmisteet',
                      'juomat/energia--ja-urheilujuomat/energiajuomat'
                      ]

store_locations = ['Tampere', 'Pirkkala', 'Lempäälä', 'Nokia']

try:
    with open('nutritional_content_data.json', 'r') as file:
        nutritional_content_dict = json.load(file)
except IOError:
    print("Could not open nutritional content data file. Using empty dictionary.")
    nutritional_content_dict = {}

try:
    with open('product_price_data.json', 'r') as file2:
        product_price_dict = json.load(file2)
except IOError:
    print("Could not open product price data file. Using empty dictionary.")
    product_price_dict = {}

driver = uc.Chrome()
driver.get("https://www.k-ruoka.fi/?kaupat&kauppahaku=Tampere&ketju=kcitymarket&ketju=ksupermarket")

wait = WebDriverWait(driver, 10)
try:
    accept_cookies = wait.until(EC.element_to_be_clickable((By.XPATH, f"//button[@id='onetrust-accept-btn-handler']")))
    # Click accept cookies button
    accept_cookies.click()
except TimeoutException:
    print('Prompt to accept cookies did not pop up')

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

counter = 0

while counter < number_of_stores:
    wait = WebDriverWait(driver, 10)
    try:
        store_list_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-component='store-list']")))
        stores = store_list_element.find_elements(By.XPATH, ".//li[@data-component='store-list-item']")
        new_number_of_stores = len(stores)

        while True:
            if new_number_of_stores == number_of_stores:
                break
            store_list_container = driver.find_element(By.XPATH, "//div[starts-with(@class, 'StoreSelector__StyledVerticalScrollAwareContainer')]")
            store_list_container.click()
            store_list_container.send_keys(Keys.PAGE_DOWN)

            time.sleep(1)

            stores = store_list_element.find_elements(By.XPATH, ".//li[@data-component='store-list-item']")
            new_number_of_stores = len(stores)

        store_name = stores[counter].get_attribute("data-store")
        store_location = stores[counter].find_element(By.XPATH, ".//div[@data-testid='store-location']").text

        if store_location not in store_locations:
            print(f"Skipping store {counter + 1} - {store_name} as {store_location} is not in the specified locations")
            counter += 1
            continue

        store = stores[counter].find_element(By.XPATH, f".//button[@data-select-store='{store_name}']")
        driver.execute_script("arguments[0].scrollIntoView()", store)

        time.sleep(1)
        store.click()
        time.sleep(2)

        counter += 1

    except TimeoutException:
        print("Store list element not found or not visible")
        counter += 1
        continue

    product_urls = {}
    for category in product_categories:
        driver.get(f"https://www.k-ruoka.fi/kauppa/tuotehaku/{category}")
        time.sleep(1)
        
        wait = WebDriverWait(driver, 3)
        try:
            products_list = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-testid='product-search-results']")))
            product_cards = products_list.find_elements(By.XPATH, ".//li[@data-testid='product-card']")
            last_list_len = len(product_cards)
            while True:
                driver.execute_script("arguments[0].scrollIntoView()", product_cards[-1])
                time.sleep(3)
                product_cards = driver.find_elements(By.XPATH, "//ul[@data-testid='product-search-results']//li[@data-testid='product-card']")
                new_list_len = len(product_cards)
                if new_list_len == last_list_len:
                    break
                last_list_len = new_list_len
        except TimeoutException:
            print("No products found in", store_name, "for category", category)
            continue
        
        wait = WebDriverWait(driver, 3)
        try:
            products_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-testid='product-search-results']")))
            product_cards = products_element.find_elements(By.XPATH, ".//li[@data-testid='product-card']")
        except TimeoutException:
            print("No products found in", store_name, "for category", category, "after scrolling")
            continue

        for card in product_cards:
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
                        continue
                    ean_code = ean_code_string[:hyphen_index] if hyphen_index != -1 else ean_code_string

                    discount_badge_elements = card.find_elements(By.XPATH, ".//div[starts-with(@class, 'ProductCard__Discount')]")
                    normal_price_elements = card.find_elements(By.XPATH, ".//div[@data-testid='product-normal-price']")
                    if len(discount_badge_elements) > 0 or len(normal_price_elements) > 0:
                        discount = 'Yes'
                    else:
                        discount = 'No'
                    
                except Exception as e:
                    print("Could not get product url or EAN code for", card.text, e)
                    continue

                if nutritional_content_dict.get(ean_code, 'Unknown') != 'Unknown':
                    if nutritional_content_dict[ean_code].get('Vegan', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Lactose Free', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Gluten Free', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Sydänmerkki', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Hyvää Suomesta', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Organic', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Category', 'Unknown') == 'Unknown':
                        product_urls.update({url: ean_code})
                        if product_price_dict.get(ean_code, "None") != "None":
                            product_price_dict[ean_code]['Size (kg)'] = size
                            product_price_dict[ean_code]['Store'] = store_name
                            #product_price_dict[ean_code]['Category'] = category

                        else:
                            product_price_dict[ean_code] = {}
                            product_price_dict[ean_code]['Size (kg)'] = size
                            product_price_dict[ean_code]['Store'] = store_name
                            #product_price_dict[ean_code]['Category'] = category
                    else:
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

                        if product_price_dict.get(ean_code, "None") != "None":
                            if product_price_dict[ean_code].get('Price per Unit', 'Unknown') != 'Unknown':
                                if product_price_dict[ean_code]['Price per Unit'] > unit_price:
                                    product_price_dict[ean_code]['Price per Unit'] = unit_price
                                    product_price_dict[ean_code]['Unit'] = unit_type
                                    product_price_dict[ean_code]['Size (kg)'] = size
                                    product_price_dict[ean_code]['Store'] = store_name
                                    #product_price_dict[ean_code]['Category'] = category
                                    if discount == 'Yes':
                                        if product_price_dict[ean_code].get('Discount valid until', 'Unknown') == 'Unknown':
                                            product_urls.update({url: ean_code})
                                        else:
                                            if date.today() > product_price_dict[ean_code]['Discount valid until']:
                                                product_urls.update({url: ean_code})

                            else:
                                product_price_dict[ean_code]['Price per Unit'] = unit_price
                                product_price_dict[ean_code]['Unit'] = unit_type
                                product_price_dict[ean_code]['Size (kg)'] = size
                                product_price_dict[ean_code]['Store'] = store_name
                                #product_price_dict[ean_code]['Category'] = category
                                if discount == 'Yes':
                                    if product_price_dict[ean_code].get('Discount valid until', 'Unknown') == 'Unknown':
                                        product_urls.update({url: ean_code})
                                    else:
                                        if date.today() > product_price_dict[ean_code]['Discount valid until']:
                                            product_urls.update({url: ean_code})

                        else:
                            product_price_dict[ean_code] = {}
                            product_price_dict[ean_code]['Price per Unit'] = unit_price
                            product_price_dict[ean_code]['Unit'] = unit_type
                            product_price_dict[ean_code]['Size (kg)'] = size
                            product_price_dict[ean_code]['Store'] = store_name
                            #product_price_dict[ean_code]['Category'] = category
                            if discount == 'Yes':
                                if product_price_dict[ean_code].get('Discount valid until', 'Unknown') == 'Unknown':
                                    product_urls.update({url: ean_code})
                                else:
                                    if date.today() > product_price_dict[ean_code]['Discount valid until']:
                                        product_urls.update({url: ean_code})

                    product_price_dict[ean_code].update(nutritional_content_dict[ean_code])

                else:
                    product_urls.update({url: ean_code})
                    product_price_dict[ean_code] = {}
                    product_price_dict[ean_code]['Name'] = product_name
                    product_price_dict[ean_code]['Size (kg)'] = size
                    product_price_dict[ean_code]['Store'] = store_name
                    product_price_dict[ean_code]['Category'] = category

            else:
                try:
                    driver.execute_script("arguments[0].scrollIntoView()", card)
                    nayta_tuotteet_button = card.find_element(By.XPATH, ".//button[@aria-label='Näytä tuotteet']")
                    nayta_tuotteet_button.click()
                except Exception as e:
                    try:
                        time.sleep(1)
                        nayta_tuotteet_button = card.find_element(By.XPATH, ".//button[@aria-label='Näytä tuotteet']")
                        driver.execute_script("arguments[0].scrollIntoView()", nayta_tuotteet_button)
                        time.sleep(1)
                        nayta_tuotteet_button.click()
                    except Exception as e:
                        try:
                            driver.find_element(By.XPATH, "//button[@title='Sulje']").click()
                            time.sleep(3)

                            nayta_tuotteet_button = card.find_element(By.XPATH, ".//button[@aria-label='Näytä tuotteet']")
                            driver.execute_script("arguments[0].scrollIntoView()", nayta_tuotteet_button)
                            time.sleep(3)
                            nayta_tuotteet_button.click()
                            
                        except Exception as e:
                            print("Could not click 'Näytä tuotteet' button for", card.text, e)
                            continue

                wait = WebDriverWait(driver, 5)
                try:
                    products_list_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-testid='offer-products']")))
                    product_elements = driver.find_elements(By.XPATH, "//ul[@data-testid='offer-products']//li[@data-testid='product-card']")
                except TimeoutException:
                    print("Product elements not found or not visible for", card.text)
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

                    except Exception as e:
                        print("Could not get product url or EAN code for", product.text, e)
                        continue

                    if nutritional_content_dict.get(ean_code, 'Unknown') != 'Unknown':
                        if nutritional_content_dict[ean_code].get('Vegan', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Lactose Free', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Gluten Free', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Sydänmerkki', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Hyvää Suomesta', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Organic', 'Unknown') == 'Unknown' or nutritional_content_dict[ean_code].get('Category', 'Unknown') == 'Unknown':
                            product_urls.update({url: ean_code})
                            if product_price_dict.get(ean_code, "None") != "None":
                                product_price_dict[ean_code]['Size (kg)'] = size
                                product_price_dict[ean_code]['Store'] = store_name
                                #product_price_dict[ean_code]['Category'] = category
                            else:
                                product_price_dict[ean_code] = {}
                                product_price_dict[ean_code]['Size (kg)'] = size
                                product_price_dict[ean_code]['Store'] = store_name
                                #product_price_dict[ean_code]['Category'] = category
                        else:
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

                            if product_price_dict.get(ean_code, "None") != "None":
                                if product_price_dict[ean_code].get('Price per Unit', 'Unknown') != 'Unknown':
                                    if product_price_dict[ean_code]['Price per Unit'] > unit_price:
                                        product_price_dict[ean_code]['Price per Unit'] = unit_price
                                        product_price_dict[ean_code]['Unit'] = unit_type
                                        product_price_dict[ean_code]['Size (kg)'] = size
                                        product_price_dict[ean_code]['Store'] = store_name
                                        #product_price_dict[ean_code]['Category'] = category

                                else:
                                    product_price_dict[ean_code]['Price per Unit'] = unit_price
                                    product_price_dict[ean_code]['Unit'] = unit_type
                                    product_price_dict[ean_code]['Size (kg)'] = size
                                    product_price_dict[ean_code]['Store'] = store_name
                                    #product_price_dict[ean_code]['Category'] = category

                            else:
                                product_price_dict[ean_code] = {}
                                product_price_dict[ean_code]['Price per Unit'] = unit_price
                                product_price_dict[ean_code]['Unit'] = unit_type
                                product_price_dict[ean_code]['Size (kg)'] = size
                                product_price_dict[ean_code]['Store'] = store_name
                                #product_price_dict[ean_code]['Category'] = category

                        product_price_dict[ean_code].update(nutritional_content_dict[ean_code])

                    else:
                        product_urls.update({url: ean_code})
                        product_price_dict[ean_code] = {}
                        product_price_dict[ean_code]['Name'] = product_name
                        product_price_dict[ean_code]['Size (kg)'] = size
                        product_price_dict[ean_code]['Store'] = store_name
                        #product_price_dict[ean_code]['Category'] = category
                        
                driver.find_element(By.XPATH, "//button[@title='Sulje']").click()
                time.sleep(2)

    for url, ean in product_urls.items():
        if url is not None:
            driver.get(url)
        else:
            print("Could not open link")
            continue

        try:
            wait = WebDriverWait(driver, 5)
            header = wait.until(EC.visibility_of_element_located((By.XPATH, "//h1[@data-testid='product-name']")))
            product_name = header.text
            size = extract_size_in_kg(product_name)
            category_elements = driver.find_elements(By.XPATH, "//li[starts-with(@class, 'Breadcrumbs__BreadcrumbsItem')]")
            if len(category_elements) > 0:
                category = category_elements[-1].text
            else:
                category = 'Unknown'
        except TimeoutException:
            print("Product name not found for", url)
            continue

        if ean == 'Unknown' or ean == '':
            try:
                wait = WebDriverWait(driver, 5)
                product_info_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//h2[text()='Tuotetiedot']")))
                product_info_header.click()
                time.sleep(1)
            except TimeoutException:
                print("Product info header not found for", product_name)
                continue

            try:
                wait = WebDriverWait(driver, 5)
                ean_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//h3[text()='EAN-koodi']//following-sibling::p")))
                ean_code = ean_element.text
                product_info_header.click()
                time.sleep(1)
            except TimeoutException:
                print("EAN code not found for", product_name)
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

        normal_price_elements = driver.find_elements(By.XPATH, "//h1[@data-testid='product-name']//following-sibling::div//div[@data-testid='product-normal-price']")
        if len(normal_price_elements) > 0:
            discount = 'Yes'
            valid_during_elements = driver.find_elements(By.XPATH, "//h1[@data-testid='product-name']//following-sibling::div[starts-with(@class, 'ProductSidebarContent__Info')]")
            if len(valid_during_elements) > 0:
                valid_during_string = valid_during_elements[0].text
                
                search_res = re.findall(r'(\d+(?:[.,]\d+)?)', valid_during_string)
                if len(search_res) > 1:
                    valid_starting_string = search_res[0]
                    valid_until_string = search_res[1]

                    today = date.today()
                    valid_starting = datetime(today.year, int(valid_starting_string.split('.')[1]), int(valid_starting_string.split('.')[0]))
                    valid_until = datetime(today.year, int(valid_until_string.split('.')[1]), int(valid_until_string.split('.')[0]))
                else:
                    valid_starting = 'Unknown'
                    valid_until = 'Unknown'
            else:
                valid_starting = 'Unknown'
                valid_until = 'Unknown'
        else:
            discount = 'No'
            valid_starting = None
            valid_until = None

        if product_price_dict.get(ean_code, 'Unknown') != 'Unknown':
            if product_price_dict[ean_code].get('Price per Unit', 'Unknown') != 'Unknown':
                if product_price_dict[ean_code]['Price per Unit'] > unit_price:
                    product_price_dict[ean_code]['Price per Unit'] = unit_price
                    product_price_dict[ean_code]['Unit'] = unit_type
                    product_price_dict[ean_code]['Size (kg)'] = size
                    product_price_dict[ean_code]['Store'] = store_name

            else:
                product_price_dict[ean_code]['Price per Unit'] = unit_price
                product_price_dict[ean_code]['Unit'] = unit_type
                product_price_dict[ean_code]['Size (kg)'] = size
                product_price_dict[ean_code]['Store'] = store_name

        else:
            product_price_dict[ean_code] = {}
            product_price_dict[ean_code]['Name'] = product_name
            product_price_dict[ean_code]['Price per Unit'] = unit_price
            product_price_dict[ean_code]['Unit'] = unit_type
            product_price_dict[ean_code]['Size (kg)'] = size
            product_price_dict[ean_code]['Store'] = store_name

        if discount == 'Yes':
            product_price_dict[ean_code]['Discount valid starting'] = valid_starting
            product_price_dict[ean_code]['Discount valid until'] = valid_until

        if nutritional_content_dict.get(ean_code, 'Unknown') != 'Unknown':
            nutritional_content_dict[ean_code]['Category'] = category
            if category == 'Energiajuomat' or category == 'Urheiluvalmisteet':
                if nutritional_content_dict[ean_code].get('Kofeiini (per 100ml)', 'Unknown') == 'Unknown':
                    caffeine_content, caffeine_amount = get_caffeine_amount(driver)
                    nutritional_content_dict[ean_code]['Kofeiini (per 100ml)'] = caffeine_content
                    nutritional_content_dict[ean_code]['Kofeiini (per tuote)'] = caffeine_amount
            nutritional_content_dict[ean_code]['Vegan'] = vegan if nutritional_content_dict[ean_code].get('Vegan', 'Unknown') == 'Unknown' or vegan == 'Yes' else nutritional_content_dict[ean_code]['Vegan']
            nutritional_content_dict[ean_code]['Gluten Free'] = gluten_free if nutritional_content_dict[ean_code].get('Gluten Free', 'Unknown') == 'Unknown' or gluten_free == 'Yes' else nutritional_content_dict[ean_code]['Gluten Free']
            nutritional_content_dict[ean_code]['Lactose Free'] = lactose_free if nutritional_content_dict[ean_code].get('Lactose Free', 'Unknown') == 'Unknown' or lactose_free == 'Yes' else nutritional_content_dict[ean_code]['Lactose Free']
            nutritional_content_dict[ean_code]['Organic'] = luomu if nutritional_content_dict[ean_code].get('Organic', 'Unknown') == 'Unknown' or luomu == 'Yes' else nutritional_content_dict[ean_code]['Organic']
            nutritional_content_dict[ean_code]['Sydänmerkki'] = sydanmerkki if nutritional_content_dict[ean_code].get('Sydänmerkki', 'Unknown') == 'Unknown' or sydanmerkki == 'Yes' else nutritional_content_dict[ean_code]['Sydänmerkki']
            nutritional_content_dict[ean_code]['Hyvää Suomesta'] = hyvaa_suomesta if nutritional_content_dict[ean_code].get('Hyvää Suomesta', 'Unknown') == 'Unknown' or hyvaa_suomesta == 'Yes' else nutritional_content_dict[ean_code]['Hyvää Suomesta']
            product_price_dict[ean_code].update(nutritional_content_dict[ean_code])
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
            if category == 'Energiajuomat' or category == 'Urheiluvalmisteet':
                caffeine_content, caffeine_amount = get_caffeine_amount(driver)
                nutritional_content_dict[ean_code]['Kofeiini (per 100ml)'] = caffeine_content
                nutritional_content_dict[ean_code]['Kofeiini (per tuote)'] = caffeine_amount
            try:
                wait = WebDriverWait(driver, 2)
                nutritional_content_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//h2[text()='Ravintosisältö']")))
                nutritional_content_header.click()
                time.sleep(1)
            except ElementClickInterceptedException:
                nutritional_content_header = driver.find_element(By.XPATH, "//h2[text()='Ravintosisältö']")
                driver.execute_script("arguments[0].scrollIntoView()", nutritional_content_header)
                time.sleep(1)
                nutritional_content_header.click()
                time.sleep(1)
            except TimeoutException:
                print("Nutritional content header not found for", product_name)
                product_price_dict[ean_code].update(nutritional_content_dict[ean_code])
                continue

            keys_list = []
            values_list = []
            try:
                wait = WebDriverWait(driver, 3)
                table = wait.until(EC.visibility_of_element_located((By.XPATH, "//h2[(text()='Ravintosisältö')]//parent::button//following::table[starts-with(@class, 'NewNutritionalDetails__Table')]")))
                unit_size = table.find_element(By.XPATH, ".//th[starts-with(@class, 'NewNutritionalDetails__NutritionContentTableColumnHeading')]//div").text
                product_price_dict[ean_code]['Nutritional Value per'] = unit_size
                keys = table.find_elements(By.XPATH, "//th[starts-with(@class, 'NewNutritionalDetails__NutritionContentTableRowHeading')]")
                values = table.find_elements(By.XPATH, "//td[starts-with(@class, 'NewNutritionalDetails__NutritionContentTableCell')][1]")

                for key in keys:
                    tokens = key.text.split('\n')
                    for token in tokens:
                        keys_list.append(token)

                for value in values:
                    tokens = value.text.split('\n')
                    for token in tokens:
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
                                values_list.append(value) if value != None else 0.0

                kv_pairs = dict(zip(keys_list, values_list))
                
            except TimeoutException:
                try:
                    unit_size_element = driver.find_element(By.XPATH, "//h2[text()='Ravintosisältö']//parent::button//following::h3").text
                    unit_size = extract_portion_size(unit_size_element)
                    keys = driver.find_elements(By.XPATH, "//h2[text()='Ravintosisältö']//parent::button//following::dt[starts-with(@id, 'product-nutritional-detail')]")
                    values = driver.find_elements(By.XPATH, "//h2[text()='Ravintosisältö']//parent::button//following::dd[starts-with(@class, 'NewNutritionalDetails')]")

                    for key in keys:
                        keys_list.append(key.text.strip())

                    for value in values:
                        value_string = value.text.strip()
                        value_in_grams = extract_size_in_g(value_string)
                        values_list.append(value_in_grams) if value_in_grams is not None else 0.0
                    kv_pairs = dict(zip(keys_list, values_list))

                except Exception as e:
                    print("Could not get nutritional content for", product_name, e)
                    product_price_dict[ean_code].update(nutritional_content_dict[ean_code])
                    continue

            nutritional_content_dict[ean_code]['Nutritional Value per'] = unit_size
            nutritional_content_dict[ean_code].update(kv_pairs)
            product_price_dict[ean_code].update(nutritional_content_dict[ean_code])

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
        with open("product_price.json", "w") as outfile:
            outfile.write(product_price_json)
    except Exception as e:
        print("Could not write product price data to JSON file", e)

    print(f"Finished scraping store {counter} of {number_of_stores} - {store_name}")

    driver.get("https://www.k-ruoka.fi/?kaupat&kauppahaku=Tampere&ketju=kcitymarket&ketju=ksupermarket")

current_time = datetime.now()
file_name = f"{current_time.strftime('%d')}_{current_time.strftime('%b')}_product_unit_prices_kruoka.xlsx"

for products, product_data in product_price_dict.items():
    portion_size_string = product_data.get('Nutritional Value per', 'Unknown')
    portion_size_in_grams = extract_size_in_g(portion_size_string)

    if portion_size_in_grams != 0 and portion_size_in_grams is not None:
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
    (product_data_df['Unit'].isin(['kg', 'l'])) & (product_data_df['Proteiinia per 100g'] > 0),
    (product_data_df['Price per Unit'] / 10.00) * (100 / product_data_df['Proteiinia per 100g']),
    np.where(
        (product_data_df['Proteiinia per 100g'] > 0) & (product_data_df['Size (kg)'] > 0),
        ((product_data_df['Price per Unit'] / product_data_df['Size (kg)']) / 10) * (100 / product_data_df['Proteiinia per 100g']),
        np.nan
    )
)
product_data_df.index.name = 'EAN-code'

with pd.ExcelWriter(f"{file_name}") as writer:
    product_data_df.to_excel(writer, sheet_name='Products')

sorted_df = product_data_df.sort_values(by=['Euroa per 100g Proteiinia'], ascending=[True])
display(sorted_df.head(10))

    