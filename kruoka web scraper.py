from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import undetected_chromedriver as uc
import time
import pandas as pd
from IPython.display import display
from datetime import datetime
import json
import os

try:
    with open('nutritional_content_data.json', 'r') as file:
        nutritional_content_dict = json.load(file)
except IOError:
    nutritional_content_dict = {}

driver = uc.Chrome()
driver.get("https://www.k-ruoka.fi/?kaupat&kauppahaku=Tampere")
wait = WebDriverWait(driver, 10)
try:
    accept_cookies = wait.until(EC.element_to_be_clickable((By.XPATH, f"//button[@id='onetrust-accept-btn-handler']")))
    # Click accept cookies button
    accept_cookies.click()
except TimeoutException:
    print('Prompt to accept cookies did not pop up')
wait = WebDriverWait(driver, 10)
try:
    store_list_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-component='store-list']")))
    number_of_stores = len(driver.find_elements(By.XPATH, "//li[@data-component='store-list-item']"))
except TimeoutException:
    print("Store list element not found or not visible")
    driver.quit()
    exit()
counter = 0
product_dict = {}
while counter < 5:
    wait = WebDriverWait(driver, 10)
    try:
        store_list_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-component='store-list']")))
        stores = driver.find_elements(By.XPATH, "//li[@data-component='store-list-item']")
        store_name = stores[counter].get_attribute("data-store")
        store = stores[counter].find_element(By.XPATH, f"//button[@data-select-store='{store_name}']")
        store.click()
        time.sleep(2)
        counter += 1
        driver.get("https://www.k-ruoka.fi/kauppa/tuotehaku/liha-ja-kasviproteiinit")
    except TimeoutException:
        print("Store list element not found or not visible")
        continue
    wait = WebDriverWait(driver, 10)
    try:
        products_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-testid='product-search-results']")))
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollBy(0, 6000)")
            time.sleep(3)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
    except TimeoutException:
        print("Could not scroll to bottom of the page")
    
    driver.execute_script("window.scrollTo(0, 0)")
    wait = WebDriverWait(driver, 10)
    try:
        products_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-testid='product-search-results']")))
        product_cards = driver.find_elements(By.XPATH, "//li[@data-testid='product-card']")
        driver.execute_script("document.body.style.zoom='50%'")
    except TimeoutException:
        print("Product cards not found or not visible")
        continue
    product_urls = []
    
    for card in product_cards:
        url_elements = card.find_elements(By.XPATH, ".//a[@data-testid='product-link']")
        if len(url_elements) > 0:
            try:
                url = url_elements[0].get_attribute("href")
                ean_code = card.get_attribute("data-product-id")
            except Exception as e:
                print("Could not get product url or EAN code for", card.text, e)
                continue
            try:
                unit_price = card.find_element(By.XPATH, ".//div[@data-testid='product-unit-price']").text
                backslash_index = unit_price.find("/")
                if backslash_index != -1:
                    unit_type = unit_price[backslash_index+1:]
                    unit_price = unit_price[:backslash_index]
                    unit_price = float(unit_price.replace(',', '.'))
                else:
                    unit_price = float(unit_price.replace(',', '.'))
                    unit_type = 'Unknown'
            except NoSuchElementException:
                try:
                    price_element_integer = card.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__IntegerPart')]")
                    unit_price_integer = price_element_integer.text
            
                    price_element_decimal = card.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__DecimalPart')]")
                    unit_price_decimal = price_element_decimal.text

                    unit_element = card.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__Extra')]")
                    unit_type = unit_element.text.replace('/', '')
                    
                    unit_price = float(unit_price_integer + '.' + unit_price_decimal)
                except Exception as e:
                    print("Could not get product price for", ean_code, e)
                    product_urls.append(url)
                    continue
            if product_dict.get(ean_code, "None") != "None":
                if product_dict[ean_code]['Price per Unit'] > unit_price:
                    product_dict[ean_code]['Price per Unit'] = unit_price
                    product_dict[ean_code]['Store'] = store_name
            else:
                if nutritional_content_dict.get(ean_code, 'Unknown') != 'Unknown':
                    product_dict[ean_code] = {}
                    product_dict[ean_code]['Price per Unit'] = unit_price
                    product_dict[ean_code]['Unit'] = unit_type
                    product_dict[ean_code]['Store'] = store_name
                    product_dict[ean_code].update(nutritional_content_dict[ean_code])
                else:
                    product_urls.append(url)
        else:
            driver.execute_script("arguments[0].scrollIntoView()", card)
            wait = WebDriverWait(driver, 5)
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, ".//button[@aria-label='Näytä tuotteet']"))).click()
            except TimeoutException:
                print("Could not find 'Näytä tuotteet' button for", card.text)
                continue
            wait = WebDriverWait(driver, 5)
            try:
                products_list_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@data-testid='offer-products']")))
                product_elements = driver.find_elements(By.XPATH, "//ul[@data-testid='offer-products']//li[@data-testid='product-card']")
            except TimeoutException:
                print("Product elements not found or not visible for", card.text)
                continue
            for product in product_elements:
                ean_code = product.get_attribute("data-product-id")
                url = product.find_element(By.XPATH, ".//a[@data-testid='product-link']").get_attribute("href")
                try:
                    unit_price = product.find_element(By.XPATH, ".//div[@data-testid='product-unit-price']").text
                    backslash_index = unit_price.find("/")
                    if backslash_index != -1:
                        unit_type = unit_price[backslash_index+1:]
                        unit_price = unit_price[:backslash_index]
                        unit_price = float(unit_price.replace(',', '.'))
                    else:
                        unit_price = float(unit_price.replace(',', '.'))
                        unit_type = 'Unknown'
                except NoSuchElementException:
                    try:
                        price_element_integer = product.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__IntegerPart')]")
                        unit_price_integer = price_element_integer.text
                
                        price_element_decimal = product.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__DecimalPart')]")
                        unit_price_decimal = price_element_decimal.text

                        unit_element = product.find_element(By.XPATH, ".//div[starts-with(@class, 'ProductPrice__Extra')]")
                        unit_type = unit_element.text.replace('/', '')
                
                        unit_price = float(unit_price_integer + '.' + unit_price_decimal)
                    except Exception as e:
                        print("Could not get product price for", ean_code, e)
                        continue
                if product_dict.get(ean_code, "None") != "None":
                    if product_dict[ean_code]['Price per Unit'] > unit_price:
                        product_dict[ean_code]['Price per Unit'] =  unit_price
                        product_dict[ean_code]['Store'] = store_name
                else:
                    if nutritional_content_dict.get(ean_code, 'Unknown') != 'Unknown':
                        product_dict[ean_code] = {}
                        product_dict[ean_code]['Price per Unit'] = unit_price
                        product_dict[ean_code]['Unit'] = unit_type
                        product_dict[ean_code]['Store'] = store_name
                        product_dict[ean_code].update(nutritional_content_dict[ean_code])
                    else:
                        product_urls.append(url)
            driver.find_element(By.XPATH, "//button[@title='Sulje']").click()

    product_urls = list(set(product_urls))  # Remove duplicates
    for product_url in product_urls:
        if product_url is not None:
            driver.get(product_url)
        else:
            print("Could not open link")
            continue
        try:
            wait = WebDriverWait(driver, 5)
            header = wait.until(EC.visibility_of_element_located((By.XPATH, "//h1[@data-testid='product-name']")))
            product_name = header.text
            driver.execute_script("document.body.style.zoom='50%'")
        except TimeoutException:
            print("Product name not found for", product_url)
            continue
        try:
            wait = WebDriverWait(driver, 5)
            price_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//h1[@data-testid='product-name']//div[@data-testid='product-unit-price']")))
            unit_price = price_element.text
            backslash_index = unit_price.find("/")
            if backslash_index != -1:
                unit_type = unit_price[backslash_index+1:]
                unit_price = unit_price[:backslash_index]
                unit_price = float(unit_price.replace(',', '.'))
            else:
                unit_price = float(unit_price.replace(',', '.'))
                unit_type = 'Unknown'
        except TimeoutException:
            try:
                price_element_integer = driver.find_element(By.XPATH, "//div[@data-testid='product-details-sidebar']//div[starts-with(@class, 'ProductPrice__IntegerPart')]")
                unit_price_integer = price_element_integer.text
                
                price_element_decimal = driver.find_element(By.XPATH, "//div[@data-testid='product-details-sidebar']//div[starts-with(@class, 'ProductPrice__DecimalPart')]")
                unit_price_decimal = price_element_decimal.text

                unit_element = driver.find_element(By.XPATH, "//div[@data-testid='product-details-sidebar']//div[starts-with(@class, 'ProductPrice__Extra')]")
                unit_type = unit_element.text.replace('/', '')
                
                unit_price = float(unit_price_integer + '.' + unit_price_decimal)
            except Exception as e:
                print("Could not get product price for", product_name, e)
                continue
        try:
            wait = WebDriverWait(driver, 5)
            product_info_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//h2[text()='Tuotetiedot']")))
            product_info_header.click()
        except TimeoutException:
            print("Product info header not found for", product_name)
            continue
        try:
            wait = WebDriverWait(driver, 5)
            ean_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//h3[text()='EAN-koodi']//following-sibling::p")))
            ean_code = ean_element.text
        except TimeoutException:
            print("EAN code not found for", product_name)
            continue
        if product_dict.get(ean_code, 'Unknown') != 'Unknown':
            if product_dict[ean_code]['Price per Unit'] > unit_price:
                product_dict[ean_code]['Price per Unit'] = unit_price
                product_dict[ean_code]['Store'] = store_name
            continue
        product_info_header.click()

        product_dict[ean_code] = {}
        product_dict[ean_code]['Name'] = product_name
        product_dict[ean_code]['Price per Unit'] = unit_price
        product_dict[ean_code]['Unit'] = unit_type
        product_dict[ean_code]['Store'] = store_name
         
        try:
            wait = WebDriverWait(driver, 5)
            nutritional_content_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//h2[text()='Ravintosisältö']")))
            time.sleep(1)
            nutritional_content_header.click()

            keys_list = []
            values_list = []

            try:
                wait = WebDriverWait(driver, 5)
                table = wait.until(EC.visibility_of_element_located((By.XPATH, "//h2[(text()='Ravintosisältö')]//parent::button//following::table[@class='NewNutritionalDetails__Table-sc-1sztb3r-4 eGGmNt']")))
                unit_size = table.find_element(By.XPATH, ".//th[@class='NewNutritionalDetails__NutritionContentTableColumnHeading-sc-1sztb3r-7 gACOYE']//div").text
                product_dict[ean_code]['Nutritional Value per'] = unit_size
                keys = table.find_elements(By.XPATH, ".//*[self::th[@class!='NewNutritionalDetails__NutritionContentTableColumnHeading-sc-1sztb3r-7 gACOYE']]")
                values = table.find_elements(By.XPATH, ".//*[self::td[@class!='NewNutritionalDetails__NutritionContentTableColumnHeading-sc-1sztb3r-7 gACOYE']]")
                for key in keys:
                    tokens = key.text.split('\n')
                    for token in tokens:
                        keys_list.append(token)
                for value in values:
                    tokens = value.text.split('\n')
                    for token in tokens:
                        kcal_index = token.find("kcal")
                        backslash_index = token.find("/")
                        if kcal_index != -1:
                            value_string = token[backslash_index+1:kcal_index].replace(',', '.').strip()
                            values_list.append(value_string)
                        else:
                            value_string = token.replace(',', '.').replace('g', '').replace('kJ', '').strip()
                            values_list.append(value_string)
                kv_pairs = dict(zip(keys_list, values_list))
                product_dict[ean_code].update(kv_pairs)
                if nutritional_content_dict.get(ean_code, 'Unknown') == 'Unknown':
                    nutritional_content_dict[ean_code] = {}
                    nutritional_content_dict[ean_code]['Name'] = product_name
                    nutritional_content_dict[ean_code]['Nutritional Value per'] = unit_size
                    nutritional_content_dict[ean_code].update(kv_pairs)
            except TimeoutException:
                print("Nutritional information not found for", product_name)
        except TimeoutException:
            print("Nutritional content header not found for", product_name)
    driver.get("https://www.k-ruoka.fi/?kaupat&kauppahaku=Tampere")

# Serializing json
nutritional_content_json = json.dumps(nutritional_content_dict, indent=4)

# Writing to sample.json
with open("nutritional_content_data.json", "w") as outfile:
    outfile.write(nutritional_content_json)

current_time = datetime.now()
file_name = f"{current_time.strftime('%d')}_{current_time.strftime('%b')}_product_unit_prices.xlsx"

product_data_df = pd.DataFrame.from_dict(product_dict, orient='index')
product_data_df['Proteiini'] = product_data_df['Proteiini'].apply(lambda x: float(x) if isinstance(x, (float, int,  str)) else 0)
product_data_df['Euroa per 100g Proteiinia'] = ((product_data_df['Price per Unit'] / 10.00) * (100 / product_data_df['Proteiini']))
product_data_df.index.name = 'EAN-code'

with pd.ExcelWriter(f"{file_name}") as writer:
    product_data_df.to_excel(writer, sheet_name='Products')

sorted_df = product_data_df.sort_values(by=['Euroa per 100g Proteiinia'], ascending=[True])
display(sorted_df.head(10))

    