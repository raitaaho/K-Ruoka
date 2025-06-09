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

driver = uc.Chrome()
driver.get("https://www.k-ruoka.fi/?kaupat&kauppahaku=Tampere")
wait = WebDriverWait(driver, 10)
try:
    accept_cookies = wait.until(EC.element_to_be_clickable((By.XPATH, f"//button[@id='onetrust-accept-btn-handler']")))
    # Click accept cookies button
    accept_cookies.click()
except TimeoutException:
    print('Prompt to accept cookies did not pop up')
time.sleep(5)

number_of_stores = len(driver.find_elements(By.XPATH, "//li[@data-component='store-list-item']"))
counter = 0
product_dict = {}
while counter < number_of_stores:
    try:
        stores = driver.find_elements(By.XPATH, "//li[@data-component='store-list-item']")
        store_name = stores[counter].get_attribute("data-store")
        store = stores[counter].find_element(By.XPATH, f"//button[@data-select-store='{store_name}']")

        store.click()
        time.sleep(3)
        counter += 1
        driver.get("https://www.k-ruoka.fi/kauppa/tuotehaku/liha-ja-kasviproteiinit")
        time.sleep(3)
        driver.execute_script("document.body.style.zoom='50%'")
        time.sleep(1)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)
        product_cards = driver.find_elements(By.XPATH, "//a[@class='ProductCard__ProductLink-sc-12u3k8m-4 dTXQcz']")
        product_urls = [card.get_attribute("href") for card in product_cards if card.get_attribute("href")]
    except Exception as e:
        print("Could not fetch product urls")
        continue

    for product_url in product_urls:
        if product_url is not None:
            driver.get(product_url)
            time.sleep(3)
            driver.execute_script("document.body.style.zoom='50%'")
            time.sleep(1)
        else:
            print("Could not open link")
            continue
        header = driver.find_element(By.XPATH, "//h1[@data-testid='product-name']")
        product_name = header.text
        
        try:
            price_element = driver.find_element(By.XPATH, "//div[@data-testid='product-unit-price']")
            unit_price = price_element.text
            backslash_index = unit_price.find("/")
            if backslash_index != -1:
                unit_price = unit_price[:backslash_index]
        except NoSuchElementException:
            try:
                price_element_integer = driver.find_element(By.XPATH, "//div[@class='ProductPrice__IntegerPart-sc-u2ag1v-6 cIkeLN']")
                unit_price_integer = price_element_integer.text
                
                price_element_decimal = driver.find_element(By.XPATH, "//div[@class='ProductPrice__DecimalPart-sc-u2ag1v-7 eolJwY']")
                unit_price_decimal = price_element_decimal.text
                
                unit_price = unit_price_integer + '.' + unit_price_decimal
            except Exception as e:
                print("Could not get product price")
                continue
        
        if product_dict.get(product_name, 'None') == 'None':
            product_dict[product_name] = {}
            product_dict[product_name]['Price per Kilogram'] = float(unit_price.replace(',', '.'))
            product_dict[product_name]['Store'] = store_name
        else:
            if float(unit_price.replace(',', '.')) < product_dict[product_name]['Price per Kilogram']:
                product_dict[product_name]['Price per Kilogram'] = float(unit_price.replace(',', '.'))
                product_dict[product_name]['Store'] = store_name

        try:
            nutritional_content_header = driver.find_element(By.XPATH, "//h2[text()='Ravintosisältö']")
            nutritional_content_header.click()
            time.sleep(5)

            keys_list = []
            values_list = []
            table = driver.find_element(By.XPATH, "//h2[(text()='Ravintosisältö')]//parent::button//following::table[@class='NewNutritionalDetails__Table-sc-1sztb3r-4 eGGmNt']")
            unit_size = table.find_element(By.XPATH, "//descendant::th[@class='NewNutritionalDetails__NutritionContentTableColumnHeading-sc-1sztb3r-7 gACOYE']//descendant::div").text
            product_dict[product_name]['Unit Size'] = unit_size
            for row in table.find_elements(By.XPATH, "//tr"):
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
                            values_list.append(token[backslash_index+1:kcal_index].strip())
                        else:
                            values_list.append(token.replace(',', '.').replace('g', '').strip())
            kv_pairs = dict(zip(keys_list, values_list))
            product_dict[product_name].update(kv_pairs)
             
        except NoSuchElementException:
            print("Nutritional information not found for", product_name)
    driver.get("https://www.k-ruoka.fi/?kaupat&kauppahaku=Tampere")
    time.sleep(3)

product_data_df = pd.DataFrame.from_dict(product_dict, orient='index')
product_data_df['Euroa per 80g Proteiinia'] = ((product_data_df['Price per Kilogram'] / 1000) * (80 / (product_data_df['Proteiini'] / 100)))
product_data_df.index.name = 'Product'

with pd.ExcelWriter(f"product_prices.xlsx") as writer:
    product_data_df.to_excel(writer, sheet_name='Products')

sorted_df = product_data_df.sort_values(by=['Euroa per 80g Proteiinia'], ascending=[True])
display(sorted_df.head(10))

    