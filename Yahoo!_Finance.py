from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd
import numpy as np


driver = webdriver.Chrome()
driver.maximize_window()

# explicit wait
wait = WebDriverWait(driver,10)

def wait_for_page_load(driver,wait):
    page_title = driver.title
    try:
        wait.until(
            lambda d : d.execute_script("return document.readyState") == "complete"
        )
    except:
        print(f"page is \"{page_title}\" not loaded in given time")
    else:
        print(f"Page \"{page_title}\" is Loaded Sucessfully")

   

url = "https://finance.yahoo.com/"
driver.get(url)
wait_for_page_load(driver,wait)


# hovering on market menu

actions = ActionChains(driver)

market_menu = wait.until(
    EC.presence_of_element_located((By.XPATH , '/html/body/div[2]/header/div/div/div/div[4]/div/div/ul/li[3]/a'))
)
actions.move_to_element(market_menu).perform()

# click On Trending ticker

trending_ticker = wait.until(
    EC.element_to_be_clickable(
    (By.XPATH, '//a[contains(@href,"trending-tickers")]')
)
)

trending_ticker.click()
wait_for_page_load(driver,wait)

# Click on mst active

most_active = wait.until(
    EC.presence_of_element_located((By.XPATH , '/html/body/div[2]/main/section/section/section/article/section[1]/div/nav/ul/li[1]/a'))
)
most_active.click()
wait_for_page_load(driver,wait)


# Scraping data

data = []
while True:
    # scraping
    
    wait.until(
        EC.presence_of_element_located((By.TAG_NAME,'table'))
    )
    rows = driver.find_elements(By.CSS_SELECTOR,'table tbody tr')
    for row in rows:
        values = row.find_elements(By.TAG_NAME,'td')
        stoke = {
            'name' : values[1].text,
            'symbol' : values[0].text,
            'price' : values[3].text,
            'change' : values[4].text,
            'volume' : values[6].text,
            'market_cap' : values[8].text,
            'PE_ratio' : values[9].text,
        }
        data.append(stoke)


    # checking Clickable

    # click next page

    try:
        next_button = wait.until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="nimbus-app"]/section/section/section/article/section[1]/div/div[3]/div[3]/button[3]'))
        )

        if next_button.get_attribute("disabled"):
            print("Next button is disabled. No more pages.")
            break
        
    except:
        print('Next Page Is not available we go through all pages')
        break
    else:
        next_button.click()
        time.sleep(2)


stoke_table = pd.DataFrame(data).apply(lambda col:col.str.strip() if col.dtype == "object" else col).assign(
    price = lambda df_: pd.to_numeric(df_.price),
    change = lambda df_: pd.to_numeric(df_.change.str.replace('+','')),
    volume = lambda df_: pd.to_numeric(df_.volume.str.replace("M","")),
    market_cap = lambda df_:df_.market_cap.apply(lambda val : float(val.replace("B","")) if "B" in val else float(val.replace("T","")) * 1000),
    PE_ratio = lambda df_:(df_.PE_ratio.replace("--",np.nan).str.replace(",","").pipe(lambda col:pd.to_numeric(col)))
).rename(columns = {
    "price" : "Price($)",
    "volume" : "Volume(Million)",
    "market_cap" : "Market_cap(Billon)"
})



stoke_table


get_ipython().system('pip install openpyxl')

stoke_table.to_excel("Yahoo-Stoke-data.xlsx", index=False) 





