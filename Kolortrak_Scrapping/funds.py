from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
import undetected_chromedriver as uc
import csv
import time

Path = "chromedriver.exe"
service = Service(Path)
user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'

cho = Options()
cho.add_argument("--disable-dev-shm-usage")
cho.add_argument("disable-infobars")
cho.add_argument("--disable-extensions")
cho.add_argument("--no-sandbox")
cho.add_argument(f"user-agent={user_agent}")

driver = uc.Chrome(options=cho, service=service)

username = "ahalabi@acceinfo.com"
password = "Halabi15$"

driver.get("https://app.kolortrak.com/#/signin")

username_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located(
        (
            By.XPATH,
            "/html/body/div[1]/div[2]/div[1]/div/div[2]/div[2]/form/div[1]/input",
        )
    )
)
password_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located(
        (
            By.XPATH,
            "/html/body/div[1]/div[2]/div[1]/div/div[2]/div[2]/form/div[2]/input",
        )
    )
)

username_input.send_keys(username)
password_input.send_keys(password)

login_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (
            By.XPATH,
            "/html/body/div[1]/div[2]/div[1]/div/div[2]/div[2]/form/div[4]/button",
        )
    )
)
login_button.click()

dashboard_element = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/div[1]/div/div/div[2]/div/div"))
)

funds = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[1]/div[2]/div[1]/div/div/div[2]/div/div/div/ul/li[3]/div")
        
    )
)
funds.click()


funds_research = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[1]/div[2]/div[1]/div/div/div[2]/div/div/div/ul/li[3]/div/div/button[1]")
    )
)
funds_research.click()

dashboard_container = WebDriverWait(driver, 100).until(
    EC.presence_of_element_located(
        (
            By.XPATH,
            "/html/body/div[1]/div[2]/div[1]/div/main/div/div/div/div/div[2]/div[2]/div[1]/div",
        )
    )
)

dashboard = WebDriverWait(dashboard_container, 200).until(
    EC.visibility_of_element_located(
        (
            By.XPATH,
            "/html/body/div[1]/div[2]/div[1]/div/main/div/div/div/div/div[2]/div[2]/div[1]/div",
        )
    )
)


setting = WebDriverWait(driver, 200).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[1]/div[2]/div[1]/div/main/div/div/div/div/div[2]/div[1]/div[3]/button")
    )
)
setting.click()


element_click = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CLASS_NAME,"popover-body"))
)

time.sleep(2)


check_boxes = [
    '#id-toggleCheckBox',
    '#id-manager',
    '#id-fundseries',
    '#id-ESG',
    '#id-minimumInvestment',
    '#id-holdingNumber',
    '#id-feesAndSalesInfo',
    '#id-alpha1y',
    '#id-alpha3y',
    '#id-alpha5y',
    '#id-beta1y',
    '#id-beta3y',
    '#id-beta5y',
    '#id-sharpe1y',
    '#id-sharpe3y',
    '#id-sharpe5y',
    '#id-sortino1y',
    '#id-sortino3y',
    '#id-sortino5y',
    '#id-infoRatio1y',
    '#id-infoRatio3y',
    '#id-infoRatio5y',
    '#id-growth3m',
    '#id-growth6m',
    '#id-growth3y',
    '#id-growth5y',
    '#id-growth10y',
    '#id-growthGraph',
    '#id-risk1y',
    '#id-risk3y',
    '#id-risk5y',
    '#id-risk10y',
    '#id-assetBond',
    '#id-assetStock',
    '#id-sector_45',
    '#id-sector_20',
    '#id-sector_10',
    '#id-sector_40',
    '#id-sector_60',
    '#id-sector_55',
    '#id-sector_30',
    '#id-sector_15',
    '#id-sector_35',
    '#id-sector_25',
    '#id-sector_50'
]

for button in check_boxes:
    check = driver.find_element(By.CSS_SELECTOR, button)
    driver.execute_script("arguments[0].scrollIntoView();", check)
    time.sleep(0.2)
    driver.execute_script("arguments[0].click();", check)

    
time.sleep(1)
setting.click()


hide_sidebar = driver.find_element(By.XPATH,'/html/body/div[1]/div[2]/div[1]/div/main/div/div/div/div/div[1]/div[2]/button')
hide_sidebar.click()


time.sleep(2)


 



def scroll_and_scrape_table(driver, csv_filename):
    table_data = []

    with open(csv_filename, "w", newline="", encoding="utf-8") as csvfile:
        csv_writer = csv.writer(csvfile)

     
        count = 39483
        for i in range(1,count+1):
            try:                
                while True:
                    try:
                        
                        grid = driver.find_element(By.CLASS_NAME,'ReactVirtualized__Grid__innerScrollContainer').find_elements(By.CLASS_NAME,'ReactVirtualized__Table__row')
                        last_element = grid[-1]
                        last_element_rowindex=last_element.get_attribute('aria-rowindex')                
                        
                        sabih_element = driver.find_element(By.CSS_SELECTOR,f'div[aria-rowindex="{i}"]')
                        haseebs_element=sabih_element.find_elements(By.CLASS_NAME, "ReactVirtualized__Table__rowColumn")
                        data = [column.text.strip() for column in haseebs_element]
                        csv_writer.writerow(data)
                        #print('i',i)
                        #print('last_element_rowindex',last_element_rowindex)
                        if str(i)==last_element_rowindex :
                            print('we scraped last one now')
                            print(last_element)
                            driver.execute_script("arguments[0].scrollIntoView();", last_element)
                        
                        break
 
                    except:
                        continue
                

            except Exception as e:
                print(f"Error: {str(e)}")
                break


csv_filename = "output.csv"

table_data = scroll_and_scrape_table(driver, csv_filename)
