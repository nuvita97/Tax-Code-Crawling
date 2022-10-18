from selenium import webdriver
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import openpyxl

# Function to get value from excel file
def get_value_excel(filename, cellname):
    wb = openpyxl.load_workbook(filename)
    Sheet1 = wb['Sheet1']
    wb.close()
    return Sheet1[cellname].value

# Function to update value to excel file
def update_value_excel(filename, cellname, value):
    wb = openpyxl.load_workbook(filename)
    Sheet1 = wb['Sheet1']
    Sheet1[cellname].value = value
    wb.close()
    wb.save(filename)

# Columns to update value
col_name_mst = "A"
col_name_company = "B"
col_name_address = "C"

filename = "Data-customer.xlsx"
exception = "N/A"
start_row = 2
end_row = 2263

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get("https://masothue.com/")

# Update to cell
for i_row in range (start_row, end_row):
    cell_name_mst = "%s%s"%(col_name_mst, i_row)
    cell_name_company = "%s%s"%(col_name_company, i_row)
    cell_name_address = "%s%s"%(col_name_address, i_row)

    mst = get_value_excel(filename, cell_name_mst)

    # Fill in value to Search
    mstSearch = driver.find_element("id", "search")
    sleep(1)
    mstSearch.send_keys(mst)
    sleep(1)

    # Click Search
    mstSeachBtn = driver.find_element("xpath", "/html/body/div[1]/nav/div/form/div/div[2]/button")
    sleep(1)
    mstSeachBtn.click()
    sleep(1)
        

    # Place to get values from web
    try:
        mstSearchCompany = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div[2]/main/section[1]/div/table[1]/thead/tr/th/span")
        sleep(1) 
        mstSearchAddress = driver.find_element(By.CSS_SELECTOR, "td[itemprop='address']")
        sleep(1)
    except:
        update_value_excel(filename, cell_name_company, exception)
        update_value_excel(filename, cell_name_address, exception)
        print(cell_name_company, exception)
        # Restart windows
        # driver.close()
        # sleep(1)
        # driver = webdriver.Chrome(ChromeDriverManager().install())
        # driver.get("https://masothue.com/")
    else:
        update_value_excel(filename, cell_name_company, mstSearchCompany.text)
        update_value_excel(filename, cell_name_address, mstSearchAddress.text)
        print(cell_name_company, mstSearchCompany.text)

driver.close()
sleep(1)
