from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd

driver = webdriver.Chrome()

driver.get('https://itdashboard.gov/')
# Get button 'Dive in' by XPATH and click
dive_in = driver.find_element(By.XPATH, '/html/body/main/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/a')
dive_in.click()
del dive_in

# 'Lazy load' site problem
driver.implicitly_wait(2)

# Get amounts of agencies
amounts = driver.find_elements_by_css_selector('span.h1.w900')
amounts_list = []
for amount in amounts:
    amounts_list.append([amount.text])
df = pd.DataFrame(amounts_list)
writer = pd.ExcelWriter('output/table.xlsx')
# Put data into Excel file
df.to_excel(writer, sheet_name='Agencies', index=False, header=False)
del (amounts, amounts_list, df)

# Get 'view' button by XPATH and click on it
view = driver.find_element(By.XPATH,
                           '/html/body/main/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[1]/div[2]/div/div/div/div[2]/a')
view.click()
del view

# Again 'lazy load'
driver.implicitly_wait(10)

# Get table of investments and put in on new sheet in the Excel file
table = driver.find_element(By.ID, 'investments-table-object')
table = table.get_attribute('outerHTML')
df = pd.read_html(table)[0]
df.to_excel(writer, sheet_name='Agency', index=False, header=False)
writer.save()
writer.close()

# Get links in UII column
table = driver.find_element(By.ID, 'investments-table-object')
rows = table.find_element(By.TAG_NAME, 'tbody')
rows = rows.find_elements(By.TAG_NAME, 'tr')
col = rows[0].find_element(By.TAG_NAME, 'td')
a = col.find_element(By.TAG_NAME, 'a')
if a.get_attribute('href'):
    a.click()
    # Download PDF file
    download = driver.find_element(By.ID, 'business-case-pdf')
    download = download.find_element(By.TAG_NAME, 'a')
    download.click()
    # Lets the file be downloaded
    sleep(10)
    del download
del (table, rows, col, a, writer)

driver.quit()
