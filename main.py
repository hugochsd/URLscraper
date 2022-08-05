import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

year = []
month = []
description = []
urls = []
df = pd.read_excel('CISARSS.xlsx', sheet_name='mysheet')

year = df['Year'].tolist()
month = df['Month'].tolist()
description = df['Description'].tolist()

mydesc = []

for de in description:
    if de != None:
        mydesc.append(de.partition('\n')[0])

string = "article"
print(len(description))
print(len(month))
driver = webdriver.Chrome(executable_path="C:\chromedriver.exe")

for desc in mydesc:
    driver.get('http://www.google.com')

    searchbox = driver.find_element('name','q')
    searchbox.send_keys(desc)
    searchbox.submit()

    driver.find_element(By.XPATH, '(//h3)[1]/../../a').click()
    #time.sleep(1)
    print(driver.current_url)
    urls.append(driver.current_url)

driver.quit()


mydata = pd.DataFrame({'Year':year, 'Month':month, 'Link': urls, 'Description': description})
writer = pd.ExcelWriter('mapdata.xlsx')
mydata.head(627).to_excel(writer, sheet_name='mysheet')
workbook = writer.book
worksheet = writer.sheets['mysheet']
worksheet.set_column(1,1,20)
worksheet.set_column(2,2,30)
worksheet.set_column(3,3,120)
worksheet.set_column(4,4,120)

writer.save()