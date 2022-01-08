from selenium import webdriver
import time
from openpyxl import Workbook, load_workbook

web = webdriver.Chrome('/Users/genesissales/PycharmProjects/RobotChulllienForm/venv/lib/python3.8/site-packages/selenium/chromedriver')
web.get('https://forms.office.com/Pages/DesignPage.aspx?auth_pvr=WindowsLiveId&auth_upn=genxsolarservices%40gmail.com&lang=en-PH&origin=OfficeDotCom&route=Start#Analysis=true&FormId=DQSIkWdsW0yxEjajBLZtrQAAAAAAAAAAAANAAT2kdaNUODY4NVBROU5JT1FIQk84VU1EUjZQTFlWMi4u')

time.sleep(2)

# Enter Password of microsoft
Password = '09058714319Gen$1'
Passxpath = web.find_element_by_xpath('//*[@id="i0118"]')
Passxpath.send_keys(Password)

# Sign in
Signin=web.find_element_by_xpath('//*[@id="idSIButton9"]')
Signin.click()
Yes=web.find_element_by_xpath('//*[@id="idSIButton9"]')
Yes.click()

time.sleep(1)

#Download excel File
Yes=web.find_element_by_xpath('//*[@id="analyzeViewPrintChildContainer"]/div[3]/div[2]/div/button/div[1]/span')
Yes.click()

time.sleep(3)
#Open excel file
wb = load_workbook('/Users/genesissales/Downloads/Chulien Pre-order Form (3).xlsx')
ws = wb.active