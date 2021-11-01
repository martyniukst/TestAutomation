import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import xlsxwriter
# from bs4 import BeautifulSoup
# import csv

chrome_options = Options()
options = selenium.webdriver.ChromeOptions()
driver = webdriver.Chrome(ChromeDriverManager().install(),options=chrome_options)

workbook = xlsxwriter.Workbook('Agencies.xlsx')
worksheet = workbook.add_worksheet('Agencies')
def scraping():
    driver.get('https://itdashboard.gov/')
    time.sleep(3)
    btn = driver.find_element(By.XPATH, '//a[@href="#home-dive-in"]')
    btn.click()
    agencies = driver.find_elements(By.CLASS_NAME, 'h4.w200')
    spending = driver.find_elements(By.CLASS_NAME, 'h1.w900')
    list_agencies = []
    list_spending = []
    i = 1
    for item in agencies:
        if item.text != '':
            worksheet.write('A'+str(i), item.text)
            i = i+1
    i = 1
    for item in spending:
        if item.text != '':
            worksheet.write('B'+str(i), item.text)
            i = i+1
    print (list_agencies, list_spending)
    workbook.close()
    driver.close()

if __name__ == "__main__":
    scraping()
