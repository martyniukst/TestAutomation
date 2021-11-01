from selenium.webdriver.support.ui import Select
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import xlsxwriter
from bs4 import BeautifulSoup
import os

ROOT_DIR = os.path.dirname(os.path.abspath("main.py"))

chrome_options = Options()
chrome_options = selenium.webdriver.ChromeOptions()
prefs = {"download.default_directory" : str(ROOT_DIR)+'/output'}
chrome_options.add_experimental_option("prefs",prefs)
driver = webdriver.Chrome(ChromeDriverManager().install(),options=chrome_options)


workbook = xlsxwriter.Workbook('Agencies.xlsx')
worksheet = workbook.add_worksheet('Agencies')
def scraping():
    try:
        driver.get('https://itdashboard.gov/')
        time.sleep(3)
        btn = driver.find_element(By.XPATH, '//a[@href="#home-dive-in"]')
        btn.click()
        agencies = driver.find_elements(By.CLASS_NAME, 'h4.w200')
        spending = driver.find_elements(By.CLASS_NAME, 'h1.w900')
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

        #get page National Science Foundation
        worksheet2 = workbook.add_worksheet('National_Science_Foundation')
        driver.get('https://itdashboard.gov/drupal/summary/422')
        time.sleep(10)
        select = Select(driver.find_element(By.NAME, 'investments-table-object_length'))
        select.select_by_visible_text('All')
        time.sleep(10)
        table = driver.find_element(By.CLASS_NAME, "dataTables_scrollBody")
        code = table.get_attribute('outerHTML')
        soup = BeautifulSoup(code, 'html.parser')
        target = soup.find_all('tr')
        worksheet2.write('A1', 'UII')
        worksheet2.write('B1', 'Bureau')
        worksheet2.write('C1', 'Investment Title')
        worksheet2.write('D1', 'Total FY2021 Spending ($M)')
        worksheet2.write('E1', 'Type')
        worksheet2.write('F1', 'CIO Rating')
        worksheet2.write('G1', '# of Projects')
        worksheet2.write('H1', 'Href')
        i = -1
        for item in target:
            res = []
            for elem in item.find_all('td', style=False):
                res.append(elem.text)
                if 'href' in str(elem):
                    url = 'https://itdashboard.gov'+str(elem).split('"')[3]
                    res.append(url)
                    driver.get(url)
                    time.sleep(3)
                    pdf = driver.find_element(By.LINK_TEXT, "Download Business Case PDF")
                    pdf.click()
                    time.sleep(10)
            try:
                if len(res)>7:
                    res.append(res[1])
                    res.pop(1)
            except:
                pass
            for idx, month in enumerate(res):
                worksheet2.write(i, idx, res[idx])
            i = i+1
        workbook.close()
    finally:
        driver.close()

if __name__ == "__main__":
    scraping()
