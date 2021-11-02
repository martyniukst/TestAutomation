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

def open_url(url):
    driver.get(url)

def find_element(by, term):
    return driver.find_element(by, term)

def find_elements(by, term):
    return driver.find_elements(by, term)

def save_pdf(by, term):
    driver.find_element(by, term).click()


def main():
    try:
        open_url('https://itdashboard.gov/')
        time.sleep(3)
        find_element(By.XPATH, '//a[@href="#home-dive-in"]').click()
        agencies = find_elements(By.CLASS_NAME, 'h4.w200')
        spending = find_elements(By.CLASS_NAME, 'h1.w900')
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
        open_url('https://itdashboard.gov/drupal/summary/422')
        time.sleep(10)
        select = Select(driver.find_element(By.NAME, 'investments-table-object_length'))
        select.select_by_visible_text('All')
        time.sleep(10)
        table = find_element(By.CLASS_NAME, "dataTables_scrollBody")
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
                    open_url(url)
                    time.sleep(3)
                    save_pdf(By.LINK_TEXT, "Download Business Case PDF")
                    time.sleep(10)
            try:
                if len(res)>7:
                    res.append(res[1])
                    res.pop(1)
            except:
                pass #without href
            for idx, month in enumerate(res):
                worksheet2.write(i, idx, res[idx])
            i = i+1
        workbook.close()
    finally:
        driver.close()

def parse_pdf():
    import pdfplumber
    from os import listdir
    from os.path import isfile, join
    import pandas as pd
    df = pd.ExcelFile('Agencies.xlsx').parse('National_Science_Foundation')
    uii=[]
    for elem in df['UII']:
        uii.append(elem)
    investment=[]
    for elem in df['Investment Title']:
        investment.append(elem)
    mypath = 'output'
    files = [f for f in listdir(mypath) if isfile(join(mypath, f))]
    for file in files:
        with pdfplumber.open(r'output/'+str(file)) as pdf:
            first_page = pdf.pages[0]
            list = first_page.extract_text().splitlines()
            for item in list:
                if 'Name of this Investment' in item:
                    print (item.split(': ')[1] in investment)
                elif 'Unique Investment Identifier (UII)' in item:
                    print (item.split(': ')[1] in uii)



if __name__ == "__main__":
    main()
    parse_pdf()