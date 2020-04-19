from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import re
import pandas as pd
import os
import time
import openpyxl
os.chdir('c:\\Python\\colefitzpatrick_python')

url = "https://www.fangraphs.com/projections.aspx?pos=all&stats=bat&type=steamer"
pages = 131       #adjust this number based on the number of players

wb_write = openpyxl.load_workbook('steamer.xlsx')
ws_write = wb_write["hitters"]

# create a new Firefox session
driver = webdriver.Firefox()
driver.implicitly_wait(8)
driver.get(url)
writerow = 2
for page in range(1,pages):         #can use range(1,x) if you only want x number of pages and not the full dataset
    python_button = driver.find_element_by_css_selector('#ProjectionBoard1_dg1_ctl00 > thead:nth-child(2) > tr:nth-child(1) > td:nth-child(1) > div:nth-child(1) > div:nth-child(3) > button:nth-child(1)')  #finds the right arrow button to advance to the next page
    soup_level1=BeautifulSoup(driver.page_source, 'lxml')
    
    for tr in soup_level1.findAll("tr", {"class": "rgRow"}):   #the table is structured where each row has an alternating tr-class, the first loop pulls the odds, second loop pulls the even rows
        tds = tr.findAll("td", {"class": "grid_line_regular"})
        writecol=1
        for td in tds:
            ws_write.cell(row=writerow, column=writecol).value = td.text     #writes the values into spreadsheet
            writecol += 1
        writerow += 1
        
    for tr in soup_level1("tr", {"class": "rgAltRow"}):
        tds = tr.findAll("td", {"class": "grid_line_regular"})
        writecol=1
        for td in tds:
            ws_write.cell(row=writerow, column=writecol).value = td.text
            writecol += 1
        writerow += 1

    python_button.click()       #clicks the right arrow to advance to the next page
    driver.implicitly_wait(5)

wb_write.save('steamer.xlsx')
