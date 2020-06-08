from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import re
import pandas as pd
import os
import time
import openpyxl
os.chdir('c:\\Python\\colefitzpatrick_python\\FantasyBaseballProjection')

url = "https://www.fangraphs.com/projections.aspx?pos=all&stats=bat&type=steamer"
url2 = "https://www.fangraphs.com/projections.aspx?pos=all&stats=pit&type=steamer&team=0&lg=all&players=0"
pages = 130       #adjust this number based on the number of players

wb_write = openpyxl.load_workbook('steamer.xlsx')
ws_write = wb_write["hitters"]
ws2_write = wb_write["pitchers"]
wb3 = openpyxl.load_workbook('fantrax.xlsx')
ws3 = wb3["Sheet1"]

# create a new Firefox session
driver = webdriver.Firefox()
driver.implicitly_wait(8)
driver.get(url)
writerow = 2
for page in range(1,pages):         #can use range(1,x) if you only want x number of pages and not the full dataset   
    rightbutton = driver.find_element_by_xpath("//button[@class='t-button rgActionButton rgPageNext']")
    soup_level1=BeautifulSoup(driver.page_source, 'lxml')
    
    for tr in soup_level1.findAll("tr", {"class": "rgRow"}):   #the table is structured where each row has an alternating tr-class, the first loop pulls the odds, second loop pulls the even rows
        tds = tr.findAll("td", {"class": "grid_line_regular"})
        writecol=1
        for td in tds:
            if writecol < 4:
                ws_write.cell(row=writerow, column=writecol).value = td.text     #writes the name and team values into spreadsheet
                writecol += 1
            else:
                ws_write.cell(row=writerow, column=writecol).value = float(td.text)     #writes the statistical values into spreadsheet
                writecol += 1
        writerow += 1
        
    for tr in soup_level1("tr", {"class": "rgAltRow"}):
        tds = tr.findAll("td", {"class": "grid_line_regular"})
        writecol=1
        for td in tds:
            if writecol < 4:
                ws_write.cell(row=writerow, column=writecol).value = td.text     #writes the name and team values into spreadsheet
                writecol += 1
            else:
                ws_write.cell(row=writerow, column=writecol).value = float(td.text)     #writes the statistical values into spreadsheet
                writecol += 1
        writerow += 1

    driver.execute_script("arguments[0].click();", rightbutton)       #clicks the right arrow to advance to the next page
    time.sleep(3)

driver.get(url2)
writerow = 2

for page in range(1,pages):         #can use range(1,x) if you only want x number of pages and not the full dataset  
    rightbutton = driver.find_element_by_xpath("//button[@class='t-button rgActionButton rgPageNext']")
    soup_level1=BeautifulSoup(driver.page_source, 'lxml')
    
    for tr in soup_level1.findAll("tr", {"class": "rgRow"}):   #the table is structured where each row has an alternating tr-class, the first loop pulls the odds, second loop pulls the even rows
        tds = tr.findAll("td", {"class": "grid_line_regular"})
        writecol=1
        for td in tds:
            if writecol < 4:
                ws2_write.cell(row=writerow, column=writecol).value = td.text     #writes the name and team values into spreadsheet
                writecol += 1
            else:
                ws2_write.cell(row=writerow, column=writecol).value = float(td.text)     #writes the statistical values into spreadsheet
                writecol += 1
        writerow += 1
        
    for tr in soup_level1("tr", {"class": "rgAltRow"}):
        tds = tr.findAll("td", {"class": "grid_line_regular"})
        writecol=1
        for td in tds:
            if writecol < 4:
                ws2_write.cell(row=writerow, column=writecol).value = td.text     #writes the name and team values into spreadsheet
                writecol += 1
            else:
                ws2_write.cell(row=writerow, column=writecol).value = float(td.text)     #writes the statistical values into spreadsheet
                writecol += 1
        writerow += 1

    driver.execute_script("arguments[0].click();", rightbutton)       #clicks the right arrow to advance to the next page
    time.sleep(3)

hitterrows = ws_write.max_row
pitcherrows = ws2_write.max_row
fantraxrows = ws3.max_row

for row in range(1,hitterrows+1):
    if ws_write.cell(row=row, column=3).value == 'Angels':
        ws_write.cell(row=row, column=2).value = 'LAA'
    elif ws_write.cell(row=row, column=3).value == 'Athletics':
        ws_write.cell(row=row, column=2).value = 'OAK' 
    elif ws_write.cell(row=row, column=3).value == 'Astros':
        ws_write.cell(row=row, column=2).value = 'HOU'
    elif ws_write.cell(row=row, column=3).value == 'Blue Jays':
        ws_write.cell(row=row, column=2).value = 'TOR'
    elif ws_write.cell(row=row, column=3).value == 'Braves':
        ws_write.cell(row=row, column=2).value = 'ATL' 
    elif ws_write.cell(row=row, column=3).value == 'Brewers':
        ws_write.cell(row=row, column=2).value = 'MIL'
    elif ws_write.cell(row=row, column=3).value == 'Cardinals':
        ws_write.cell(row=row, column=2).value = 'STL'
    elif ws_write.cell(row=row, column=3).value == 'Cubs':
        ws_write.cell(row=row, column=2).value = 'CHC' 
    elif ws_write.cell(row=row, column=3).value == 'Diamondbacks':
        ws_write.cell(row=row, column=2).value = 'ARI'
    elif ws_write.cell(row=row, column=3).value == 'Dodgers':
        ws_write.cell(row=row, column=2).value = 'LAD'
    elif ws_write.cell(row=row, column=3).value == 'Giants':
        ws_write.cell(row=row, column=2).value = 'SF'
    elif ws_write.cell(row=row, column=3).value == 'Indians':
        ws_write.cell(row=row, column=2).value = 'CLE' 
    elif ws_write.cell(row=row, column=3).value == 'Mariners':
        ws_write.cell(row=row, column=2).value = 'SEA'
    elif ws_write.cell(row=row, column=3).value == 'Marlins':
        ws_write.cell(row=row, column=2).value = 'MIA'
    elif ws_write.cell(row=row, column=3).value == 'Mets':
        ws_write.cell(row=row, column=2).value = 'NYM' 
    elif ws_write.cell(row=row, column=3).value == 'Nationals':
        ws_write.cell(row=row, column=2).value = 'WSH'
    elif ws_write.cell(row=row, column=3).value == 'Orioles':
        ws_write.cell(row=row, column=2).value = 'BAL'
    elif ws_write.cell(row=row, column=3).value == 'Padres':
        ws_write.cell(row=row, column=2).value = 'SD' 
    elif ws_write.cell(row=row, column=3).value == 'Phillies':
        ws_write.cell(row=row, column=2).value = 'PHI'
    elif ws_write.cell(row=row, column=3).value == 'Pirates':
        ws_write.cell(row=row, column=2).value = 'PIT'
    elif ws_write.cell(row=row, column=3).value == 'Rangers':
        ws_write.cell(row=row, column=2).value = 'TEX' 
    elif ws_write.cell(row=row, column=3).value == 'Rays':
        ws_write.cell(row=row, column=2).value = 'TB'
    elif ws_write.cell(row=row, column=3).value == 'Red Sox':
        ws_write.cell(row=row, column=2).value = 'BOS'
    elif ws_write.cell(row=row, column=3).value == 'Reds':
        ws_write.cell(row=row, column=2).value = 'CIN' 
    elif ws_write.cell(row=row, column=3).value == 'Rockies':
        ws_write.cell(row=row, column=2).value = 'COL'
    elif ws_write.cell(row=row, column=3).value == 'Royals':
        ws_write.cell(row=row, column=2).value = 'KC'
    elif ws_write.cell(row=row, column=3).value == 'Tigers':
        ws_write.cell(row=row, column=2).value = 'DET' 
    elif ws_write.cell(row=row, column=3).value == 'Twins':
        ws_write.cell(row=row, column=2).value = 'MIN'
    elif ws_write.cell(row=row, column=3).value == 'White Sox':
        ws_write.cell(row=row, column=2).value = 'CHW'
    elif ws_write.cell(row=row, column=3).value == 'Yankees':
        ws_write.cell(row=row, column=2).value = 'NYY' 
    else:
        continue

for row in range(1,pitcherrows+1):
    if ws2_write.cell(row=row, column=3).value == 'Angels':
        ws2_write.cell(row=row, column=2).value = 'LAA'
    elif ws2_write.cell(row=row, column=3).value == 'Athletics':
        ws2_write.cell(row=row, column=2).value = 'OAK' 
    elif ws2_write.cell(row=row, column=3).value == 'Astros':
        ws2_write.cell(row=row, column=2).value = 'HOU'
    elif ws2_write.cell(row=row, column=3).value == 'Blue Jays':
        ws2_write.cell(row=row, column=2).value = 'TOR'
    elif ws2_write.cell(row=row, column=3).value == 'Braves':
        ws2_write.cell(row=row, column=2).value = 'ATL' 
    elif ws2_write.cell(row=row, column=3).value == 'Brewers':
        ws2_write.cell(row=row, column=2).value = 'MIL'
    elif ws2_write.cell(row=row, column=3).value == 'Cardinals':
        ws2_write.cell(row=row, column=2).value = 'STL'
    elif ws2_write.cell(row=row, column=3).value == 'Cubs':
        ws2_write.cell(row=row, column=2).value = 'CHC' 
    elif ws2_write.cell(row=row, column=3).value == 'Diamondbacks':
        ws2_write.cell(row=row, column=2).value = 'ARI'
    elif ws2_write.cell(row=row, column=3).value == 'Dodgers':
        ws2_write.cell(row=row, column=2).value = 'LAD'
    elif ws2_write.cell(row=row, column=3).value == 'Giants':
        ws2_write.cell(row=row, column=2).value = 'SF'
    elif ws2_write.cell(row=row, column=3).value == 'Indians':
        ws2_write.cell(row=row, column=2).value = 'CLE' 
    elif ws2_write.cell(row=row, column=3).value == 'Mariners':
        ws2_write.cell(row=row, column=2).value = 'SEA'
    elif ws2_write.cell(row=row, column=3).value == 'Marlins':
        ws2_write.cell(row=row, column=2).value = 'MIA'
    elif ws2_write.cell(row=row, column=3).value == 'Mets':
        ws2_write.cell(row=row, column=2).value = 'NYM' 
    elif ws2_write.cell(row=row, column=3).value == 'Nationals':
        ws2_write.cell(row=row, column=2).value = 'WSH'
    elif ws2_write.cell(row=row, column=3).value == 'Orioles':
        ws2_write.cell(row=row, column=2).value = 'BAL'
    elif ws2_write.cell(row=row, column=3).value == 'Padres':
        ws2_write.cell(row=row, column=2).value = 'SD' 
    elif ws2_write.cell(row=row, column=3).value == 'Phillies':
        ws2_write.cell(row=row, column=2).value = 'PHI'
    elif ws2_write.cell(row=row, column=3).value == 'Pirates':
        ws2_write.cell(row=row, column=2).value = 'PIT'
    elif ws2_write.cell(row=row, column=3).value == 'Rangers':
        ws2_write.cell(row=row, column=2).value = 'TEX' 
    elif ws2_write.cell(row=row, column=3).value == 'Rays':
        ws2_write.cell(row=row, column=2).value = 'TB'
    elif ws2_write.cell(row=row, column=3).value == 'Red Sox':
        ws2_write.cell(row=row, column=2).value = 'BOS'
    elif ws2_write.cell(row=row, column=3).value == 'Reds':
        ws2_write.cell(row=row, column=2).value = 'CIN' 
    elif ws2_write.cell(row=row, column=3).value == 'Rockies':
        ws2_write.cell(row=row, column=2).value = 'COL'
    elif ws2_write.cell(row=row, column=3).value == 'Royals':
        ws2_write.cell(row=row, column=2).value = 'KC'
    elif ws2_write.cell(row=row, column=3).value == 'Tigers':
        ws2_write.cell(row=row, column=2).value = 'DET' 
    elif ws2_write.cell(row=row, column=3).value == 'Twins':
        ws2_write.cell(row=row, column=2).value = 'MIN'
    elif ws2_write.cell(row=row, column=3).value == 'White Sox':
        ws2_write.cell(row=row, column=2).value = 'CHW'
    elif ws2_write.cell(row=row, column=3).value == 'Yankees':
        ws2_write.cell(row=row, column=2).value = 'NYY' 
    else:
        continue

for player in range(1,fantraxrows):
    namelist = []
    namelist = ws3.cell(row=player, column=1).value.split()
    fantraxmlbteam = ws3.cell(row=player, column=3).value
    fantraxplayername = ws3.cell(row=player, column=1).value
    for fangraphplayer in range(2,hitterrows):
        fgnamelist = []
        fgnamelist = ws_write.cell(row=fangraphplayer, column=1).value.split() #checks hitters first
        fgmlbteam = ws_write.cell(row=fangraphplayer, column=2).value
        fgplayername = ws_write.cell(row=fangraphplayer, column=1).value
        if namelist[0][:3].lower() == fgnamelist[0][:3].lower() and namelist[1][:4].lower() == fgnamelist[1][:4].lower() and fantraxmlbteam == fgmlbteam:
            ws3.cell(row=player, column=7).value = ws_write.cell(row=fangraphplayer, column=1).value
        elif fgplayername[:3].lower() == fantraxplayername[:3].lower() and fgplayername[-3:].lower() == fantraxplayername[-3:].lower() and fantraxmlbteam == fgmlbteam:
            ws3.cell(row=player, column=7).value = ws_write.cell(row=fangraphplayer, column=1).value
            print("RYU EXCEPTION")
        else:
            continue
        print(fantraxplayername[:3])
        print(fgplayername[-3:])

for player in range(1,fantraxrows):
    namelist = []
    namelist = ws3.cell(row=player, column=1).value.split()
    fantraxmlbteam = ws3.cell(row=player, column=3).value
    for fangraphplayer in range(2,pitcherrows):
        fgnamelist = []
        fgnamelist = ws2_write.cell(row=fangraphplayer, column=1).value.split() #checks pitchers second
        fgmlbteam = ws2_write.cell(row=fangraphplayer, column=2).value
        if namelist[0][:3].lower() == fgnamelist[0][:3].lower() and namelist[1][:4].lower() == fgnamelist[1][:4].lower() and fantraxmlbteam == fgmlbteam:
            ws3.cell(row=player, column=7).value = ws2_write.cell(row=fangraphplayer, column=1).value
        else:
            continue

wb_write.save('steamer.xlsx')
wb3.save('fantrax.xlsx')
