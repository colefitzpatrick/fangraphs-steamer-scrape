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

wb_write = openpyxl.load_workbook('steamer.xlsx')
ws_write = wb_write["hitters"]
ws2_write = wb_write["pitchers"]
wb3 = openpyxl.load_workbook('fantrax.xlsx')
ws3 = wb3["Sheet1"]
wb4 = openpyxl.load_workbook('teamacronyms.xlsx')
ws4 = wb4["Sheet1"]

#### scrapes either the pitcher or hitter projections from fangraphs, specificy the worksheet to write to and the number of pages to scrape ####

def steamerscrape(link, pages, worksheet):
    driver = webdriver.Firefox()
    driver.implicitly_wait(8)
    driver.get(link)
    writerow = 2
    for page in range(1,pages):         
        rightbutton = driver.find_element_by_xpath("//button[@class='t-button rgActionButton rgPageNext']")
        soup_level1=BeautifulSoup(driver.page_source, 'lxml')
        
        for tr in soup_level1.findAll("tr", {"class": "rgRow"}):   #the table is structured where each row has an alternating tr-class, the first loop pulls the odds, second loop pulls the even rows
            tds = tr.findAll("td", {"class": "grid_line_regular"})
            writecol=1
            for td in tds:
                if writecol < 4:
                    worksheet.cell(row=writerow, column=writecol).value = td.text     #writes the name and team values into spreadsheet
                    writecol += 1
                else:
                    worksheet.cell(row=writerow, column=writecol).value = float(td.text)     #writes the statistical values into spreadsheet
                    writecol += 1
            writerow += 1
            
        for tr in soup_level1("tr", {"class": "rgAltRow"}):
            tds = tr.findAll("td", {"class": "grid_line_regular"})
            writecol=1
            for td in tds:
                if writecol < 4:
                    worksheet.cell(row=writerow, column=writecol).value = td.text     #writes the name and team values into spreadsheet
                    writecol += 1
                else:
                    worksheet.cell(row=writerow, column=writecol).value = float(td.text)     #writes the statistical values into spreadsheet
                    writecol += 1
            writerow += 1

        driver.execute_script("arguments[0].click();", rightbutton)       #clicks the right arrow to advance to the next page
        time.sleep(3)
    wb_write.save('steamer.xlsx')
    print("Scrape complete for: " + str(link))


#### converts each team name to the 2 or 3 letter acronym that matches Fantrax convention ####

def teamacronyms(worksheet):
    rows = worksheet.max_row
    for row in range(1,rows+1):
        for team in range(1,31):
            if worksheet.cell(row=row, column=3).value == ws4.cell(row=team, column=1).value:
                worksheet.cell(row=row, column=2).value = ws4.cell(row=team, column=2).value
            else:
                continue
    wb_write.save('steamer.xlsx')
    print("Acronyms assigned for: " + str(worksheet))


#### calculates the points per game for hitters ####
 
def hitterppgproj():
    hitterrows = ws_write.max_row
    for row in range(2,hitterrows+1):
        mobg = (ws_write.cell(row=row, column = 18).value * ws_write.cell(row=row, column = 4).value * 1.9651 - 36.2)         #MOBG
        hits = ws_write.cell(row=row, column = 7).value
        doubles = ws_write.cell(row=row, column = 8).value
        triples = ws_write.cell(row=row, column = 9).value
        hr = ws_write.cell(row=row, column = 10).value
        singles = hits - doubles - triples - hr
        run = ws_write.cell(row=row, column = 11).value
        rbi = ws_write.cell(row=row, column = 12).value
        bbhbp = ws_write.cell(row=row, column = 13).value + ws_write.cell(row=row, column = 15).value
        so = ws_write.cell(row=row, column = 14).value
        sb = ws_write.cell(row=row, column = 16).value
        cs = ws_write.cell(row=row, column = 17).value
        fpts = singles + doubles*2 + triples*3 + hr*4 + run*2 + rbi + bbhbp - so + sb*2 - cs*2 +mobg*2
        ppg = fpts / ws_write.cell(row=row, column = 4).value
        ws_write.cell(row=row, column = 26).value = fpts                      #Fantasy Points
        ws_write.cell(row=row, column = 27).value = ppg                       #PPG
    wb_write.save('steamer.xlsx')
    print("Hitter calculations complete")

#### calculates the points per game for pitchers ####

def pitcherppgproj():
    pitcherrows = ws2_write.max_row
    for row in range(2,pitcherrows+1):
        gs = ws2_write.cell(row=row, column=7).value
        pitcher_wins = ws2_write.cell(row=row, column=4).value
        pitcher_losses = ws2_write.cell(row=row, column=5).value
        ip = ws2_write.cell(row=row, column=10).value
        hitsallowed = ws2_write.cell(row=row, column=11).value
        er = ws2_write.cell(row=row, column=12).value
        pitcher_so = ws2_write.cell(row=row, column=14).value
        pitcher_bb = ws2_write.cell(row=row, column=15).value
        pitcher_fpts = (pitcher_wins*3 - pitcher_losses*3 + ip*3 - hitsallowed*0.5 - er*2 + pitcher_so - pitcher_bb) * 1.09    #1.09 coefficient accounts for the things that steamer does not project (eg. SHO, CG, NH)
        pitcher_ppg = pitcher_fpts / ws2_write.cell(row=row, column=8).value
        if gs > 4:
            if pitcher_ppg > 8.5:
                ws2_write.cell(row=row, column=23).value = pitcher_ppg
            else:
                ws2_write.cell(row=row, column=23).value = 8.5
        else:
            ws2_write.cell(row=row, column=23).value = 0
    wb_write.save('steamer.xlsx')
    print("Pitcher calculations complete")
        

####  links points per game projection from fangraphs steamer projection into the fantrax ownership sheet ####

def ppglinker():
    hitterrows = ws_write.max_row
    pitcherrows = ws2_write.max_row
    fantraxrows = ws3.max_row
    for player in range(1,fantraxrows):
        namelist = []
        namelist = ws3.cell(row=player, column=1).value.split()
        fantraxmlbteam = ws3.cell(row=player, column=3).value
        fantraxplayername = ws3.cell(row=player, column=1).value
        fantraxposition = ws3.cell(row=player, column=2).value
        if fantraxposition in ['SP','RP','SP,RP']:
            for fangraphplayer in range(2,pitcherrows): #links pitchers first
                fgnamelist = []
                fgnamelist = ws2_write.cell(row=fangraphplayer, column=1).value.split() 
                fgmlbteam = ws2_write.cell(row=fangraphplayer, column=2).value
                fgplayername = ws2_write.cell(row=fangraphplayer, column=1).value
                if namelist[0][:3].lower() == fgnamelist[0][:3].lower() and namelist[1][:4].lower() == fgnamelist[1][:4].lower() and fantraxmlbteam == fgmlbteam:
                    ws3.cell(row=player, column=7).value = ws2_write.cell(row=fangraphplayer, column=23).value
                elif fgplayername[:3].lower() == fantraxplayername[:3].lower() and fgplayername[-3:].lower() == fantraxplayername[-3:].lower() and fantraxmlbteam == fgmlbteam:
                    ws3.cell(row=player, column=7).value = ws2_write.cell(row=fangraphplayer, column=23).value
                else:
                    continue
        else:
            for fangraphplayer in range(2,hitterrows): #links hitters seconds
                fgnamelist = []
                fgnamelist = ws_write.cell(row=fangraphplayer, column=1).value.split() 
                fgmlbteam = ws_write.cell(row=fangraphplayer, column=2).value
                fgplayername = ws_write.cell(row=fangraphplayer, column=1).value
                if namelist[0][:3].lower() == fgnamelist[0][:3].lower() and namelist[1][:4].lower() == fgnamelist[1][:4].lower() and fantraxmlbteam == fgmlbteam:
                    ws3.cell(row=player, column=7).value = ws_write.cell(row=fangraphplayer, column=27).value
                elif fgplayername[:3].lower() == fantraxplayername[:3].lower() and fgplayername[-3:].lower() == fantraxplayername[-3:].lower() and fantraxmlbteam == fgmlbteam:
                    ws3.cell(row=player, column=7).value = ws_write.cell(row=fangraphplayer, column=27).value
                else: 
                    continue
    wb3.save('fantrax.xlsx')
    print("Linker complete")

#### section where you call the various functions depending on what you need ####

steamerscrape(url, 25, ws_write)               #scrapes hitter projections
steamerscrape(url2, 25, ws2_write)              #scrapers pitcher projections

teamacronyms(ws_write)                          #assign team acronyms to hitter sheet
teamacronyms(ws2_write)                         #assign team acronyms to pitcher sheet

hitterppgproj()                                #calculate hitter ppg
pitcherppgproj()                               #calculate pitcher ppg

ppglinker()                                    #link steamer projections into the fantrax sheet
        

