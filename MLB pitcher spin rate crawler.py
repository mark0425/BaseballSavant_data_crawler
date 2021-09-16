#!/usr/bin/env python
# coding: utf-8

import pip

def import_or_install(package):
    try:
        __import__(package)
    except ImportError:
        pip.main(['install', package]) 
import_or_install("selenium")
import_or_install("pandas")
import_or_install("unidecode")
import_or_install("lxml")
import os
import sys
import time
import random
from bs4 import BeautifulSoup
from selenium import webdriver  
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from collections import OrderedDict
import pandas as pd
from selenium.webdriver.support.ui import Select
import numpy as np
import platform
import unidecode
import datetime
def imput_day(on):
    day=0; st = ["開始", "結束"]; st2 = ["(開季時間是3/28)", "(季末時間是9/30)"];
    while day<1 or day >31:
        print("請輸入搜尋"+str(st[on])+"日期"+str(st2[on]))
        try:
            day = int(input())
        except:
            print("不是正確的日期")
    return day
def imput_month(on):
    mon=0; st = ["開始", "結束"]; st2 = ["(開季時間是3/28)", "(季末時間是9/30)"];
    while mon<1 or mon >12:
        print("請輸入搜尋"+str(st[on])+"月份"+str(st2[on]))
        try:
            mon = int(input())
        except:
            print("不是正確的月份")
    return mon
startMonth = imput_month(0)
startDay = imput_day(0)
terMonth = imput_month(1)
terDate = imput_day(1) + 1
# baseball savant.com crawler #

# 參數定義
os_base = platform.system()
if os_base == 'Darwin':
    driverPath = "/chromedriver"
    keyControl = Keys.COMMAND
else:
    driverPath = "chromedriver_win32"
    keyControl = Keys.CONTROL
path_Chrome = os.getcwd()+"\\selenium_driver\\chromedriver\\chromedriver.exe"
# 使用 Google Chrome WebDriver 開啟本機 Google Chrome 瀏覽器
browser = webdriver.Chrome(driverPath)
# 如果網站在 10 秒內回應則繼續執行下一步，否則等待 10 秒
browser.implicitly_wait(10)
# 開啟網站
browser.get("https://baseballsavant.mlb.com/statcast_search")
# 瀏覽器視窗最大化
browser.maximize_window()
# 設定球種為四縫線速球
browser.find_element_by_xpath("//div[@class='glyphicon glyphicon-chevron-down']").click()
browser.find_element_by_id('lbl_PT_FF').click()
browser.find_element_by_xpath("//div[@class='glyphicon glyphicon-chevron-down']").click()
# 設定投出速球最小值
select = Select(browser.find_element_by_id('min_results'))
select.select_by_value('10')
# 設定group_by
select = Select(browser.find_element_by_id('group_by'))
select.select_by_visible_text('Player & Game Date')
# 設定以平均轉速排序
select = Select(browser.find_element_by_id('sort_col'))
select.select_by_visible_text('Avg. Spin Rate')
# 加入其他變數
browser.find_element_by_id('include_stats').click()
browser.find_element_by_xpath("//label[@for='chk_stats_pa']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_hits']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_hrs']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_so']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_k_percent']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_bb']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_babip']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_slg']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_obp']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_velocity']").click()
browser.find_element_by_xpath("//label[@for='chk_stats_launch_speed']").click()
browser.find_element_by_id('include_stats').click()

# 重複搜尋
# 定義變數
someThingWrong = False
playerNameList = []; dateList=[]; varArray = np.empty((0,14), int); spinArray = np.empty((0,1), int)
endMonth = 10
asg = ["2019-7-8", "2019-7-9", "2019-7-10"]
#年
y = 2019
# 設定賽季
seasonValue = "lbl_Sea_" + str(y)
browser.find_element_by_id('boxSea').click()
browser.find_element_by_id('lbl_Sea_2021').click()
browser.find_element_by_id(seasonValue).click()
browser.find_element_by_id('boxSea').click()
# 月
for m in range(startMonth, endMonth+1):
    # 31天或30天
    oddMonth = [3,5,7,8,10]
    if m in oddMonth:
        endDate = 31
    else:
        endDate = 30
    #日
    for d in range(startDay, endDate+1):
        # 設定日期
        dayValue = str(y) + "-" + str(m) + "-" + str(d)
        print(dayValue)
        # 終止時間
        if dayValue == '2019-'+str(terMonth)+'-'+str(terDate):
            someThingWrong = True
            break
        browser.find_element_by_id('game_date_gt').click()
        browser.find_element_by_id('game_date_gt').send_keys(keyControl, 'a')
        browser.find_element_by_id('game_date_gt').send_keys(Keys.BACKSPACE)
        browser.find_element_by_id('game_date_gt').send_keys(dayValue)
        browser.find_element_by_id('game_date_gt').send_keys(Keys.RETURN)
        browser.find_element_by_id('game_date_lt').click()
        browser.find_element_by_id('game_date_lt').send_keys(keyControl, 'a')
        browser.find_element_by_id('game_date_lt').send_keys(Keys.BACKSPACE)
        browser.find_element_by_id('game_date_lt').send_keys(dayValue)
        browser.find_element_by_id('game_date_lt').send_keys(Keys.RETURN)
        # 開始搜尋
        browser.find_element_by_xpath("//input[@class='btn btn-default btn-search-green']").click()
        # 定義參數
        pNList = [];dList = []
        page_result = []
        page_result.append(browser.page_source)
        # 以 BeautifulSoup 解析 HTML 程式碼
        soup = BeautifulSoup(page_result[0], "lxml")
        # 若沒資料且不在明星賽休賽則終止迴圈
        if not soup.find_all("table", attrs={"id": "search_results"}):
            if dayValue not in asg:
                someThingWrong = True
                break
        # 以 CSS 的 table 抓出各類資訊，就算他長得很醜很複雜，只要有一項是固定的就沒問題！
        count = 0
        for tag in soup.find_all("table", attrs={"id": "search_results"}):
            for Name in tag.find_all("td", attrs={"class": "player_name tr-data align-left table-static-column-two"}):
                name = Name.text
                name = "".join(line.strip() for line in name.split("\n"))
                pNList.append(name)
                count += 1

            count = 0
            for Date in tag.find_all("td", attrs={"class": "tr-data align-left"}):
                date = Date.text
                if count % 2 != 0:
                    dList.append(date)
                count += 1
            
            vArray = np.random.rand(len(pNList),14)
            rowCount = 0; colCount = 0
            for Variable in tag.find_all("td", attrs={"class": "tr-data align-right"}):
                if not Variable.text and (colCount == 7 or colCount == 13):
                    vArray[rowCount, colCount] = 0
                else:
                    vArray[rowCount, colCount] = float(Variable.text)        
                colCount += 1
                if colCount % 14 == 0:
                    rowCount += 1
                    colCount = 0
                    
            sArray = np.random.rand(len(pNList),1); count = 0
            for Spin in tag.find_all("td", attrs={"class": "tr-data align-right column-sort"}):
                    sArray[count, 0] = float(Spin.text)        
                    count += 1
        # 整合每日資料
        if dayValue not in asg:
            playerNameList.extend(pNList)
            dateList.extend(dList)
            varArray = np.append(varArray, vArray, axis=0)
            spinArray = np.append(spinArray, sArray, axis=0)
    # startDay改變
    startDay = 1
    if someThingWrong:
        break 




a = pd.DataFrame(playerNameList)
b = pd.DataFrame(varArray)
c = pd.DataFrame(dateList)
d = pd.DataFrame(spinArray)
df = pd.concat([a, c, b, d], axis=1)
df = df.set_axis(['Name', 'Date', 'Pitches', 'Total', 'Pitch%','PA',
                  'Hits', 'HR', 'SO', 'K%', 'BB', 'BABIP', 'SLG',
                  'OBP', 'Pitch(MPH)', 'EV(MPH)','Spin(MPH)'], axis=1, inplace=False)

print(df)
df.to_excel('output.xlsx')
