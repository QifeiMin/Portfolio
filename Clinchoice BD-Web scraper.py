#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jun 13 22:26:03 2023

@author:  QifeiMin
"""

# -*- codeing = utf-8 -*-
                         # time函数
import xlwt
import os.path
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
    # 等待条件

from xlutils.copy import copy

import xlrd

//total run time: 8 minutes

#红杉','君联','夏尔巴','礼来','启明','康桥','高瓴','元禾','杏泽','元生','同创伟业','泰康','银盛泰','高特佳','北极光'
list = ['红杉','君联','夏尔巴','礼来','启明','康桥','高瓴','元禾','杏泽','元生','同创伟业','泰康','银盛泰','高特佳','北极光'] //tags of the Clinchoice Investors to keep track of
go = 0 
book = xlwt.Workbook(encoding="utf-8",style_compression=0)
savepath = "/Users/yongmin/Desktop/爬虫资讯.xlsx" //saving to local excel file
sheet = book.add_sheet('1')
col = ("时间", "标题","链接")  //"published time","title","link", info of the investment news that keep pulled of the web
for i in range(0, len(col)):
    sheet.write(0, i, col[i])      # 创建工作表。sheet页名为dataTime   


# 获取全量数据的 selenium 驱动
def readHtml(baseurl, flag):
    print("—————————— Read Html ——————————")
    # 打开浏览器
    
    service = Service()
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service, options=options)
    options.add_argument('user-agent="Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36"')
    driver.get(baseurl)
    wait = WebDriverWait(driver, 10)
    driver.get('https://www.pedaily.cn')
    input = driver.find_element(By.ID, "top_searchkey")
    input.send_keys(i)
    input.send_keys(Keys.ENTER)
    
    wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div/div[2]/div/div[1]/div[1]/ul/li[1]/div[1]/a/img")))
   
    return driver
 
# 将 selenium 驱动转 bs 形式的 html 页面
def getHtml(driver):
    print("—————————— Get Html ——————————")
    # 获取完整渲染的网页源代码
    pageSource = driver.page_source
    soup = BeautifulSoup(pageSource, 'html.parser')
    soup.prettify()
    return soup
 
# 从 html 页面爬取数据
def getData(soup):
    print("—————————— Get Data ——————————")
    # 1. 时间
    
    # 2. 数据，查找符合要求的字符串                        # 排行
    
    item = soup.find_all('div', class_='img' )  
    
   
   
    timelist = []
    linklist=[]
    titlelist=[]
    t = soup.find_all('div',class_='info')  
    for j in t: 
        T = j.find('span',{'class':'date'}).get_text()
        timelist.append(T)
        num = len(timelist)
        global go 
        go = 0
        go = go + num 
        
              # 用来存储爬取的网页信息
   
    for i in item:
        
        
        link = i. find('a')['href']
        linklist.append(link)
        
     
    for i in item: 
        
        title = i.find('img')['alt']
        titlelist.append(title)
        
    
    print(timelist)   
    print(linklist)    
    print(titlelist)
    
    try:   //save the info to local excel
        workbook = xlrd.open_workbook(savepath)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
        new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
   
   
        for i in range(0,len(timelist)):
            new_worksheet.write(i+1+rows_old,0,timelist[i])
        for j in range(0, len(linklist)):
            new_worksheet.write(j+1+rows_old,2,linklist[j])
        for j in range(0, len(titlelist)):
            new_worksheet.write(j+1+rows_old,1,titlelist[j])
            new_workbook.save(savepath) 
        print('—————————— 写入成功 ——————————') //"saving to local excel succeeded"
    except Exception as e:
        print('—————————— 写入失败 ——————————',e) //"saving to local excel failed"

    
    
    
    
    return titlelist
            
   

 
if __name__ == "__main__":
    print("—————————— 开始执行 ——————————")  //"start of scraper "
    
    for i in list:
        # 1. 读取url
        html = readHtml("https://www.pedaily.cn", True)
        # 2. selenium转BeautifulSoup
        soup = getHtml(html)
        # 3. 处理Html数据
        dataList = getData(soup)
        
    
       
        print("—————————— 爬取完毕 ——————————") //"end of scraping"

        
