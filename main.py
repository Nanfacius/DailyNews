# -*-  coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
import os,time,re,json
from datetime import datetime
import pandas as pd
from time import sleep

class Firefox_driver():
    def __init__(self,URL='https://vbdata.cn/newsList',path=os.path.join(os.getcwd(),'DATA')):
        self.URL=URL
        self.path=path
        if not os.path.exists(path):
            os.mkdir(path)
        self.dict={}
        self.dataframe=pd.DataFrame(columns=['Date','Time','Tag','Title','Content'])
        self.driver=webdriver.Firefox()
        self.driver.implicitly_wait(20)
        
    def get_URL(self):
        self.driver.get(self.URL)
        checkboxes=self.driver.find_elements(By.CLASS_NAME,'ivu-checkbox-input')
        checkboxes[1].click()
        while True:
            more=self.driver.find_element(By.CLASS_NAME,'more')
            if more.text=='加载更多':
                break
            else:
                sleep(0.5)
        button=more.find_element(By.TAG_NAME,'button')
        button.click()
        sleep(3)
        
    def get_data(self):
        dates=self.driver.find_elements(By.CSS_SELECTOR,'div.info>p')
        newslist=self.driver.find_element(By.CLASS_NAME,'news_list')
        items=newslist.find_elements(By.CLASS_NAME,'item')
        return items,dates
    
    def process_data(self,item,dates):
        y_dates=[date.location['y'] for date in dates]
        i=0
        for item in items:
            if len(y_dates)==2 and item.location['y']>y_dates[1]:
                i=1
            time=item.find_element(By.CLASS_NAME,'time').text
            a1=item.find_element(By.CLASS_NAME,'a1')
            tag=a1.find_element(By.CLASS_NAME,'tag').text
            title=a1.text
            a2=item.find_element(By.CLASS_NAME,'a2')
            content=a2.text
            self.dataframe=self.dataframe.append({'Date':re.findall('[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日',dates[i].text)[0],'Time':time,'Tag':tag,'Title':title,'Content':content},ignore_index=True)
            self.dataframe.columns=['Date','Time','Tag','Title','Content']

    def save_data(self):
        self.dataframe.to_excel(os.path.join(self.path,'每日新闻'+time.strftime("%Y%m%d_%H%M", time.localtime())+'.xlsx'),index=False)
    
    def concat_files(self,date1=None,date2=time.strftime("%Y%m%d", time.localtime()),filepath=None):#合并文件，日期从date1到date2，格式为‘YYYYmmDD’,date1若赋值'today'则合并当天新闻
        if not date1:
            date1='20000101'
        if date1=='today':
            date1=time.strftime("%Y%m%d", time.localtime())
        if not filepath:
            filepath=os.path.join(self.path,'每日新闻合并.xlsx')
        dfs=[self.dataframe,]
        for file in os.listdir(self.path):
            if file[-5:]=='.xlsx':
                date=re.findall('[0-9]{8}',file)
                if date:
                    if int(date[0])>=int(date1) and int(date[0])<=int(date2):
                        dfs.append(pd.read_excel(os.path.join(self.path,file)))
        df=pd.concat(dfs,ignore_index=True).drop_duplicates().sort_values(by=['Date','Time'],ascending=False)
        df.to_excel(filepath,index=False)
    def quit(self):
        self.driver.quit()
    
if __name__ == '__main__':
    driver=Firefox_driver()
    driver.get_URL()
    items,dates=driver.get_data()
    driver.process_data(items,dates)
    driver.quit()
    driver.save_data()
    driver.concat_files('today')
