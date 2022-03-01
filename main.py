# -*-  coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
import os,time,re,json
from datetime import datetime
import pandas as pd
from time import sleep
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

class Chrome_driver():
    def __init__(self,URL='https://vbdata.cn/newsList',path=os.path.join(os.getcwd(),'DATA'),driver_path=None):
        print('Starting driver......')
        self.URL=URL
        self.path=path
        if not os.path.exists(path):
            os.mkdir(path)
        self.dict={}
        self.dataframe=pd.DataFrame(columns=['Date','Time','Tag','Title','Content'])
        option = webdriver.ChromeOptions()
        option.add_argument('headless')
        if driver_path:
            self.driver=webdriver.Chrome(driver_path,options=option)
        else:
            self.driver=webdriver.Chrome('/usr/bin/chromedriver',chrome_options=option)
        self.driver.implicitly_wait(20)
        print('Driver started!')
        
    def get_URL(self):
        print('Connecting......')
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
        print('Connected!')
        
    def get_data(self):
        print('Getting data.......')
        dates=self.driver.find_elements(By.CSS_SELECTOR,'div.info>p')
        newslist=self.driver.find_element(By.CLASS_NAME,'news_list')
        items=newslist.find_elements(By.CLASS_NAME,'item')
        return items,dates
    
    def process_data(self,item,dates):
        print('Processing data......')
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
        print('Saving data......')
        self.dataframe.to_excel(os.path.join(self.path,'每日新闻'+time.strftime("%Y%m%d_%H%M", time.localtime())+'.xlsx'),index=False)
        print('Data saved!')
        
    def concat_files(self,date1=None,date2=time.strftime("%Y%m%d", time.localtime()),filepath=None):#合并文件，日期从date1到date2，格式为‘YYYYmmDD’,date1若赋值'today'则合并当天新闻
        print('Concatenating files......')
        self.filepath=filepath
        if not date1:
            date1='20000101'
        if date1=='today':
            date1=time.strftime("%Y%m%d", time.localtime())
        if not filepath:
            filepath=os.path.join(self.path,'每日新闻合并'+time.strftime("%Y%m%d", time.localtime())+'.xlsx')
        dfs=[self.dataframe,]
        for file in os.listdir(self.path):
            if file[-5:]=='.xlsx':
                date=re.findall('[0-9]{8}',file)
                if date:
                    if int(date[0])>=int(date1) and int(date[0])<=int(date2):
                        dfs.append(pd.read_excel(os.path.join(self.path,file)))
        df=pd.concat(dfs,ignore_index=True).drop_duplicates().sort_values(by=['Date','Time'],ascending=False)
        df.to_excel(filepath,index=False)
        print('Files contenated!')
        
    def quit(self):
        self.driver.quit()
        print('Driver quited!')
        
    def send_mail(self):
        #使用第三方邮件服务
        mail_host="smtp.qq.com"  #设置服务器
        mail_user="1696239338@qq.com"    #用户名
        mail_pass="oimlkxiajuigcfjf"   #口令 

        sender = '1696239338@qq.com'
        receivers = ['1696239338@qq.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱

        #创建一个带附件的实例
        message = MIMEMultipart()
        message['From'] = Header("爬虫", 'utf-8')
        message['To'] =  Header("尊敬的用户", 'utf-8')
        subject = '今日医药新闻'
        message['Subject'] = Header(subject, 'utf-8')

        #邮件正文内容 
        # 构造附件1
        att1 = MIMEText(open(self.filepath, 'rb').read(), 'base64', 'utf-8')
        att1["Content-Type"] = 'application/octet-stream'
        # 这里的filename可以任意写，写什么名字，邮件中显示什么名字（不要使用中文）
        att1["Content-Disposition"] = 'attachment; filename="DailyNews20220301.xlsx"'
        message.attach(att1)

        try:
            smtpObj = smtplib.SMTP() 
            smtpObj.connect(mail_host, 25)    # 25 为 SMTP 端口号
            smtpObj.login(mail_user,mail_pass)  
            smtpObj.sendmail(sender, receivers, message.as_string())
            print("邮件发送成功")
        except smtplib.SMTPException:
            print("Error: 无法发送邮件")
            
print(os.getcwd())
# if __name__ == '__main__':
driver=Chrome_driver() #传递driver路径，如果driver保存在Python.exe相同目录下则可以不传递参数
driver.get_URL()
items,dates=driver.get_data()
driver.process_data(items,dates)
driver.quit()
driver.save_data()
filepath='每日新闻合并'+time.strftime("%Y%m%d", time.localtime())+'.xlsx'
driver.concat_files('today',filepath=filepath)
driver.send_mail()
