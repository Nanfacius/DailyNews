# -*-  coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
import os,time,re
import pandas as pd
from time import sleep
import schedule

# 如果不需要发邮件，可以把下面几行注释掉
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

class Chrome_driver():
    def __init__(self,URL='https://vbdata.cn/newsList',path=os.path.join(os.getcwd(),'DATA'),driver_path=None):
        print('Starting driver......')
        self.URL=URL
        self.path=path
        self.raw_path=os.path.join(self.path,'raw_file')
        self.concat_path=os.path.join(self.path,'concat_file')
        
        #创建目录
        if not os.path.exists(self.path):
            os.mkdir(self.path)
        if not os.path.exists(self.raw_path):
            os.mkdir(self.raw_path)
        if not os.path.exists(self.concat_path):
            os.mkdir(self.concat_path)
            
        self.dict={}
        self.dataframe=pd.DataFrame(columns=['Date','Time','Tag','Title','Content'])
        
        option = webdriver.ChromeOptions()
        #设置Webdriver静默模式，若要进行debug请注释掉下面这一行
        option.add_argument('headless')
        
        #创建Webdriver对象
        if driver_path:
            self.driver=webdriver.Chrome(driver_path,options=option)
        else:
            self.driver=webdriver.Chrome(options=option)
            
        #设置隐含等待时间
        self.driver.implicitly_wait(30)
        
        print('Driver started!')
        
    def get_URL(self):
        print('Connecting......')
        
        #打开链接
        self.driver.get(self.URL)
        
        #等待网页内容出现后点击“只看推荐的”选框
        checkboxes=self.driver.find_elements(By.CLASS_NAME,'ivu-checkbox-input')
        sleep(15)
        checkboxes=self.driver.find_elements(By.CLASS_NAME,'ivu-checkbox-input')
        checkboxes[1].click()
        
        #点击“加载更多”
        while True:
            more=self.driver.find_element(By.CLASS_NAME,'more')
            if more.text=='加载更多':
                break
            else:
                sleep(0.5)
        button=more.find_element(By.TAG_NAME,'button')
        button.click()
        sleep(5)
        
        print('Connected!')
        
    def get_data(self):
        print('Getting data.......')

        #获取新闻
        newslist=self.driver.find_element(By.CLASS_NAME,'news_list')
        items=newslist.find_elements(By.CLASS_NAME,'item')
        
        #获取日期 可能存在2个日期（今天和昨天）
        dates=[]
        date1=self.driver.find_element(By.CSS_SELECTOR,'div.info>p')
        try:
            date2=newslist.find_element(By.CSS_SELECTOR,'div.d-time>span')
            dates=[date1,date2]
        except:
            dates=[date1]
            
        return items,dates
    
    def process_data(self,items,dates):
        print('Processing data......')
        
        #得到日期在网页上的位置
        y_dates=[date.location['y'] for date in dates]
        
        #遍历新闻
        for item in items:
            
            #确定新闻日期
            if len(y_dates)==2 and item.location['y']>y_dates[1]:
                date=re.findall('[0-9]{4}年',dates[0].text)[0]+dates[1].text
            else:
                date=re.findall('[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日',dates[0].text)[0]
            
            #获取新闻时间、标签、标题和正文内容
            time=item.find_element(By.CLASS_NAME,'time').text
            a1=item.find_element(By.CLASS_NAME,'a1')
            tag=a1.find_element(By.CLASS_NAME,'tag').text
            title=a1.text
            a2=item.find_element(By.CLASS_NAME,'a2')
            content=a2.text
            
            #保存为pandas DataFrame
            self.dataframe=self.dataframe.append({'Date':date,'Time':time,'Tag':tag,'Title':title,'Content':content},ignore_index=True)
            self.dataframe.columns=['Date','Time','Tag','Title','Content']
    
    #保存文件
    def save_data(self):
        print('Saving data......')
        self.dataframe.to_excel(os.path.join(self.raw_path,'每日新闻'+time.strftime("%Y%m%d_%H%M", time.localtime())+'.xlsx'),index=False)
        print('Data saved!')
        
    def save_data_txt(self):
        print('Saving data txt......')
        self.dataframe.to_csv(os.path.join(self.raw_path,'每日新闻'+time.strftime("%Y%m%d_%H%M", time.localtime())+'.txt'),index=False)
        print('Data saved!')
        
    #合并从date1到date2已保存的的全部新闻文件
    def concat_files(self,date1=None,date2=time.strftime("%Y%m%d", time.localtime()),filepath=None):#合并文件，日期从date1到date2，格式为‘YYYYmmDD’,date1若赋值'today'则合并当天新闻
        print('Concatenating files......')
        self.filepath=filepath
        
        #判断日期
        if not date1:
            date1='20000101'
        if date1=='today':
            date1=time.strftime("%Y%m%d", time.localtime())
        
        #若未传递filepath则采用默认路径
        if not filepath:
            filepath=os.path.join(self.concat_path,'每日新闻合并'+time.strftime("%Y%m%d", time.localtime())+'.xlsx')
        
        #读取文件
        dfs=[self.dataframe,]
        for file in os.listdir(self.raw_path):
            if file[-5:]=='.xlsx':
                date=re.findall('[0-9]{8}',file)
                if date:
                    if int(date[0])>=int(date1) and int(date[0])<=int(date2):
                        dfs.append(pd.read_excel(os.path.join(self.raw_path,file)))
        
        #合并文件、删除重复并按日期和时间排序
        df=pd.concat(dfs,ignore_index=True).drop_duplicates().sort_values(by=['Date','Time'],ascending=False)
        
        #保存
        df.to_excel(filepath,index=False)
        
        print('Files contenated!')
                
    def create_report(self):
        print('Creating roport......')
        self.report_path=os.path.join(self.concat_path,'report'+time.strftime("%Y%m%d", time.localtime())+'.txt')
        df=pd.read_excel(self.filepath)
        text='[太阳]早上好~医疗行业“健康早餐”\n来啦！\n'+time.strftime("%Y/%m/%d", time.localtime())+'\n┈┈┈┈┈┈┈┈┈┈┈┈┈┈\n'
        for i in range(len(df)):
            text+=str(i+1)+df.loc[i,'Title']+'\n'
        text+='┈┈┈┈┈┈┈┈┈┈┈┈┈┈\nECV Digital Healthcare致力于推动数字医疗行业持续发展。ECV医疗社群欢迎分享干货与实践交流，更有行业报告及一手资源。入群或更多行业需求，请微信联系：joeys_1116'
        with open(self.report_path,'w') as f:
            f.write(text)
        print('Report created!')
        
    def quit(self):
        self.driver.quit()
        print('Driver quited!')
        
    def send_mail(self):
        print('Sending mail......')
        
        #从环境变量获取邮箱服务器、账号和密码
        mail_host=os.getenv("EMAIL_HOST")
        mail_user=os.getenv("EMAIL_ACCOUNT")
        mail_pass=os.getenv("EMAIL_PASSWORD") 
        sender=mail_user
        receivers=[mail_user]
        
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
        att1["Content-Disposition"] = 'attachment; filename="DailyNews'+time.strftime("%Y%m%d", time.localtime())+'.xlsx"'
        message.attach(att1)
        
        att2 = MIMEText(open(self.report_path, 'rb').read(), 'base64', 'utf-8')
        att2["Content-Type"] = 'application/octet-stream'
        # 这里的filename可以任意写，写什么名字，邮件中显示什么名字（不要使用中文）
        att2["Content-Disposition"] = 'attachment; filename="NewsReport'+time.strftime("%Y%m%d", time.localtime())+'.txt"'
        message.attach(att2)

        #发送邮件
        try:
            smtpObj = smtplib.SMTP() 
            smtpObj.connect(mail_host, 25)    # 25 为 SMTP 端口号
            smtpObj.login(mail_user,mail_pass)  
            smtpObj.sendmail(sender, receivers, message.as_string())
            print("Mail Sended")
        except smtplib.SMTPException:
            print("Failed")
            
def main():
    os.environ['TZ'] = 'Asia/Shanghai'
    driver=Chrome_driver(driver_path='/usr/bin/chromedriver') #传递driver路径，如果driver保存在Python.exe相同目录下则可以不传递参数
#     driver=Chrome_driver() #传递driver路径，如果driver保存在Python.exe相同目录下则可以不传递参数
    driver.get_URL()
    items,dates=driver.get_data()
    driver.process_data(items,dates)
    driver.quit()
    driver.save_data()
    driver.save_data_txt()
    filepath='DATA/concat_file/每日新闻合并'+time.strftime("%Y%m%d", time.localtime())+'.xlsx'
    driver.concat_files('today',filepath=filepath)
    driver.create_report()
    return driver

def main_with_mail():# 运行main并发送邮件
    driver=main()
    driver.send_mail()


# # 定时运行
# if __name__ == '__main__':
#     main_with_mail()
#     schedule.every().day.at('07:00').do(main)
#     while True:
#         try:
#             schedule.run_pending() # 运行所有可运行的任务
#         except Exception as e:
#             print(e)
#             continue
#         sleep(60)
        
# 单次运行
if __name__ == '__main__':
#     main()
    main_with_mail()



    
    
