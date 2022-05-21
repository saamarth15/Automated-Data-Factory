#import libraries
import datetime
import time
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotVisibleException
from selenium import webdriver
from bs4 import BeautifulSoup
import pyodbc
from selenium.webdriver.common.action_chains import ActionChains
#connectiong to database
server = 'abdpl-l2.database.windows.net'
database = 'ABDPL'
username = 'abdadmin'
password = 'abdpl@ho13'
driver= '{ODBC Driver 13 for SQL Server}'
db = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+ password)
cur = db.cursor()
yesterday = datetime.datetime.today() - datetime.timedelta(days=1)
yesterday_day = yesterday.strftime("%d")
yesterday_month = yesterday.strftime("%B")
yesterday_year = yesterday.strftime("%Y")
yesterday1 = yesterday.strftime("%d/%m/%Y")
#delete the old values if you rerun the program for same day
cur.execute("SELECT * FROM JHARKHAND WHERE DATE='"+yesterday1+"'")
s = cur.fetchall()
if s:
    cur.execute("DELETE FROM JHARKHAND WHERE DATE = '"+yesterday1+"'")
    db.commit()
#data extraction procedure using selenium webdriver and storing to mysql
url = "https://jharkhandutpad.nic.in/jsbcl/"
browser = webdriver.Chrome(executable_path='C:\Users\dloadmin\Desktop\chromedriver.exe')
browser.get(url)
username = browser.find_element_by_name("txtUserName")
password1 = browser.find_element_by_name("txtPassword")
firstelement = browser.find_element_by_id("txtNum1")
secondelement = browser.find_element_by_id("txtNum2")
result = browser.find_element_by_name("txtResult")
answer = int(firstelement.text) + int(secondelement.text)
login = browser.find_element_by_id("btnLogin")
username.send_keys("sup_all")
password1.send_keys("ABD@#$%")
result.send_keys(answer)
login.click()
browser.implicitly_wait(10)
hover_element = browser.find_element_by_xpath("(//a[contains(text(),'Sales Report |')])")
hover = ActionChains(browser).move_to_element(hover_element)
hover.perform()
browser.find_element_by_xpath("(//a[contains(text(),'NewSaleReport')])").click()
browser.implicitly_wait(10)
browser.get("https://jharkhandutpad.nic.in/jsbcl/Reports/NewSaleReport.aspx")
browser.find_element_by_xpath('//input[@id="ImgBtnCalf"]').click()
browser.find_element_by_xpath('//*[@id="CE_title"]').click()
time.sleep(2)
browser.find_element_by_xpath('//*[@id="CE_title"]').click()
time.sleep(2)
browser.find_element_by_xpath("//*[@id='CE_yearsBody']/tr/td/div[text()='"+yesterday_year+"']").click()
time.sleep(2)
browser.find_element_by_xpath("//*[@id='CE_monthsBody']/tr/td/div[@title='"+yesterday_month+", "+yesterday_year+"']").click()
time.sleep(2)
browser.find_element_by_xpath("//*[@id='CE_daysBody']/tr/td/div[contains(@title,'"+yesterday_month+" "+yesterday_day+"')]").click()
browser.find_element_by_xpath('//input[@id="ImgBtnCal"]').click()
browser.find_element_by_xpath('//*[@id="CE2_title"]').click()
time.sleep(2)
browser.find_element_by_xpath('//*[@id="CE2_title"]').click()
time.sleep(2)
browser.find_element_by_xpath("//*[@id='CE2_yearsBody']/tr/td/div[text()='"+yesterday_year+"']").click()
time.sleep(2)
browser.find_element_by_xpath("//*[@id='CE2_monthsBody']/tr/td/div[@title='"+yesterday_month+", "+yesterday_year+"']").click()
time.sleep(2)
browser.find_element_by_xpath("//*[@id='CE2_daysBody']/tr/td/div[contains(@title,'"+yesterday_month+" "+yesterday_day+"')]").click()
browser.find_element_by_id("RadioButtonView").click()
p = []
z = browser.find_elements_by_xpath("//select[@name='ddlGodown']/option[@value]")
for i in z[1:]:
    p.append(str(i.text))
for k in  p:
        browser.find_element_by_xpath("//select[@name='ddlGodown']/option[text()= '"+k+"']").click()
        browser.find_element_by_name("btnShow").click()
        time.sleep(10)
        try:
            browser.find_element_by_id("ReportViewer1_ctl05_ctl00_Last_ctl00_ctl00").click()
        except ElementNotVisibleException:
            pass
        time.sleep(2)
        n=browser.find_element_by_id("ReportViewer1_ctl05_ctl00_TotalPages")
        try:
            browser.find_element_by_id("ReportViewer1_ctl05_ctl00_First_ctl00_ctl00").click()
        except ElementNotVisibleException:
            pass
        time.sleep(2)
        button = browser.find_element_by_id("ReportViewer1_ctl05_ctl00_Next_ctl00_ctl00")
        for i in range(0,int(n.text)):
            datalist = []
            page = browser.page_source
            soup = BeautifulSoup(page,"lxml")
            table = soup.find_all('table')[3]
            table_row = table.find_all('tr')[25]
            td = table_row.find_all('td')[2]
            box = td.find_all('div')
            for i in range(22,(len(box)-10),9):
                data1 = []
                for j in range(0,9):
                    data1.append(((box[i+j].text).encode('ascii','ignore').decode('ascii')).replace(',',''))
                datalist.append(data1)
            #data insertion
            for i in datalist:
                cur.execute("insert into JHARKHAND(GODOWN_NAME,DATE,LICENCE_NAME,DISTRICT,BRAND,LABEL_NAME,PACK_SIZE,CASE1,BTL) values(?,?,?,?,?,?,?,?,?)",  (i[0],i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8]))
            db.commit()
            try:
                button.click()
            except ElementNotVisibleException:
                pass
            time.sleep(5)
#data cleaning. Removing garbage data
cur.execute("DELETE FROM JHARKHAND WHERE DATE = 'NO Record Found'")
cur.execute("DELETE FROM JHARKHAND WHERE ISNUMERIC(GODOWN_NAME) = 1")
db.commit()
db.close()
browser.close()
print "jharkhand task completed "+yesterday1
