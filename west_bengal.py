#import libraries
import datetime
import time
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
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
yesterday1 = yesterday.strftime("%d/%m/%Y")
yesterday1_day = yesterday.strftime("%d")

#delete the old values if you re-run the program for same day
cur.execute("SELECT * FROM WB1 WHERE DATE='"+yesterday1+"'")
s = cur.fetchall()
if s:
    cur.execute("DELETE FROM WB1 WHERE DATE = '"+yesterday1+"'")
    db.commit()

#data extraction procedure using selenium webdriver and storing to mysql
url='https://excise.wb.gov.in/WBSBCL/Bevco/NIC/UserLogin/Login.aspx'
browser = webdriver.Chrome(executable_path='C:\Users\dloadmin\Desktop\chromedriver.exe')
browser.get(url)
username = browser.find_element_by_name("txt_username")
password = browser.find_element_by_name("txt_password")
captcha = browser.find_element_by_name("CodeNumberTextBox")
login_button = browser.find_element_by_name("ImageButton1")
username.send_keys("B/2018/016")
password.send_keys("Va@achDR")
x=raw_input("enter captcha\n")
captcha.send_keys(x)
login_button.click()
try:
    alert_obj = browser.switch_to.alert
    alert_obj.accept()
except NoAlertPresentException:
    pass
try:
    browser.find_element_by_id("ctl00_ContentPlaceHolder1_ImageButton1").click()
except NoSuchElementException:
    pass
browser.find_element_by_name("ctl00$img_bttn_BI").click()
browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_optTabs"]/label[5]').click()
time.sleep(5)
z=['Warehouse at Siliguri','Bankura','Sonari','Contai','warehouse at Dankuni','Durgapur','Cossipore','Panchla','Murshidabad','Usthi','Habra','Kasba','Krishnanagar','Maheshtala','Khardah','Kharagpur','Jamalpur','Jalpaiguri','Gangarampur','Warehouse at Malda','Suri','Purulia','Rampurhat']
for k in z:
    browser.find_element_by_xpath("//select[@name='ctl00$ContentPlaceHolder1$ddl_date']/option[text()='"+yesterday1+"']").click()
    browser.find_element_by_xpath("//select[@name='ctl00$ContentPlaceHolder1$ddl_Warehouse']/option[contains(text(),'"+k+"')]").click()
    browser.find_element_by_name("ctl00$ContentPlaceHolder1$btn_Show_grid").click()
    time.sleep(10)
    j=2
    while True:
        datalist = []    
        soup = BeautifulSoup(browser.page_source,'lxml')
        try:
            tables=soup.find_all('table')[17]
        except IndexError:
            break
        table_rows=tables.find_all('tr')
        for tr in table_rows[4:-1]:
            if len(tr)==10:
               td = tr.find_all('td')
               row = [(((i.text).encode('ascii','ignore').decode('ascii')).strip('\n')).replace(',','') for i in td]
               datalist.append(row)
        for i in datalist:
               i.append(u""+yesterday1)
        for i in datalist:
               z=i[3].split('-')
               i[3:4]=z
        for i in datalist:
            i.append(i[0][-13:-1])
            i[0]=i[0].replace(i[0][-14:],'')
        for i in datalist:
            if('750' in i[2]):
                m=float(i[3])+float(i[4])/12
            elif('500' in i[2]):
                m=float(i[3])+float(i[4])/18
            elif('1000' in i[2]):
                m=float(i[3])+float(i[4])/9
            elif('180' in i[2]):
                m=float(i[3])+float(i[4])/48
            elif('375' in i[2]):
                m=float(i[3])+float(i[4])/24
            else:
                m=float(i[3])
            i.append(str(m).decode('utf-8'))
        #data insertion
        for i in datalist:
                    cur.execute("insert into WB1(RETAILER,INVOICE,PRODUCT,CASE1,BOTTLES,LANDED_COST,BEVCO_MARGIN,TCS_AMOUNT,INVOICE_AMOUNT,DATE,code,tot_cases) values(?,?,?,?,?,?,?,?,?,?,?,?)",  (i[0],i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8],i[9],i[10],i[11]))
        db.commit()
        try:
            browser.find_element_by_xpath("//*[@id='ctl00_ContentPlaceHolder1_grid_Invoice_Details']/tbody/tr[1]/td/table/tbody/tr/td/a[text()='"+str(j)+"']").click()
            time.sleep(2)
            j=j+1
        except NoSuchElementException:
            try:  
                browser.find_element_by_xpath("//*[@id='ctl00_ContentPlaceHolder1_grid_Invoice_Details']/tbody/tr[1]/td/table/tbody/tr/td[12]/a[text()='...']").click()
                time.sleep(2)
                j=j+1
            except NoSuchElementException:
                try:  
                    browser.find_element_by_xpath("//*[@id='ctl00_ContentPlaceHolder1_grid_Invoice_Details']/tbody/tr[1]/td/table/tbody/tr/td[11]/a[text()='...']").click()
                    time.sleep(2)
                    j=j+1
                except NoSuchElementException:
                    while True:
                        try:
                              browser.find_element_by_xpath("//*[@id='ctl00_ContentPlaceHolder1_grid_Invoice_Details']/tbody/tr[1]/td/table/tbody/tr/td[1]/a[text()='1']").click()
                              time.sleep(5)
                              break
                        except NoSuchElementException:
                              try:
                                  browser.find_element_by_xpath("//*[@id='ctl00_ContentPlaceHolder1_grid_Invoice_Details']/tbody/tr[1]/td/table/tbody/tr/td[1]/a[text()='...']").click()
                                  time.sleep(5)
                              except NoSuchElementException:
                                  break
                    break
        time.sleep(5)
db.close()
browser.close()
print "west bengal task complete for "+yesterday1

