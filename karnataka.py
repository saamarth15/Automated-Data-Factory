#importing libraries
import datetime
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from bs4 import BeautifulSoup
import pyodbc
#connectiong to database
server = 'abdpl-l2.database.windows.net'
database = 'ABDPL'
username = 'abdadmin'
password = 'abdpl@ho13'
driver= '{ODBC Driver 13 for SQL Server}'
db = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+ password)
cur = db.cursor()
yesterday = datetime.datetime.today() - datetime.timedelta(days=2)
yesterday = yesterday.strftime("%d/%m/%Y")
#delete the old values if you rerun the program for same day
cur.execute("SELECT * FROM KARNATAKA1 WHERE DATE='"+yesterday+"'")
s = cur.fetchall()
if s:
    cur.execute("DELETE FROM KARNATAKA1 WHERE DATE = '"+yesterday+"'")
    db.commit()
#data extraction procedure using selenium webdriver and storing to mysql
url = "http://webreport.ksbcl.com/p5rmwebreports/default.aspx"
browser = webdriver.Ie(executable_path='C:\Users\dloadmin\Desktop\IEDriverServer.exe')
browser.get(url)
username = browser.find_element_by_name("uname")
password1 = browser.find_element_by_id("pwd1")
username.send_keys("0011")
password1.send_keys("nsp@123")
submitButton = browser.find_element_by_id("image1") 
submitButton.click()
browser.get("http://webreport.ksbcl.com/p5rmwebreports/left.aspx")
window_before = browser.window_handles[0]
browser.find_element_by_xpath("(//a[contains(text(),'WR-98')])").click()
browser.implicitly_wait(10)
window_after = browser.window_handles[1]
browser.implicitly_wait(10)
browser.switch_to.window(window_after)
browser.implicitly_wait(10)
p = []
z = browser.find_elements_by_xpath("//select[@name='depot']/option[@value]")
for i in z:
    p.append(str(i.get_attribute("value")))
for k in  p:
    try:
        fromdate=browser.find_element_by_xpath("//input[@name='rfromdt']")
        fromdate.clear()
        fromdate.send_keys(yesterday)
    except NoSuchElementException:
        pass
    try:    
        todate=browser.find_element_by_xpath("//input[@name='rtodt']")
        todate.clear()
        todate.send_keys(yesterday)
    except NoSuchElementException:
        pass
    datalist = []
    browser.find_element_by_xpath("//select[@name='depot']/option[text()= '"+k+"']").click()
    browser.find_element_by_xpath("//input[@src='./img/vreport.gif']").click()
    browser.implicitly_wait(10)
    page = browser.page_source
    soup = BeautifulSoup(page,"lxml")
    table = soup.find_all('table')[2]
    table_rows = table.find_all('tr')
    for tr in table_rows[1:]:
        td = tr.find_all('td')
        row = [((i.text).encode('ascii','ignore').decode('ascii')).replace(',','') for i in td]
        datalist.append(row)
    for i in datalist:
        i.append(u""+yesterday)
    #data insertion
    for i in datalist:
        cur.execute("insert into KARNATAKA1(MMYY,PARTYNAME,SUPPLIERNAME,DEPOTNAME,ITEMNAME,NO_OF_BOTTLES,QTYINCBS,QTYINBTLS,CATEGORY,RANGENAME,DISTRICT,STATUS,MCAT,DATE) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",  (i[0],i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8],i[9],i[10],i[11],i[12],i[13]))
    db.commit()
    browser.back()
    browser.implicitly_wait(10)
#data cleaning . Removing garbage data
cur.execute("delete from KARNATAKA1 where MMYY = 'Grand Total' and date='"+yesterday+"'")
cur.execute("delete from KARNATAKA1 where DEPOTNAME = NCHAR(160) and date='"+yesterday+"'")
cur.execute("delete from KARNATAKA1 where NO_OF_BOTTLES='0' and date='"+yesterday+"'")
db.commit()
cur.execute("update KARNATAKA1 set QTYINBTLS='0' where QTYINBTLS='' or QTYINBTLS=NCHAR(160) and date='"+yesterday+"'")
db.commit()
cur.execute("update KARNATAKA1 set QTYINCBS='0' where QTYINCBS='' or QTYINCBS=NCHAR(160) and date='"+yesterday+"'")
db.commit()
db.close()
browser.close()
browser.switch_to.window(window_before)
browser.close()
print "\n\nkarnataka task complete for "+yesterday        
        
    


    
