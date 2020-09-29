import pandas as pd
import xlrd
from datetime import datetime
from pandas import ExcelWriter
from pandas import ExcelFile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui

pyautogui.FAILSAFE = False
# Empty because they are my personal information
usernameStr = ''
passwordStr = ''
# open a log file to put missing or PTs not found into the log
f = open('log.txt', 'w')
count = 1
# selenium open chrome then login
browser = webdriver.Chrome()
browser.get(('https://apps.epilepsygroup.com/nereg/servlet/com.nereg.seguridad.base.login'))
username = browser.find_element_by_id('vUSUARIOCODIGO')
username.send_keys(usernameStr)
global date
password = WebDriverWait(browser, 10).until(
EC.presence_of_element_located((By.ID, 'vUSUARIOPASSWORD')))
password.send_keys(passwordStr)
 
signInButton = browser.find_element_by_name('BTN_ENTER')
signInButton.click()
time.sleep(3)
browser.get(('https://apps.epilepsygroup.com/nereg/servlet/com.nereg.findpatient'))
# Function checks if the patient was found after search is executed
# counter is the current patient in the list because it is recursive
# It matches the date the patient was hooked up with the excel sheet
def findMatch(counter):
    global date
    try:
        aD = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.ID, 'span_W0016EMU_HOOKUPDATE_000' + str(counter))))
        if(aD.get_attribute("innerText") != date):
            counter +=1
            findMatch(counter)
        else:
            browser.execute_script("""gx.evt.execEvt('W0016',false,'W0016ECTL_UPDATE.CLICK.000""" + str(counter) + """',this,17); """)
            counter = 1
    except:
        print('Patient not Found: ' + df['Chart #'][i])
        f.write('Patient not Found: ' + df['Chart #'][i] + '\n')
        
            
#reads the excel sheet using pandas         
df = pd.read_excel('Data.xlsx', sheet_name='Sheet1')
 
for i in df.index:
    global date
    date = df['Adm.date'][i].date().strftime('%m/%d/%y')
    try:
        search = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.ID, 'vSEARCH')))
    except:
        browser.execute_script("""document.querySelector('#W0009SECCHANGE > a').click(); """)
        time.sleep(1)
        search = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.ID, 'vSEARCH')))
        print('Patient not Found from search: ' + df['Chart #'][i])
    # start search
    
    search.send_keys(df['Chart #'][i])
    sButton = browser.find_element_by_name('BTNBUSCAR')
    sButton.click()
    time.sleep(1)
    # Write to file if missing 
    isPresent = browser.find_elements_by_id("Grid_resultContainerRow_0001")
    if(len(isPresent) < 0):
        print('No values')
    pyautogui.moveTo(750, 405)
    time.sleep(1)
    pyautogui.moveTo(750, 406)
    pyautogui.click()
    time.sleep(1)
    
    browser.get(('https://apps.epilepsygroup.com/nereg/servlet/com.nereg.patientgestiondispatcher'))
    pyautogui.moveTo(140, 736)
    pyautogui.click()
    findMatch(count)
    time.sleep(3)
    
    # if it finds the patient then it goes into the chart information and searches
    # for the dvd field and enters the corresponding dvd number into the drive else
    # continues onto the next patient
    try:
        dvd = browser.find_element_by_id('vEMU_DVDEXTENALHARDDRIVE')
        dvd.clear()
        dvd.send_keys(str(df['DVD'][i]).split(".0", 1)[0])
        time.sleep(1)
        cButton = browser.find_element_by_name('ALTA')
        cButton.click()
        time.sleep(1)
        browser.execute_script("""document.querySelector('#W0009SECCHANGE > a').click(); """)
        time.sleep(1)
    except:
        continue
f.close()    
