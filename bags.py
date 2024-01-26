from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import date
import pandas as pd
from tkinter import filedialog
import os

file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=(("Excel files", "*.xlsx;*.xls"),('CSV Files','*.csv'), ("all files", "*.*")))
file_path = os.path.abspath(file_path)
df = pd.read_excel(file_path,header=None)
df.columns=['BagID']
df['driver Name'] = ''
df['company Name'] = ''

def getdrivertaWebsite(baglist):
    options = Options()
    options.add_experimental_option('detach',True)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                        options = options)
    driver.execute_script("window.open('', '_blank');")
    driver.switch_to.window(driver.window_handles[0])
    url = ""
    driver.get(url)
    stationselector = driver.find_element('xpath','//*[@EXAMPLE_TAG=""]')
    stationselector.click()
    stationsearchbar = driver.find_element('xpath','//*[@EXAMPLE_TAG=""]')
    stationsearchbar.click()
    stationsearchbar.send_keys("")
    stationsearchbar.send_keys(Keys.RETURN)
    time.sleep(2)
    for bag in baglist:
        iframe = driver.find_element('xpath', "//iframe[@EXAMPLE_TAG='']")
        driver.switch_to.frame(iframe)
        containerelement = driver.find_element('xpath',"//EXAMPLE_TAG/EXAMPLE_TAG[contains(text(),'')]")
        containerelement.click()
        searchwindow = driver.find_element('xpath',"//EXAMPLE_TAG[@EXAMPLE_TAG = '']")
        searchwindow.click()
        # searchwindow.send_keys(Keys.CONTROL+'A')
        # searchwindow.send_keys(Keys.DELETE)
        searchwindow.send_keys(bag)
        searchwindow.send_keys(Keys.RETURN)
        searchbutton = driver.find_element('xpath',"//EXAMPLE_TAG/EXAMPLE_TAG[contains(text(),'Search')]")
        searchbutton.click()
        time.sleep(1)
        packages_element = driver.find_element('xpath',"//EXAMPLE_TAG[@EXAMPLE_TAG='']/parent::span")
        packages_element.click()
        time.sleep(2)
        try:
            pkgCode_element = driver.find_element('xpath',"//EXAMPLE_TAG[contains(text(),'')]/ancestor::tr//a[contains(text(),'')]")
            pkgCode = pkgCode_element.text
            driver.switch_to.window(driver.window_handles[-1])
            driver.get(url)
            searchwindow = driver.find_element('xpath','//*[@EXAMPLE_TAG=" "]')
            searchwindow.click()
            searchwindow.send_keys(Keys.CONTROL+'A')
            searchwindow.send_keys(Keys.DELETE)
            searchwindow.send_keys(pkgCode)
            searchwindow.send_keys(Keys.RETURN)
            time.sleep(2)
            iframe = driver.find_element('xpath', "//iframe[@EXAMPLE_TAG='']")
            driver.switch_to.frame(iframe)
            companyelement = driver.find_element('xpath', "//span[text()='Provider']/parent::p/following-sibling::p[@EXAMPLE_TAG")
            company = companyelement.text
            driverelement = driver.find_element('xpath',"//span[text()='Driver']/parent::p/following-sibling::p[@EXAMPLE_TAG]")
            driver = driverelement.text
            df.loc[df['BagID']==bag,'driver Name'] = driver
            df.loc[df['BagID']==bag,'company Name'] = company
        except NoSuchElementException as exception:
            print(exception,pkgCode)
        driver.switch_to.window(driver.window_handles[0])
        driver.back()
        time.sleep(1)
    driver.quit()


baglist = df.iloc[:,0].tolist()
getdrivertaWebsite(baglist)
df.to_excel(f'bags{date.today().strftime("%m-%d-%y")}.xlsx',index=False)


        

