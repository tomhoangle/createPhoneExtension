from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import pandas as pd
import numpy as np
#from openpyxl import load_workbook
#from openpyxl import Workbook
import time, os.path

def find_missing(numberList):
    start, end = numberList[0], numberList[-1]
    temp = sorted(set(range(start, end + 1)).difference(numberList))
    return [i for i in temp if i >= 1000]

#Available location: Boston, MA - 911, Chicago, IL - 911, London, UK - 911, Los Angeles, CA - 911, New York, NY - 911, Monroe, NC - 911, Plano, TX - 911
#San Francisco, CA - 911
#may set up location check later
def create_3cxAccount(firstName, lastName, emailAddress, location = 'Los Angeles, CA - 911', outbound = False):
    chrome_options = Options()
    download_directory = "C:\\3cx"
    chrome_options.add_experimental_option("prefs", {"download.default_directory": download_directory, "download.prompt_for_download": False})
    chrome_options.add_argument("--headless")
    WINDOW_SIZE = "3840,2160"
    chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
    chrome_options.add_argument("start-maximized")
    chrome_options.add_argument("disable-infobars")
    chrome_options.add_argument("--disable-extensions")
    driver = webdriver.Chrome(options=chrome_options)
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_directory}}
    command_result = driver.execute("send_command", params)
    wait = WebDriverWait(driver, 10)
    driver.get("http://3cxdev01/#/login")
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder*='User name or extension number']"))).send_keys("")#enter username here
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder*='Password']"))).send_keys("") #password here
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "button[class*='btn btn-lg btn-primary btn-block ng-scope']"))).click()
    driver.get("http://3cxdev01/#/app/extensions")
    wait.until(ec.visibility_of_element_located((By.ID, "btnExport"))).click()

    #Check if export list finish down. Max 10s
    if os.path.exists( "C:\\3cx\\extensions.csv" ):
        os.remove("C:\\3cx\\extensions.csv")
    now = time.time()
    last_time = now + 10
    while time.time() <= last_time:
        if os.path.exists( "C:\\3cx\\extensions.csv" ):
            break;

    #find list of available outbound and inbound number
    f=pd.read_csv("C:/3cx/extensions.csv")
    outboundList = pd.read_csv("exportDID.csv")
    keepCol = ['Number']
    f = f[keepCol]
    keepCol = ['MASK','INOFFICE_DEST_NUMBER']
    outboundList = outboundList[keepCol]
    availableEx = list(f.Number)
    availableEx = find_missing(availableEx) #list of available extensions
    availableEx = list(map(str, availableEx)) #convert to string
    outboundList = outboundList.loc[pd.isnull(outboundList).any(1),:]
    outboundList['MASK'].replace({'\*':''}, regex=True, inplace=True)
    outboundList = list(outboundList['MASK'])
    outboundList = [i[-4:] for i in outboundList if '310448' in i] #list of outbound number for 310448XXXX
    outboundToUse = set(availableEx).intersection(outboundList)
    outboundToUseList = list(outboundToUse)
    for x in list(outboundToUse):
        availableEx.remove(x)
    if(outbound == True):
        outboundToUseList.sort()
        extensionAdd = str(outboundToUseList[0])
    else:
        availableEx.sort()
        extensionAdd = str(availableEx[0])
    print(extensionAdd)

    #create extension
    wait.until(ec.visibility_of_element_located((By.ID, "btnAdd"))).click()
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Extension']"))).clear()
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Extension']"))).send_keys(extensionAdd)
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='First Name']"))).send_keys(firstName)
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Last Name']"))).send_keys(lastName)
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Email Address']"))).send_keys(emailAddress)
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Outbound Caller ID']"))).send_keys("310448" + extensionAdd)
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='ID']"))).clear()
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='ID']"))).send_keys(extensionAdd)
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Password']"))).clear()
    wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Password']"))).send_keys("dai123")
    element = driver.find_element_by_id("btnSave")
    coordinates = element.location_once_scrolled_into_view # returns dict of X, Y coordinates
    driver.execute_script('window.scrollTo({}, {});'.format(coordinates['x'], coordinates['y']))
    time.sleep(1)
    wait.until(ec.element_to_be_clickable((By.ID, "btnSave"))).click()
    wait.until(ec.url_changes(driver.current_url))

    if(outbound == True):
        #assign inbound rules
        driver.get("http://3cxdev01/#/app/inbound_rules")
        wait.until(ec.visibility_of_element_located((By.ID, "inputSearch"))).send_keys(extensionAdd)
        wait.until(ec.visibility_of_element_located((By.XPATH, f"//td[@class='ng-binding' and contains(., '310448{extensionAdd}')]"))).click()
        wait.until(ec.visibility_of_element_located((By.ID, "btnEdit"))).click()
        wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Inbound rule name']"))).clear()
        wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Inbound rule name']"))).send_keys(firstName + " " + lastName)
        Select(wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "select[class='form-control ng-pristine ng-untouched ng-valid ng-not-empty']")))).select_by_visible_text('Extension')
        wait.until(ec.visibility_of_element_located((By.XPATH, "//span[@class='ng-binding ng-scope' and contains(., '0000')]"))).click()
        wait.until(ec.visibility_of_element_located((By.XPATH, "//*[@id=\"app\"]/div/div[3]/div[2]/div/div[2]/inbound-rule-editor-contents/panel[2]/div/div[2]/ng-transclude/routing-control/div/div[3]/select-control/div/div/div/input[1]"))).send_keys(extensionAdd)
        wait.until(ec.visibility_of_element_located((By.XPATH, f"//span[@class='ng-binding ng-scope' and contains(., '{extensionAdd}')]"))).click()
        Select(wait.until(ec.visibility_of_element_located((By.XPATH, "//div[@class='form-group' and contains(., 'Destination for calls outside office hours')]/select")))).select_by_visible_text('Extension')
        wait.until(ec.visibility_of_element_located((By.XPATH, "//span[@class='ng-binding ng-scope' and contains(., '0000')]"))).click()
        wait.until(ec.visibility_of_element_located((By.XPATH, "//*[@id=\"app\"]/div/div[3]/div[2]/div/div[2]/inbound-rule-editor-contents/panel[2]/div/div[2]/ng-transclude/routing-control/div/div[8]/select-control/div/div/div/input[1]"))).send_keys(extensionAdd)
        driver.implicitly_wait(5)
        time.sleep(3)
        driver.find_elements_by_xpath(f"//span[@class='ng-binding ng-scope' and contains(., '{extensionAdd}')]")[1].click()
        wait.until(ec.visibility_of_element_located((By.ID, "btnSave"))).click()
        wait.until(ec.url_changes(driver.current_url))

    #Add to correct 911 group
    driver.get("http://3cxdev01/#/app/groups")
    wait.until(ec.visibility_of_element_located((By.XPATH, f"//td[@class='ng-binding' and contains(., '{location}')]"))).click()
    wait.until(ec.visibility_of_element_located((By.ID, "btnEdit"))).click()
    wait.until(ec.url_changes(driver.current_url))
    wait.until(ec.visibility_of_element_located((By.ID, "btnAdd"))).click()
    wait.until(ec.visibility_of_element_located((By.ID, "inputSearch"))).send_keys(extensionAdd)
    wait.until(ec.visibility_of_element_located((By.XPATH, f"//td[@class='checkbox-row ng-binding' and contains(., {extensionAdd})]"))).click()
    #timing issue, use sleep for now
    time.sleep(4)
    driver.find_element_by_xpath("/html/body/div[1]/div/div/div[3]/button[1]").click()
    time.sleep(4)
    driver.find_element_by_id("btnSave").click()
    wait.until(ec.url_changes(driver.current_url))
    return driver, extensionAdd

#test
#driver = create_3cxAccount('Test','Test','Test@gail.com',location = 'Los Angeles, CA - 911', outbound = False)
