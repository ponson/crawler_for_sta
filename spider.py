from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep
import pdb
import zipfile
import shutil
import send2trash
import pandas as pd
import openpyxl


options = webdriver.FirefoxOptions()
options.add_argument("-headless")
options.enable_downloads = True
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.helperApps.alwaysAsk.force", False)
options.set_preference("browser.download.dir", "/usr/src/app/download/")
options.set_preference("browser.download.dir", "d:\\download")
options.set_preference("browser.helperApps.neverAsk.openFile","text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,text/html,text/plain,application/msword,application/xml,application/octet-stream,binary/octet-stream")
options.set_preference("browser.helperApps.neverAsk.saveToDisk","text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,text/html,text/plain,application/msword,application/xml,application/octet-stream,binary/octet-stream")
driver = webdriver.Firefox(options=options)
driver.set_window_size(1680, 1200)
driver.implicitly_wait(5)

# url = 'https://atlas.stackline.com/'

FOLDER_PATH = "D:\\Download\\"

with open('acc.txt', 'r') as acc:
    account = acc.readline().rstrip(' \n')
    password = acc.readline().rstrip(' \n')


def run_download_process(p_driver, p_destname, p_main_select, p_sub_select):

    # driver.find_element(By.XPATH, '//*[contains(@class, "MuiModal-backdrop")]/..').get_attribute('innerHTML')
    WebDriverWait(p_driver, 30, 0.5).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.MuiModal-backdrop[aria-hidden="true"]')))
    sleep(2)
    p_driver.find_element(By.XPATH, p_main_select).click()
    WebDriverWait(p_driver, 30, 0.5).until(EC.visibility_of_element_located((By.XPATH, p_sub_select)))

    p_driver.find_element(By.CSS_SELECTOR, '.MuiSpeedDialIcon-icon').click()
    WebDriverWait(p_driver, 30, 0.5).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[aria-label~=Download]')))
    p_driver.find_element(By.CSS_SELECTOR, 'button[aria-label~=Download]').click()
    # p_driver.find_element(By.CSS_SELECTOR, '[role="dialog"] input[checked]').click()
    p_driver.save_screenshot('export total traffic data.png')
    p_driver.find_element(By.CSS_SELECTOR, '[role="dialog"] .sl-primary-button').click()

    p_driver.find_element(By.CSS_SELECTOR, '[aria-label="Download List"]').click()
    # WebDriverWait(p_driver, 30, 0.5).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.queued')))
    p_driver.save_screenshot('download file1.png')
    sleep(5)
    WebDriverWait(p_driver, 60, 0.5).until_not(EC.visibility_of_element_located((By.CSS_SELECTOR, '.queued')))
    file_name = p_driver.find_element(By.XPATH, '//*[@class="download"]/../..//p').get_attribute('aria-label')
    print(f"filename={file_name}")
    # WebDriverWait(p_driver, 60, 0.5).until_not(EC.visibility_of_element_located((By.CSS_SELECTOR, '.queued')))
    p_driver.save_screenshot('download file2.png')
    p_driver.find_element(By.XPATH, f'//*[@aria-label="{file_name}"]/../..//*[@class="download"]').click()
    p_driver.save_screenshot('download file3.png')
    zipname = file_name[0:-4] + '.zip'
    fullpath = FOLDER_PATH+zipname
    with zipfile.ZipFile(fullpath, 'r') as zf:
        zf.extractall(FOLDER_PATH)
    p_destname += file_name[-17:]
    shutil.move(FOLDER_PATH+file_name, FOLDER_PATH+p_destname)
    print(FOLDER_PATH+p_destname)
    send2trash.send2trash(FOLDER_PATH+zipname)
    driver.find_element(By.CSS_SELECTOR, 'button[aria-label~=Close]').click()

try:
    df = pd.read_excel(FOLDER_PATH + 'RequestList.xlsx')
    for ridx in range(len(df.index)):
        row = df.loc[ridx]
        ro = row['RO']
        cn = row['Country']
        pj = row['Project Name']
        url = row['Stackline URL']
        rm = row.iloc[3]
        print(f"{ro}_{cn}_{pj}_{rm}")

        driver.get(url)
        if ridx == 0:
            WebDriverWait(driver, 30, 0.5).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#signin-email')))
            driver.find_element(By.CSS_SELECTOR, '#signin-email').send_keys(account)
            driver.find_element(By.CSS_SELECTOR, '#signin-password').send_keys(password)
            driver.find_element(By.CSS_SELECTOR, '#signin-button').click()
        # WebDriverWait(driver, 30, 0.5).until(EC.presence_of_element_located((By.XPATH, '//*[text()="YTD"]')))
        # driver.find_element(By.XPATH, '//*[text()="YTD"]').click()
        # driver.find_element(By.XPATH, '//*[text()="Last Week"]').click()
        # WebDriverWait(driver, 30, 0.5).until(EC.presence_of_element_located((By.XPATH, '//*[@class="sl-header__dropdowns"]//*[text()="Last Week"]')))
        # driver.save_screenshot('selected date range.png')

        destname = ro +'_'+ cn + '_' + pj + '_'+ rm + '_Total_Clicks_'
        run_download_process(driver, destname, '//a[text()="Traffic - Total"]', '//h2[text()="Total Traffic"]')
        sleep(2)
        destname = ro +'_'+ cn + '_' + pj + '_'+ rm + '_Retail_Sales_'
        run_download_process(driver, destname, '//a[text()="Retail Sales"]', '//span[text()="Retail Sales"]')
        

except Exception as err:
        print(err)
        driver.save_screenshot('err.png')
        pdb.set_trace()
        print('error')
driver.quit()
pdb.set_trace()
print('end')