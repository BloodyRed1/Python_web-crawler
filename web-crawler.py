#爬蟲自動下載每日電錶收集的用電資料
#(新)

# -*- coding: utf-8 -*-
"""
Created on Wed Feb 21 14:11:17 2024

@author: T3615
"""
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException
import signal
import sys
import requests
import zipfile

def handle_interrupt(signum, frame):
    
    print("\n接收到中断信號，程序即将退出。")
    browser.quit()
    sys.exit(0)
def check_internet_connection():
    try:
        # 發送一個簡單的 GET 请求到一个已知的網站
        response = requests.get("https://www.google.com", timeout=5)
        # 如果響應狀態碼為 200，则表示連接正常
        return response.status_code == 200
    except requests.ConnectionError:
        # 連接錯誤，表示沒有網路連接
        return False
print("按Control+C可以退出程式。")


mm_value = "0" # 初始化 mm_value
# 獲取當前月份
current_month = datetime.now().strftime("%m")
# 獲取當前年份
current_year = datetime.now().year
YYt=int(current_year)
MMt=int(current_month)
if MMt < 13 and YYt <= current_year:
    for i in range(13):
        
        if  YYt!=2023 and i == MMt:
            yy_value= str(YYt-2023)
            mm_value = str(i-1)
internet_connected = check_internet_connection()
if internet_connected:
    print("網路連線正常。即將開始下載必備的軟體")
    
    # 指定 Chrome Driver 版本和下载链接
    chrome_driver_url = "https://storage.googleapis.com/chrome-for-testing-public/119.0.6045.105/win64/chromedriver-win64.zip"
    chrome_url="https://storage.googleapis.com/chrome-for-testing-public/119.0.6045.105/win64/chrome-win64.zip"
    # 指定保存 Chrome Driver 的文件夹路径
    chrome_driver_folder = os.getcwd()
    new_chrome_driver_path= chrome_driver_folder +"\\chromedriver-win64\\chromedriver.exe"
    new_chrome_path= chrome_driver_folder +"\\chrome-win64\\chrome.exe"
    # 如果保存 Chrome Driver 的文件夹不存在，则创建它
    if not os.path.exists(chrome_driver_folder):
        os.makedirs(chrome_driver_folder)
    # 指定保存 Chrome Driver 文件的路径
    chrome_driver_path = os.path.join(chrome_driver_folder, "chromedriver.zip")
    chrome_path = os.path.join(chrome_driver_folder, "chrome-win64.zip")
    # 下载 Chrome Driver 文件
    print("開始下载 Chrome Driver...")
    print("開始下载 Test Chrome ...")
    response = requests.get(chrome_driver_url)
    with open(chrome_driver_path, "wb") as f:
        f.write(response.content)
    response2 = requests.get(chrome_url)
    with open(chrome_path, "wb") as Z:
        Z.write(response2.content)
    print("Chrome Driver 下载完成。")
    print(" Test Chrome 下載完成...")
    print("開始解壓 Chrome Driver...")
    with zipfile.ZipFile(chrome_path, "r") as zip_ref:
        zip_ref.extractall(chrome_driver_folder)
    with zipfile.ZipFile(chrome_driver_path, "r") as zip_ref:
        zip_ref.extractall(chrome_driver_folder)    
    print("解壓 Chrome Driver...完成")
    os.remove(chrome_path)
    os.remove(chrome_driver_path)

        


# 獲取當前腳本所在的目錄
current_directory = os.getcwd()

# 指定新資料夾的名稱
new_folder_name = 'Today_Power_Meter_Data'
    # 使用os.path.join()建立新資料夾的相對路徑
relative_path = os.path.join(current_directory, new_folder_name)

# 判斷資料夾是否已經存在，如果不存在則創建
if not os.path.exists(relative_path):
    os.makedirs(relative_path)
    print(f"成功創建資料夾: {relative_path}")
else:
    print(f"資料夾已經存在: {relative_path}")

# 將當前目錄設置為新資料夾
download_folder = relative_path

chrome_driver_path = new_chrome_driver_path

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.binary_location =new_chrome_path
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
    
})

service = Service(chrome_driver_path)
browser = webdriver.Chrome(service=service, options=chrome_options)
domain_url = 'http://172.16.11.60/#home/main'
 

browser.get(
    f'{domain_url}/statistics/statisticsList?type=05&subType=225')
signal.signal(signal.SIGINT, handle_interrupt)
password_locator = (By.ID, 'loginPassword')
# 使用 WebDriverWait 等待元素可見
password = WebDriverWait(browser, 5000).until(
    EC.visibility_of_element_located(password_locator)
)

# 使用 WebDriverWait 等待元素可交互
password = WebDriverWait(browser, 5000).until(
    EC.element_to_be_clickable(password_locator)
)
password.send_keys("Admin")
password.send_keys(Keys.RETURN)
#password.send_keys(Keys.RETURN)  # 或者使用 submit 方法，取决于网页的实现
button_locator2 = (By.ID, 'popupOKButton')
def wait_for_element(driver, element_id, timeout=10):
    try:
        element_present = EC.presence_of_element_located((By.ID, element_id))
        WebDriverWait(driver, timeout).until(element_present)
        return True
    except Exception as e:
        print("登入失敗")
        signal.signal(signal.SIGINT, handle_interrupt)   
        return False
while not wait_for_element(browser,'popupOKButton'):
    password.send_keys(Keys.RETURN)
    time.sleep(1)
    print("重新嘗試登入")
    signal.signal(signal.SIGINT, handle_interrupt)   
button2 = WebDriverWait(browser, 5000).until(
    EC.visibility_of_element_located(button_locator2)
)
signal.signal(signal.SIGINT, handle_interrupt)
# 使用 WebDriverWait 等待元素可交互
button2 = WebDriverWait(browser, 5000).until(
    EC.element_to_be_clickable(button_locator2)
)
button2.send_keys(Keys.RETURN)

main_page_element = browser.find_element(By.XPATH, '//li[contains(text(), "主頁面")]')


# 點擊 "主頁面" 元素
main_page_element.click()

wait = WebDriverWait(browser, 10)

history_chart_element = wait.until(EC.presence_of_element_located((By.XPATH, '//li[contains(@location, "#home/main/hy_chart!0")]')))

# 點擊 "歷史圖表頁面" 元素
history_chart_element .click()

select_element = WebDriverWait(browser, 20,1).until(
EC.presence_of_element_located((By.ID, "select_pm"))
                )
# 創建 Select 對象
select = Select(select_element)

# 獲取所有選項
options = select.options
signal.signal(signal.SIGINT, handle_interrupt)         
for option in options:
    try:
        # 選擇下拉列表中的下一個選項
        select.select_by_value(option.get_attribute("value"))
        time.sleep(0.5)
        # 重新查找下載按鈕
        download_button = WebDriverWait(browser,20,1).until(
           EC.presence_of_element_located((By.ID, "download_btn"))
                                   )
        # 點擊下載按鈕
        
        download_button.click()
        time.sleep(0.1)
        # 等待一段時間以確保相應的數據已加載（視具體網頁情況而定）
        #time.sleep(2)
    except StaleElementReferenceException:
                print("Inner StaleElementReferenceException. Retrying...") 
    except StaleElementReferenceException:
                print("Outer StaleElementReferenceException. Retrying...")   
time.sleep(1)
browser.quit()