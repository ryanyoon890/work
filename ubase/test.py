from bs4 import BeautifulSoup
from selenium import webdriver
import win32com.client
import sys,os,time
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True  # 엑셀 프로그램 보이게 하기

if  getattr(sys, 'frozen', False): 
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path)
else:
    driver = webdriver.Chrome()


driver.get('https://support.hp.com/kr-ko/check-warranty#multiple')

file_path = '2025 HP 시리얼리스트.xlsx'
wb = excel.Workbooks.Open(os.path.abspath(file_path))

time.sleep(10)

git add .
git commit -m "최신 코드 및 chromedriver 경로 처리 추가"
git push