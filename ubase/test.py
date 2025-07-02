from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import win32com.client
import sys,os,time
import win32com.client
from datetime import datetime, timedelta


excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True  # 엑셀 프로그램 보이게 하기


file_path = '2025 HP 시리얼리스트.xlsx'
wb = excel.Workbooks.Open(os.path.abspath(file_path))
ws=wb.Sheets("22")
row =2
while True:
    date_value = ws.Cells(row, 5).Value  # 5열: 입고일시
    col3_value = ws.Cells(row, 3).Value  # 3열: 원하는 값
    product_value = ws.Cells(row, 1).Value  # 2열: 제품명
    if date_value is None:
        break
    try:
        # 엑셀 float 날짜 처리 (날짜+시간)
        if isinstance(date_value, float) or isinstance(date_value, int):
            date_value_dt = datetime(1899, 12, 30) + timedelta(days=float(date_value))
            setdate = datetime(2022, 4, 1)
        else:
            # 문자열일 경우 (예: '2022-05-18 14:46:00+00:00')
            date_value_dt = datetime.fromisoformat(str(date_value).replace('Z', '+00:00'))
            if date_value_dt.tzinfo is not None:
                setdate = datetime(2022, 4, 1, tzinfo=date_value_dt.tzinfo)
            else:
                setdate = datetime(2022, 4, 1)
        if date_value_dt >= setdate:
            print(product_value, date_value_dt, col3_value)
    except Exception as e:
        print(f"날짜 파싱 실패: {date_value} ({e})")
    row += 1

if  getattr(sys, 'frozen', False): 
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path)
else:
    driver = webdriver.Chrome()

driver.get('https://support.hp.com/kr-ko/check-warranty')
WebDriverWait(driver, 5).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="inputtextpfinder"]'))
)

