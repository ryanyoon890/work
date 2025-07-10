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
if  getattr(sys, 'frozen', False): 
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path)
else:
    driver = webdriver.Chrome()

file_path = '2025 HP 시리얼리스트.xlsx'
wb = excel.Workbooks.Open(os.path.abspath(file_path))
ws=wb.Sheets("22")
new_file_path = '시리얼코드.xlsx' # 새로운 파일 이름름
new_wb = excel.Workbooks.Add()
new_ws = new_wb.Sheets('Sheet1')  # 새로운 시트 가져오기
new_ws.Cells(1, 1).Value = '제품명'
new_ws.Cells(1, 2).Value = '시리얼'   
new_ws.Cells(1, 3).Value = '입고일시'
new_ws.Cells(1, 4).Value = '보증1'
new_ws.Cells(1, 5).Value = '보증2'
new_ws.Cells(1, 6).Value = '보증기간'
row =2
new_row=2  # 새로운 파일에 데이터를 쓸 때 사용할 행 번호
while True:
    date_value = ws.Cells(row, 5).Value  # 5열: 입고일시
    col3_value = ws.Cells(row, 3).Value  # 3열: 시리얼 값값
    product_value = ws.Cells(row, 1).Value  # 2열: 제품명
    if date_value is None:
        break
    try:
        if any(x in str(product_value) for x in ["280","400", "600", "800"]):
        # 엑셀 float 날짜 처리 (날짜+시간)
            if isinstance(date_value, float) or isinstance(date_value, int):
                date_value_dt = datetime(1899, 12, 30) + timedelta(days=float(date_value))
                setdate = datetime(2022, 4, 28)
            else:
                # 문자열일 경우 (예: '2022-05-18 14:46:00+00:00')
                date_value_dt = datetime.fromisoformat(str(date_value).replace('Z', '+00:00'))
                if date_value_dt.tzinfo is not None:
                    setdate = datetime(2022, 4, 28, tzinfo=date_value_dt.tzinfo)
                else:
                    setdate = datetime(2022, 4, 28)
            if date_value_dt >= setdate:
                driver.get('https://support.hp.com/kr-ko/check-warranty')
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="inputtextpfinder"]'))
                )
                driver.find_element(By.XPATH, '//*[@id="inputtextpfinder"]').send_keys(str(col3_value))
                driver.find_element(By.XPATH, '//*[@id="FindMyProduct"]').click()
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="warrantyStatus"]/div[2]/div[1]'))
                )
                warranty_info1= driver.find_element(By.XPATH, '//*[@id="warrantyStatus"]/div[2]/div[1]').text.strip()
                warranty_info2= driver.find_element(By.XPATH, '//*[@id="warrantyStatus"]/div[2]/div[2]').text.strip()
                tracker= driver.find_element(By.XPATH, '//*[@id="directionTracker"]/app-layout/app-check-warranty/div/div/div[2]/app-warranty-details/div/div[2]/main/div[4]/div/div[2]/div/div/div[1]/div[6]/div[2]').text.strip()
                print(product_value, date_value_dt, col3_value, warranty_info1, warranty_info2, tracker)
                new_ws.Cells(new_row,1).Value = product_value
                new_ws.Cells(new_row,2).Value = col3_value
                new_ws.Cells(new_row,3).Value = date_value_dt.strftime('%Y-%m-%d')
                new_ws.Cells(new_row,4).Value = warranty_info1
                new_ws.Cells(new_row,5).Value = warranty_info2
                new_ws.Cells(new_row,6).Value = tracker
                new_row+=1

    except Exception as e:
        print(f"날짜 파싱 실패: {date_value} ({e})")
    row += 1

new_wb.SaveAs(os.path.abspath(new_file_path))




