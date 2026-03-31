import os
import time
import pandas as pd
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 저장 폴더 설정
SAVE_DIR = "C:/daily_rate"
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

# ERP URL 확인 키워드
ERP_URL_KEYWORD = "https://qcaps.nov.com/"

# ChromeDriver 경로 설정 (수동 설치 필요)
CHROME_DRIVER_PATH =  "C:\TOOLS\chromedriver-win64/chromedriver.exe"

# 크롬 드라이버 생성
def create_driver():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    service = Service(CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# ERP 사이트 확인
def check_erp_site(driver, expected_url_keyword):
    current_url = driver.current_url
    print(f"🔍 현재 접속 중인 URL: {current_url}")  # 콘솔에 출력

    # 팝업창에도 현재 접속 URL을 띄움
    messagebox.showinfo("현재 ERP 접속 URL", f"{current_url}")

    # URL 안에 기대하는 키워드가 포함되어 있는지 확인
    if expected_url_keyword.lower() not in current_url.lower():
        messagebox.showerror(
            "오류", 
            f"❌ 현재 ERP 사이트가 지정한 ERP와 다릅니다.\n\n현재 URL: {current_url}\n기대 URL 키워드: {expected_url_keyword}"
        )
        driver.quit()
        exit(1)
    else:
        print("✅ ERP 사이트 정상 확인 완료")


# 환율 정보 가져오기 (네이버)
def get_exchange_rates(date_str, currency_codes=None):
    if currency_codes is None:
        currency_codes = ['USD', 'EUR', 'JPY', 'CNY', 'GBP']

    results = []
    date_url = datetime.strptime(date_str, '%Y-%m-%d').strftime('%Y%m%d')
    headers = {"User-Agent": "Mozilla/5.0"}

    for code in currency_codes:
        try:
            url = f"https://finance.naver.com/marketindex/exchangeDailyQuote.naver?marketindexCd=FX_{code}KRW&page=1&searchdate={date_url}"
            response = requests.get(url, headers=headers, verify=False)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            table = soup.select_one('table.tbl_exchange')

            rate = '정보 없음'
            if table:
                for row in table.select('tbody tr'):
                    cols = row.select('td')
                    if not cols:
                        continue
                    rate_date = cols[0].get_text(strip=True).replace('.', '-')
                    rate_val = cols[1].get_text(strip=True).replace(',', '')
                    if rate_date == date_str:
                        rate = float(rate_val)
                        break

            results.append({'날짜': date_str, '통화코드': code, '매매기준율(KRW)': rate})
        except Exception as e:
            results.append({'날짜': date_str, '통화코드': code, '매매기준율(KRW)': f"오류: {e}"})

    return results

# 환율 정보 엑셀로 저장
def save_to_excel(data, date_str):
    filename = f"{SAVE_DIR}/daily_{date_str.replace('-', '')}.xlsx"
    
    new_data = []
    for row in data:
        if isinstance(row['매매기준율(KRW)'], str):
            continue

        new_data.append({
            'From Currency': 'KRW',
            'To Currency': row['통화코드'],
            'Rate Date': row['날짜'],
            'Rate Type': 'KRW Daily',
            'Rate': row['매매기준율(KRW)']
        })
    
    df = pd.DataFrame(new_data)
    df.to_excel(filename, index=False)
    return filename

# 버튼1: 환율 가지고 오기
def fetch_exchange_rates():
    date_str = datetime.today().strftime("%Y-%m-%d")
    rates = get_exchange_rates(date_str)
    filename = save_to_excel(rates, date_str)
    messagebox.showinfo("완료", f"환율 저장 완료!\n{filename}")

# 버튼2: 환율 UPDATE
def update_erp_rates():
    driver = create_driver()

    # ERP 사이트 정상 체크
    #check_erp_site(driver, ERP_URL_KEYWORD)

    date_str = datetime.today().strftime("%Y-%m-%d")
    filename = f"{SAVE_DIR}/daily_{date_str.replace('-', '')}.xlsx"

    if not os.path.exists(filename):
        messagebox.showerror("오류", f"파일이 없습니다: {filename}")
        driver.quit()
        return

    df = pd.read_excel(filename)

    driver.find_element(By.XPATH, "//button[contains(text(),'Create Daily Rates')]").click()
    time.sleep(2)

    dates = [datetime.strptime(date_str, "%Y-%m-%d")]

    # 금요일이면 토/일도 입력
    weekday = dates[0].weekday()
    if weekday == 4:  # 금요일
        dates.append(dates[0] + timedelta(days=1))
        dates.append(dates[0] + timedelta(days=2))

    for date in dates:
        rate_date_str = date.strftime("%d-%b-%Y").upper()10008843-037

        for idx, row in df.iterrows():
            if pd.isna(row['Rate']):
                continue

            driver.find_element(By.XPATH, "//input[@title='From Currency']").send_keys(row['From Currency'])
            driver.find_element(By.XPATH, "//input[@title='To Currency']").send_keys(row['To Currency'])
            driver.find_element(By.XPATH, "//input[@title='Start Date']").send_keys(rate_date_str)
            driver.find_element(By.XPATH, "//select[contains(@title,'Rate Type')]").send_keys(row['Rate Type'])
            driver.find_element(By.XPATH, "//input[@title='Rate']").send_keys(str(row['Rate']))
            time.sleep(1)

        driver.find_element(By.XPATH, "//button[text()='Apply']").click()
        time.sleep(3)

    messagebox.showinfo("완료", "ERP 환율 등록 완료!")
    driver.quit()

# GUI 구성
root = tk.Tk()
root.title("환율 관리 프로그램")
root.geometry("300x200")

btn1 = tk.Button(root, text="환율 가지고 오기", command=fetch_exchange_rates, height=2)
btn1.pack(pady=10)

btn2 = tk.Button(root, text="환율 UPDATE", command=update_erp_rates, height=2)
btn2.pack(pady=10)

root.mainloop()
