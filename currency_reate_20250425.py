import os
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from dotenv import load_dotenv
load_dotenv()
import urllib3



# 🔇 SSL 인증서 경고 무시
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 🔐 .env 환경 변수 로드

#ORACLE_URL = os.getenv("ORACLE_URL")
ORACLE_URL="https://qcaps.nov.com/OA_HTML/AppsLocalLogin"
ORACLE_USER = os.getenv("ORACLE_USER")
ORACLE_PASS = os.getenv("ORACLE_PASS")
ORACLE_RESP = os.getenv("ORACLE_RESPONSIBILITY")

# 📥 네이버 환율 가져오기 (SSL 무시)
def get_exchange_rates_by_date(date_str, currency_codes=None):
    if currency_codes is None:
        currency_codes = ['USD', 'EUR', 'JPY', 'CNY', 'GBP']

    results = []
    date_url = datetime.strptime(date_str, '%Y-%m-%d').strftime('%Y%m%d')
    headers = {"User-Agent": "Mozilla/5.0"}

    for code in currency_codes:
        try:
            url = f"https://finance.naver.com/marketindex/exchangeDailyQuote.naver?marketindexCd=FX_{code}KRW&page=1&searchdate={date_url}"
            response = requests.get(url, headers=headers, verify=False)  # ⚠ SSL 인증 무시
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

# 🏦 ERP에 환율 자동 등록
def input_exchange_rates_to_erp(driver, responsibility_name, exchange_data, rate_type="KRW Daily"):
    try:
       # 1. 로그인 페이지 접속
        driver.get(ORACLE_URL)

        # 2. 로그인 필드 로딩 대기
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "usernameField")))

        # 3. 로그인 정보 입력
        driver.find_element(By.ID, "usernameField").send_keys(ORACLE_USER)
        driver.find_element(By.ID, "passwordField").send_keys(ORACLE_PASS)

        # 4. 로그인 버튼 클릭
        driver.find_element(By.XPATH, "//button[contains(@onclick, 'submitCredentials')]").click()
        
        
        # driver.find_element(By.LINK_TEXT, responsibility_name).click()
                
        # time.sleep(3)

        # driver.find_element(By.LINK_TEXT, "Daily Rates").click()
        # time.sleep(3)

        # driver.find_element(By.XPATH, "//button[contains(text(),'Create Daily Rates')]").click()
        # time.sleep(3)

        # for entry in exchange_data:
        #     if isinstance(entry['매매기준율(KRW)'], str):
        #         print(f"⚠️ {entry['통화코드']} 환율 정보 없음 → ERP 입력 생략")
        #         continue

        #     print(f"➤ {entry['통화코드']} 환율 입력 중...")

        #     driver.find_element(By.XPATH, "//input[@title='From Currency']").send_keys(entry['통화코드'])
        #     driver.find_element(By.XPATH, "//input[@title='To Currency']").send_keys("KRW")
        #     driver.find_element(By.XPATH, "//input[@title='Start Date']").send_keys(entry['날짜'])
        #     driver.find_element(By.XPATH, "//select[contains(@title,'Rate Type')]").send_keys(rate_type)
        #     driver.find_element(By.XPATH, "//input[@title='Rate']").send_keys(str(entry['매매기준율(KRW)']))
        #     driver.find_element(By.XPATH, "//button[text()='Apply']").click()
        #     time.sleep(3)

        #     print(f"✅ {entry['통화코드']} 환율 등록 완료!")

    except Exception as e:
        print(f"🚫 ERP 입력 중 오류 발생: {e}")

# 📄 엑셀 저장
def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"📄 엑셀 저장 완료: {filename}")

# ▶️ 메인 실행
if __name__ == "__main__":
    date_input = input("📅 조회할 날짜를 입력하세요 (YYYY-MM-DD): ").strip()
    try:
        datetime.strptime(date_input, '%Y-%m-%d')
        print("📡 환율 정보 가져오는 중...")
        rates = get_exchange_rates_by_date(date_input)

        for r in rates:
            print(f"{r['통화코드']} → {r['매매기준율(KRW)']}")

        #driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        # ChromeDriver 경로 설정
        driver_path = "C:\TOOLS\chromedriver-win64/chromedriver.exe"  # 실제 경로로 수정

        
        
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)  # 크롬 자동 닫힘 방지
        
        service = Service(driver_path)

        driver = webdriver.Chrome(service=service)
        print("\n🔐 Oracle ERP 로그인 및 등록 시작...")
        input_exchange_rates_to_erp(driver, ORACLE_RESP, rates)
        driver.quit()

        save_to_excel(rates, f"환율정보_{date_input}.xlsx")

    except ValueError:
        print("❌ 날짜 형식이 잘못되었습니다. 예: 2024-04-25")
