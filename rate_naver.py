import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import urllib3

# ✅ SSL 인증서 오류 경고 무시
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_exchange_rates_by_date(date_str, currency_codes=None):
    if currency_codes is None:
        currency_codes = ['USD', 'EUR', 'JPY', 'CNY', 'GBP']

    results = []
    date_url = datetime.strptime(date_str, '%Y-%m-%d').strftime('%Y%m%d')
    headers = {"User-Agent": "Mozilla/5.0"}

    for code in currency_codes:
        try:
            url = f"https://finance.naver.com/marketindex/exchangeDailyQuote.naver?marketindexCd=FX_{code}KRW&page=1&searchdate={date_url}"
            response = requests.get(url, headers=headers, verify=False)  # 🔧 SSL 인증 무시
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

def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"✅ 엑셀 저장 완료: {filename}")

if __name__ == "__main__":
    date_input = input("📅 조회할 날짜를 입력하세요 (YYYY-MM-DD): ").strip()
    try:
        datetime.strptime(date_input, '%Y-%m-%d')  # 날짜 포맷 검증
        exchange_data = get_exchange_rates_by_date(date_input)
        excel_filename = f"환율정보_{date_input}.xlsx"
        save_to_excel(exchange_data, excel_filename)
    except ValueError:
        print("❌ 날짜 형식이 잘못되었습니다. 예: 2024-04-25")
