import requests
import pandas as pd
from openpyxl import load_workbook
from bs4 import BeautifulSoup

def get_latest_drw_no_from_naver():
    """네이버 검색을 통해 현재 기준 최신 회차 번호를 가져옵니다."""
    url = "https://search.naver.com/search.naver?query=로또당첨번호"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"}
    try:
        res = requests.get(url, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        tag = soup.select_one('._lotto-btn-current') or soup.select_one('.btn_lotto_select')
        if tag:
            return int(tag.text.replace('회', '').strip())
    except:
        return None
    return None

def get_lotto_numbers(drw_no):
    url = f"https://www.dhlottery.co.kr/common.do?method=getLottoNumber&drwNo={drw_no}"
    try:
        response = requests.get(url).json()
        if response.get("returnValue") == "success":
            return [
                int(response['drwNo']),
                int(response['drwtNo1']), int(response['drwtNo2']), int(response['drwtNo3']),
                int(response['drwtNo4']), int(response['drwtNo5']), int(response['drwtNo6'])
            ]
    except:
        return None
    return None

def update_excel():
    file_name = "lotto_prob.xlsx"
    wb = load_workbook(file_name)
    ws = wb['원본']
    
    last_row = ws.max_row
    last_drw = ws.cell(row=last_row, column=1).value
    
    try:
        last_drw = int(last_drw)
    except (ValueError, TypeError):
        last_drw = 0
        
    # 네이버에서 최신 회차 확인
    latest_on_naver = get_latest_drw_no_from_naver()
    print(f"현재 엑셀 마지막 회차: {last_drw}")
    print(f"네이버 최신 회차: {latest_on_naver}")

    next_drw = last_drw + 1
    new_rows = []

    while True:
        data = get_lotto_numbers(next_drw)
        if data:
            new_rows.append(data)
            print(f"{next_drw}회차 데이터 수집 완료")
            next_drw += 1
        else:
            if latest_on_naver and next_drw <= latest_on_naver:
                print(f"{next_drw}회차가 네이버엔 있지만 API엔 아직 없습니다. 잠시 후 재시도 필요.")
            break

    if new_rows:
        for row_data in new_rows:
            ws.append(row_data)
        wb.save(file_name)
        print(f"총 {len(new_rows)}개의 회차가 추가되었습니다.")
    else:
        print("이미 최신 상태입니다.")

if __name__ == "__main__":
    update_excel()
