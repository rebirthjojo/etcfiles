import requests
import pandas as pd
from openpyxl import load_workbook

def get_lotto_numbers(drw_no):
    url = f"https://www.dhlottery.co.kr/common.do?method=getLottoNumber&drwNo={drw_no}"
    try:
        response = requests.get(url).json()
        if response.get("returnValue") == "success":
            return [
                int(response['drwNo']), # 회차
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
        last_drw = 0 # 헤더만 있는 경우 대비
        
    next_drw = last_drw + 1
    new_rows = []

    while True:
        data = get_lotto_numbers(next_drw)
        if data:
            new_rows.append(data)
            print(f"{next_drw}회차 데이터 수집 완료")
            next_drw += 1
        else:
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