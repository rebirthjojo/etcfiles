import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import os

def get_lotto_from_naver(drw_no):
    """네이버 검색 결과에서 특정 회차의 번호를 긁어옵니다."""
    url = f"https://search.naver.com/search.naver?query={drw_no}회로또"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
    }
    
    try:
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        ball_elements = soup.select('.ball')
        if not ball_elements:
            return None
            
        nums = [int(ball.text) for ball in ball_elements[:6]]
        return [drw_no] + nums
    except Exception as e:
        print(f"{drw_no}회차 수집 중 오류: {e}")
        return None

def update_excel():
    file_name = "lotto_prob.xlsx"
    if not os.path.exists(file_name):
        print(f"오류: {file_name} 파일이 없습니다.")
        return

    wb = load_workbook(file_name)
    ws = wb['원본']
    
    last_row = ws.max_row
    last_drw = ws.cell(row=last_row, column=1).value
    
    try:
        last_drw = int(last_drw)
    except:
        last_drw = 0
        
    next_drw = last_drw + 1
    new_rows = []

    print(f"현재 엑셀 마지막 회차: {last_drw}. {next_drw}회차부터 확인을 시작합니다.")

    while True:
        data = get_lotto_from_naver(next_drw)
        if data:
            new_rows.append(data)
            print(f"{next_drw}회차 네이버 데이터 수집 성공: {data[1:]}")
            next_drw += 1
        else:
            break

    if new_rows:
        for row_data in new_rows:
            ws.append(row_data)
        wb.save(file_name)
        print(f"업데이트 완료: 총 {len(new_rows)}개 회차 추가.")
    else:
        print("네이버에도 아직 새로운 회차가 없습니다. (이미 최신 상태)")

if __name__ == "__main__":
    update_excel()
