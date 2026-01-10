import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import os

def update_excel():
    file_name = "lotto_prob.xlsx"
    
    if not os.path.exists(file_name):
        print(f"Error: {file_name} not found.")
        return

    wb = load_workbook(file_name)
    ws = wb['원본']
    
    # 1. 마지막 회차 읽기
    last_row = ws.max_row
    last_drw = ws.cell(row=last_row, column=1).value
    
    try:
        last_drw = int(last_drw)
    except:
        last_drw = 0
    
    # 2. 안전장치 (1200회 미만 인식 시 1205회로 고정)
    if last_drw < 1200:
        last_drw = 1205

    target_drw = last_drw + 1
    print(f"Target: {target_drw}")

    # 3. 네이버 조회
    url = f"https://search.naver.com/search.naver?query={target_drw}회로또"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"}
    
    try:
        response = requests.get(url, headers=headers, timeout=5)
        soup = BeautifulSoup(response.text, 'html.parser')
        ball_elements = soup.select('.ball')
        
        if ball_elements and len(ball_elements) >= 6:
            nums = [int(ball.text) for ball in ball_elements[:6]]
            # 마지막 행에 중복 체크 후 추가
            if ws.cell(row=ws.max_row, column=1).value != target_drw:
                ws.append([target_drw] + nums)
                wb.save(file_name)
                print(f"Success: {target_drw}회 추가 완료")
            else:
                print("Already updated.")
        else:
            print("Data not found on Naver.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    update_excel()
