import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import os

def update_excel():
    file_name = "lotto_prob.xlsx"
    
    if not os.path.exists(file_name):
        print(f"Error: {file_name} file missing.")
        return

    try:
        # data_only=False로 열어야 2번째 시트의 수식이 깨지지 않고 보존됩니다.
        wb = load_workbook(file_name, data_only=False)
        
        # 반드시 '원본' 시트를 지정해서 수식이 있는 2번째 시트를 보호합니다.
        if '원본' not in wb.sheetnames:
            print("에러: '원본' 시트를 찾을 수 없습니다. 시트 이름을 확인하세요.")
            return
            
        ws = wb['원본']
        
        # 1. 엑셀에서 마지막 회차 읽기 (A열 기준)
        last_row = ws.max_row
        while last_row > 1:
            val = ws.cell(row=last_row, column=1).value
            if isinstance(val, int):
                last_drw = val
                break
            last_row -= 1
        else:
            last_drw = 1205 # 기본값

        target_drw = last_drw + 1
        print(f"Target: {target_drw}")

        # 2. 네이버에서 데이터 가져오기
        url = f"https://search.naver.com/search.naver?query={target_drw}회로또"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"}
        
        res = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(res.text, 'html.parser')
        balls = soup.select('.ball')

        if balls and len(balls) >= 6:
            nums = [int(b.text) for b in balls[:6]]
            
            # 3. '원본' 시트의 마지막 다음 행에 데이터 기입
            new_row_idx = last_row + 1
            ws.cell(row=new_row_idx, column=1, value=target_drw)
            for i, num in enumerate(nums):
                ws.cell(row=new_row_idx, column=i + 2, value=num)
            
            # 4. 저장 (이때 2번째 시트의 수식은 건드리지 않음)
            wb.save(file_name)
            print(f"Success: {target_drw}회차 데이터 저장 성공!")
        else:
            print(f"Fail: {target_drw}회 데이터가 네이버에 없습니다.")

    except Exception as e:
        print(f"실행 중 에러: {e}")

if __name__ == "__main__":
    update_excel()
