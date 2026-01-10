import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import os

def update_excel():
    file_name = "lotto_prob.xlsx"
    if not os.path.exists(file_name): return

    try:
        # data_only=False로 수식 보호
        wb = load_workbook(file_name, data_only=False)
        ws = wb['원본']
        
        # 1. '진짜' 마지막 회차와 행 번호 찾기
        # A열을 2000행까지 훑어서 가장 큰 회차 숫자와 그 위치를 찾습니다.
        last_drw = 0
        last_row_idx = 0
        
        # 1행(헤더) 제외하고 2행부터 검사
        for i in range(2, 2000):
            val = ws.cell(row=i, column=1).value
            if isinstance(val, int):
                if val > last_drw:
                    last_drw = val
                    last_row_idx = i
            # 너무 긴 빈 구간이 나오면 중단 (성능 최적화)
            if val is None and i > last_row_idx + 10:
                break
        
        if last_drw == 0:
            print("회차 데이터를 찾을 수 없어 기본값(1205)을 사용합니다.")
            last_drw = 1205
            # 데이터가 하나도 없을 경우 대비 (보통 헤더 밑 2행부터 시작)
            last_row_idx = 1 

        target_drw = last_drw + 1
        print(f"현재 엑셀상 최신 회차: {last_drw} (위치: {last_row_idx}행)")
        print(f"네이버 수집 목표: {target_drw}")

        # 2. 네이버 크롤링
        url = f"https://search.naver.com/search.naver?query={target_drw}회로또"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"}
        
        res = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(res.text, 'html.parser')
        balls = soup.select('.ball')

        if balls and len(balls) >= 6:
            nums = [int(b.text) for b in balls[:6]]
            
            # 3. 마지막 행 바로 다음 줄에 기록
            new_row = last_row_idx + 1
            ws.cell(row=new_row, column=1, value=target_drw)
            for i, num in enumerate(nums):
                ws.cell(row=new_row, column=i + 2, value=num)
            
            wb.save(file_name)
            print(f"성공: {target_drw}회차를 {new_row}행에 기록했습니다.")
        else:
            # 아직 당첨 번호가 안 떴을 경우 (토요일 저녁 등)
            print(f"결과: {target_drw}회차 결과가 아직 네이버에 없습니다. (수집 중단)")

    except Exception as e:
        print(f"에러 발생: {e}")

if __name__ == "__main__":
    update_excel()
