import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import os

def update_excel():
    file_name = "lotto_prob.xlsx"
    if not os.path.exists(file_name):
        print("파일을 찾을 수 없습니다.")
        return

    try:
        # data_only=False: 2번째 시트의 수식을 그대로 보존합니다.
        wb = load_workbook(file_name, data_only=False)
        ws = wb['원본']
        
        # 1. 서식을 무시하고 "숫자"가 있는 진짜 마지막 행 찾기
        real_last_row = 1
        last_drw = 0
        
        # A열을 위에서부터 훑으며 마지막 '정수'를 찾습니다.
        for i in range(1, ws.max_row + 1):
            val = ws.cell(row=i, column=1).value
            if isinstance(val, int):
                real_last_row = i
                last_drw = val
        
        # 만약 숫자를 하나도 못 찾았다면 (헤더만 있는 경우 등)
        if last_drw == 0:
            last_drw = 1205 # 기본값 설정
            
        target_drw = last_drw + 1
        print(f"포착된 최신 회차: {last_drw} (행 번호: {real_last_row})")
        print(f"네이버 수집 목표: {target_drw}")

        # 2. 네이버 크롤링
        url = f"https://search.naver.com/search.naver?query={target_drw}회로또"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"}
        
        res = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(res.text, 'html.parser')
        balls = soup.select('.ball')

        if balls and len(balls) >= 7:
            # 보너스 번호 포함 7개 수집
            nums = [int(b.text) for b in balls[:7]]
            
            # 3. 데이터 기록 (A열: 회차, B~G열: 당첨번호, H열: 보너스번호)
            write_row = real_last_row + 1
            ws.cell(row=write_row, column=1, value=target_drw)
            for i, num in enumerate(nums):
                # i가 0~6까지 돌면서 B, C, D, E, F, G, H열에 순서대로 기록합니다.
                ws.cell(row=write_row, column=i + 2, value=num)
            
            wb.save(file_name)
            print(f"성공: {target_drw}회차 (보너스 포함 7개) 저장 완료!")
        else:
            print(f"결과: {target_drw}회차 결과가 아직 없습니다.")

    except Exception as e:
        print(f"실행 중 에러 발생: {e}")

if __name__ == "__main__":
    update_excel()

