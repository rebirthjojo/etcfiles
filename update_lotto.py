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
        # data_only=False로 수식 보존
        wb = load_workbook(file_name, data_only=False)
        ws = wb['원본']
        
        # 1. 실제 데이터가 있는 진짜 마지막 행 찾기 (A열 기준)
        # max_row가 수식 때문에 뻥튀기 된 경우를 대비합니다.
        real_last_row = 0
        for i in range(ws.max_row, 0, -1):
            if ws.cell(row=i, column=1).value is not None:
                try:
                    # 셀의 값이 숫자인지 확인
                    int(ws.cell(row=i, column=1).value)
                    real_last_row = i
                    break
                except (ValueError, TypeError):
                    continue
        
        last_drw = int(ws.cell(row=real_last_row, column=1).value) if real_last_row > 0 else 1205
        target_drw = last_drw + 1
        
        print(f"포착된 마지막 회차: {last_drw} (행 번호: {real_last_row})")
        print(f"기록할 목표 회차: {target_drw} (행 번호: {real_last_row + 1})")

        # 2. 네이버 크롤링
        url = f"https://search.naver.com/search.naver?query={target_drw}회로또"
        headers = {"User-Agent": "Mozilla/5.0"}
        res = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(res.text, 'html.parser')
        balls = soup.select('.ball')

        if balls and len(balls) >= 6:
            nums = [int(b.text) for b in balls[:6]]
            
            # 3. 정확한 다음 행에 기록
            write_row = real_last_row + 1
            ws.cell(row=write_row, column=1, value=target_drw)
            for i, num in enumerate(nums):
                ws.cell(row=write_row, column=i + 2, value=num)
            
            # 4. 저장 및 확인
            wb.save(file_name)
            print(f"성공: {target_drw}회 저장 완료!")
            
            # 저장 후 용량 변화나 수정 시간을 강제로 갱신하기 위해 os.utime 사용 (Git 감지용)
            os.utime(file_name, None) 
        else:
            print(f"{target_drw}회차 데이터가 아직 없습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    update_excel()
