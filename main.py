import tkinter as tk
import os
import sys
from safian.gui import OrderApp

def main():
    root = tk.Tk()
    
    # 엑셀 파일 이름 찾기
    possible_names = [
        "2026통합발주서_영업_연습.xlsb",
        os.path.join("dist", "2026통합발주서_영업_연습.xlsb"),
        os.path.join("..", "2026통합발주서_영업_연습.xlsb"),
    ]
    
    excel_file = ""
    for name in possible_names:
        if os.path.exists(name):
             excel_file = name
             break
             
    if not excel_file:
         # 파일이 없더라도 일단 빈 문자열로 실행, 에러 메시지 출력됨
         excel_file = "2026통합발주서_영업_연습.xlsb"

    app = OrderApp(root, excel_file)
    root.mainloop()

if __name__ == "__main__":
    main()
