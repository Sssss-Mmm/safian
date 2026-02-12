import tkinter as tk
from safian.gui import OrderApp
import os
import sys

def main():
    base_path = os.path.dirname(os.path.abspath(__file__))
    master_file = os.path.join(base_path, '2026통합발주서_영업_연습.xlsb')
    
    if not os.path.exists(master_file):
        # Allow running from one level up if needed or check current dir
        if os.path.exists('2026통합발주서_영업_연습.xlsb'):
            master_file = '2026통합발주서_영업_연습.xlsb'
        else:
            print(f"Error: Master file not found at {master_file}")
            # Could show a simple messagebox before tk init if needed, or just let App handle it (App handles it)

    root = tk.Tk()
    app = OrderApp(root, master_file)
    root.mainloop()

if __name__ == "__main__":
    main()
