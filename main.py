import tkinter as tk
from safian.gui import OrderApp
import os
import sys

def main():
    # Redirect stderr to debug.log for crash capturing
    sys.stderr = open("debug.log", "a", encoding="utf-8")
    
    if getattr(sys, 'frozen', False):
        # Application is frozen
        # _MEIPASS contains bundled resources (like images/data added via --add-data)
        # sys.executable is the path to the exe file
        
        # We assume the master file is external (next to the exe)
        base_path = os.path.dirname(sys.executable)
    else:
        # Application is not frozen
        base_path = os.path.dirname(os.path.abspath(__file__))
        
    # Check for master file in common locations
    possible_paths = [
        os.path.join(base_path, '2026통합발주서_영업_연습.xlsb'),
        os.path.join(os.getcwd(), '2026통합발주서_영업_연습.xlsb'),
        '2026통합발주서_영업_연습.xlsb'
    ]
    
    master_file = '2026통합발주서_영업_연습.xlsb' # Default
    for p in possible_paths:
        if os.path.exists(p):
            master_file = p
            break
            
    if not os.path.exists(master_file):
         # If not found, let OrderApp handle it or prompt
         # But OrderApp expects a path.
         print(f"Error: Master file not found at {master_file}")
         # Could show a simple messagebox before tk init if needed, or just let App handle it (App handles it)

    root = tk.Tk()
    app = OrderApp(root, master_file)
    root.mainloop()

if __name__ == "__main__":
    main()
