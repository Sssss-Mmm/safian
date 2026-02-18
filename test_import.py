import sys
import os
sys.path.append(os.getcwd())
try:
    from safian.gui import OrderApp
    print("OrderApp imported successfully.")
    if hasattr(OrderApp, 'paste_from_clipboard'):
        print("paste_from_clipboard exists.")
    else:
        print("paste_from_clipboard MISSING.")
        print(dir(OrderApp))
except Exception as e:
    print(f"Import error: {e}")
