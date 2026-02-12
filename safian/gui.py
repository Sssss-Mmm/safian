import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from safian.core import OrderProcessor
import os

class OrderApp:
    def __init__(self, root, master_file):
        self.root = root
        self.root.title("발주서 자동화 프로그램")
        self.root.geometry("1000x600")
        
        # Robust font selection
        kor_font = self._find_korean_font()
        if kor_font:
            print(f"Applying font: {kor_font}")
            default_font = (kor_font, 10)
            self.root.option_add("*Font", default_font)
            
            style = ttk.Style()
            style.configure(".", font=default_font)
            style.configure("Treeview", font=default_font)
            style.configure("Treeview.Heading", font=default_font)
        else:
            print("Warning: No Korean font detected.")
        
        self.processor = None
        self.master_file = master_file
        self.order_list = []

        self._init_processor()
        self._create_ui()

    def _find_korean_font(self):
        from tkinter import font
        available_fonts = sorted(font.families())
        
        # Save to file for debugging
        with open("fonts.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(available_fonts))
        print("Font list saved to fonts.txt")
        
        candidates = ["Malgun Gothic", "맑은 고딕", "Gulim", "굴림", "Batang", "바탕", "NanumGothic", "Noto Sans KR"]
        
        for cand in candidates:
            if cand in available_fonts:
                return cand
        
        for f in available_fonts:
            if "nanum" in f.lower() or "gothic" in f.lower() or "korea" in f.lower():
                return f
        
        return None

    def _init_processor(self):
        try:
            self.processor = OrderProcessor(self.master_file)
            messagebox.showinfo("성공", "데이터 로드 완료!")
        except Exception as e:
            messagebox.showerror("오류", f"데이터 로드 실패: {e}\n파일 경로: {self.master_file}")

    def _create_ui(self):
        # Input Frame
        input_frame = ttk.LabelFrame(self.root, text="주문 정보 입력")
        input_frame.pack(fill="x", padx=10, pady=5)

        # Fields map: Label -> Variable name
        self.entries = {}
        fields = [
            ("거래처명", "partner"), ("주문인", "orderer"), ("수취인", "mid_recipient"),
            ("전화번호", "phone"), ("핸드폰", "mobile"), ("주소", "address"),
            ("바코드", "barcode"), ("수량", "qty"), ("수수료", "fee"),
            ("배송비", "ship_fee"), ("배송메모", "memo")
        ]

        # Use grid layout
        for idx, (label_text, var_name) in enumerate(fields):
            row = idx // 3
            col = (idx % 3) * 2
            
            ttk.Label(input_frame, text=label_text).grid(row=row, column=col, padx=5, pady=5, sticky="e")
            entry = ttk.Entry(input_frame)
            entry.grid(row=row, column=col+1, padx=5, pady=5, sticky="w")
            self.entries[var_name] = entry
            
            # Special binding for barcode to auto-lookup?
            if var_name == "barcode":
                entry.bind("<FocusOut>", self._on_barcode_leave)

        # Buttons
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(btn_frame, text="추가", command=self.add_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="선택 삭제", command=self.remove_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="엑셀 저장", command=self.export_excel).pack(side="right", padx=5)

        # Treeview (List)
        cols = ("거래처명", "주문인", "수취인", "바코드", "제품명", "수량", "배송메모")
        self.tree = ttk.Treeview(self.root, columns=cols, show="headings")
        
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
            
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)

    def _on_barcode_leave(self, event):
        barcode = self.entries["barcode"].get()
        if not barcode or not self.processor:
            return
            
        product = self.processor.lookup_product(barcode)
        if product:
            # Maybe show product name in a label or status bar?
            # For now just print to console or simple popup confirming found
            # Or better, autofill inputs if we had product name field (we don't displayed in inputs, but we will store it)
            pass
        else:
            # messagebox.showwarning("경고", "상품을 찾을 수 없습니다.") # Might be annoying on every focus out
            pass

    def add_item(self):
        data = {key: entry.get() for key, entry in self.entries.items()}
        
        # Validation
        if not data["barcode"]:
            messagebox.showwarning("경고", "바코드는 필수입니다.")
            return

        # Lookup Product
        product_name = "N/A"
        gifts = []
        if self.processor:
            prod_info = self.processor.lookup_product(data["barcode"])
            if prod_info:
                product_name = prod_info['product_name']
                gifts = prod_info['gifts']
            else:
                if not messagebox.askyesno("확인", "등록되지 않은 품번입니다. 계속하시겠습니까?"):
                    return

        # Add to local list
        item_data = {
            **data,
            "product_name": product_name,
            "gifts": ", ".join(map(str, gifts))
        }
        self.order_list.append(item_data)

        # Add to Treeview
        values = (data["partner"], data["orderer"], data["mid_recipient"], 
                  data["barcode"], product_name, data["qty"], data["memo"])
        self.tree.insert("", "end", values=values)
        
        # Clear specific fields? Or keep for convenience? 
        # Usually clear Barcode, Qty, Memo. Partner/Orderer might repeat.
        self.entries["barcode"].delete(0, "end")
        self.entries["qty"].delete(0, "end")
        self.entries["memo"].delete(0, "end")

    def remove_item(self):
        selected = self.tree.selection()
        if not selected:
            return
            
        for item in selected:
            # Remove from internal list - this is tricky with index. 
            # Simple way: rebuild list or just match by ID. 
            # For this simple app, we might rely on index if synchronization is guaranteed.
            idx = self.tree.index(item)
            del self.order_list[idx]
            self.tree.delete(item)

    def export_excel(self):
        if not self.order_list:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return

        if self.processor:
            success = self.processor.generate_order_file(self.order_list, path)
            if success:
                messagebox.showinfo("성공", "파일 저장 완료!")
            else:
                messagebox.showerror("오류", "파일 저장 실패")
