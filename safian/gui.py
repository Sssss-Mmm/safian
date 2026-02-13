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
        btn_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(btn_frame, text="추가 (Enter)", command=self.add_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="삭제 (Del)", command=self.remove_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="엑셀 데이터 붙여넣기 (Ctrl+V)", command=self.paste_from_clipboard).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="엑셀 저장", command=self.export_to_excel).pack(side="right", padx=5)

        # Bind keys
        self.root.bind("<Return>", lambda e: self.add_item())
        self.root.bind("<Delete>", lambda e: self.remove_item())                                                                         
        self.root.bind("<Control-v>", lambda e: self.paste_from_clipboard())

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

    def paste_from_clipboard(self, event=None):
        try:
            content = self.root.clipboard_get()
            if not content:
                return

            rows = content.strip().split('\n')
            if not rows:
                return
            
            # Helper to parse a single row string into a data dict
            def parse_row(row_str):
                row_vals = row_str.split('\t')
                vals = [c.strip() for c in row_vals]
                
                mapping = {
                    0: 'partner',
                    1: 'orderer',
                    2: 'mid_recipient',
                    4: 'phone',
                    5: 'mobile',
                    7: 'address',
                    8: 'barcode',
                    11: 'qty',
                    12: 'fee',
                    13: 'ship_fee',
                    14: 'memo'
                }
                
                data = {}
                # Initialize with empty strings for all known fields
                for key in self.entries.keys():
                    data[key] = ""

                for src_idx, key in mapping.items():
                    if src_idx < len(vals):
                        data[key] = vals[src_idx]
                
                # Cleanup Qty
                if not data['qty'] or data['qty'] == '0' or not data['qty'].isdigit():
                    data['qty'] = '1'
                
                return data

            # Logic:
            # If multiple rows -> Auto-add all (silent mode for individual errors, report at end)
            # If single row -> Fill inputs so user can edit before adding
            
            if len(rows) > 1:
                added_count = 0
                error_count = 0
                
                for r in rows:
                    if not r.strip(): continue
                    data = parse_row(r)
                    if self._process_and_add_order(data, silent=True):
                        added_count += 1
                    else:
                        error_count += 1
                
                msg = f"{added_count}건 추가 완료."
                if error_count > 0:
                    msg += f"\n(실패/중복/수기입력 필요: {error_count}건)"
                messagebox.showinfo("완료", msg)
                
            else:
                # Single row - fill inputs
                data = parse_row(rows[0])
                for key, val in data.items():
                    entry = self.entries.get(key)
                    if entry:
                        entry.delete(0, 'end')
                        entry.insert(0, val)
                
                # Auto-lookup for single row
                if data['barcode']:
                     self._on_barcode_leave(None)

        except Exception as e:
            print(f"Clipboard paste error: {e}")
            messagebox.showerror("오류", f"클립보드 붙여넣기 실패: {e}")

    def add_item(self):
        # Gather data from entries
        data = {key: entry.get() for key, entry in self.entries.items()}
        
        if self._process_and_add_order(data, silent=False):
            # Clear specific fields on success
            self.entries["barcode"].delete(0, "end")
            self.entries["qty"].delete(0, "end")
            # self.entries["memo"].delete(0, "end")

    def _process_and_add_order(self, data, silent=False):
        # Validation
        if not data["barcode"]:
            if not silent: messagebox.showwarning("경고", "바코드는 필수입니다.")
            return False

        # Lookup Product
        items_to_add = []
        if self.processor:
            lookup_results = self.processor.lookup_product(data["barcode"])
            if lookup_results:
                items_to_add = lookup_results
            else:
                if not silent:
                    if not messagebox.askyesno("확인", f"등록되지 않은 품번 '{data['barcode']}'입니다. 계속하시겠습니까?"):
                        return False
                # If silent (bulk) or user said yes, add as manual
                items_to_add = [{'type': '수기', 'product_name': '직접입력', 'barcode': data["barcode"]}]
        else:
             # No processor loaded
             items_to_add = [{'type': '수기', 'product_name': '직접입력', 'barcode': data["barcode"]}]

        # Add to local list and Treeview
        base_qty = int(data["qty"]) if data["qty"].isdigit() else 1
        
        for item in items_to_add:
            row_data = data.copy()
            row_data["product_name"] = item['product_name']
            row_data["barcode"] = item['barcode'] 
            row_data["type"] = item['type']
            
            self.order_list.append(row_data)

            # Treeview display
            values = (data["partner"], data["orderer"], data["mid_recipient"], 
                      row_data["barcode"], f"[{item['type']}] {item['product_name']}", 
                      data["qty"], data["memo"])
            self.tree.insert("", "end", values=values)
            
        return True


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

    def export_to_excel(self):
        if not self.order_list:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                            filetypes=[("Excel files", "*.xlsx"), ("Excel Binary", "*.xlsb")])
        if not path:
            return

        if self.processor:
            success, msg = self.processor.generate_order_file(self.order_list, path)
            if success:
                messagebox.showinfo("성공", "파일 저장 완료!")
            else:
                messagebox.showerror("오류", msg)
