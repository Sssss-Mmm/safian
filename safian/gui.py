import tkinter as tk
from tkinter import ttk, messagebox
import traceback
from safian.core import OrderProcessor
from safian.parser import parse_order_text

class OrderApp:
    def __init__(self, root, master_file):
        self.root = root
        self.root.title("주문 발주서 자동화 (클립보드 AI 파서 1.0)")
        self.root.geometry("1100x650")
        
        # 폰트
        default_font = ("Malgun Gothic", 10)
        self.root.option_add("*Font", default_font)
        style = ttk.Style()
        style.configure(".", font=default_font)
        style.configure("Treeview", font=default_font)
        style.configure("Treeview.Heading", font=default_font, font_weight="bold")

        # 코어 초기화
        self.processor = OrderProcessor(master_file)
        self.master_file = master_file
        self.order_list = [] # Treeview와 연동할 대기 리스트

        self._create_ui()

    def _create_ui(self):
        # 상단 타이틀 & 설명 영역
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=(10, 0))
        ttk.Label(top_frame, text="✅ 1. 카카오톡 주문을 복사(Ctrl+C) 한 후, [클립보드 분석] 버튼을 눌러주세요.", font=("Malgun Gothic", 11, "bold")).pack(anchor="w")
        ttk.Label(top_frame, text="✅ 2. 빈칸이 자동으로 채워지면 내용을 확인하고 [추가(Enter)]를 누릅니다.", font=("Malgun Gothic", 11)).pack(anchor="w")

        # 1. 입력 영역 (폼)
        input_frame = ttk.LabelFrame(self.root, text="[ 주문 정보 확인 / 수정 ]")
        input_frame.pack(fill="x", padx=10, pady=10)

        self.entries = {}
        fields = [
            ("상호(거래처명)", "partner"), ("주문인", "orderer"), ("수취인", "mid_recipient"),
            ("전화번호(집)", "phone"), ("핸드폰(필수)", "mobile"), ("주소", "address"),
            ("힌트(상품명)", "product_hint"), ("⭐바코드", "barcode"), ("수량", "qty"),
            ("수수료", "fee"), ("배송비", "ship_fee"), ("배송메모", "memo")
        ]

        # 4열 배치
        for idx, (label_text, var_name) in enumerate(fields):
            row = idx // 4
            col = (idx % 4) * 2
            
            lbl = ttk.Label(input_frame, text=label_text)
            lbl.grid(row=row, column=col, padx=5, pady=8, sticky="e")
            
            entry = ttk.Entry(input_frame, width=22)
            entry.grid(row=row, column=col+1, padx=5, pady=8, sticky="w")
            self.entries[var_name] = entry
            
            # 주소와 메모는 조금 더 길게
            if var_name in ["address", "memo", "product_hint"]:
                entry.config(width=35)
                
            # 바코드 입력 시 엔터 치면 제품 다시 검색
            if var_name == "barcode":
                entry.bind("<Return>", lambda e: self._on_barcode_manual_search())

        # 2. 버튼 영역
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        btn_paste = tk.Button(btn_frame, text="📋 클립보드 분석 (Ctrl+V)", bg="#dcedc1", font=("Malgun Gothic", 10, "bold"), height=2, command=self.paste_and_analyze)
        btn_paste.pack(side="left", padx=5)
        
        btn_add = tk.Button(btn_frame, text="➕ 아래 목록에 추가 (Enter)", bg="#a8e6cf", font=("Malgun Gothic", 10, "bold"), height=2, command=self.add_item)
        btn_add.pack(side="left", padx=5)

        ttk.Button(btn_frame, text="➖ 선택 삭제 (Del)", command=self.remove_item).pack(side="left", padx=5)
        
        btn_save = tk.Button(btn_frame, text="💾 엑셀파일에 저장", bg="#ffaaa5", fg="white", font=("Malgun Gothic", 10, "bold"), height=2, command=self.export_to_excel)
        btn_save.pack(side="right", padx=5)

        # 3. 리스트 (Treeview) 영역
        tree_frame = ttk.LabelFrame(self.root, text="[ 발주 대기 리스트 ]")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 엑셀에 저장될 컬럼 구성
        cols = ("주문인", "주소", "핸드폰", "바코드", "상품명", "수량", "배송메모", "타입")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=10)
        
        # 컬럼 너비 설정
        widths = {"주문인": 70, "주소": 300, "핸드폰": 120, "바코드": 110, "상품명": 150, "수량": 50, "배송메모": 150, "타입": 60}
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=widths.get(col, 100), anchor="center")
            
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        # 단축키 설정
        self.root.bind("<Control-v>", lambda e: self.paste_and_analyze())
        self.root.bind("<Return>", lambda e: self.add_item())
        self.root.bind("<Delete>", lambda e: self.remove_item())

    # ------------------ 핵심 로직 ------------------ #
    def paste_and_analyze(self, event=None):
        """텍스트를 붙여넣고 곧바로 주문 데이터로 파싱"""
        # 만약 사용자가 특정 입력창(Entry)에 커서를 두고 Ctrl+V를 눌렀다면,
        # 전체 분석을 실행하지 않고 기본 텍스트 붙여넣기에 맡깁니다.
        if event and isinstance(event.widget, (tk.Entry, ttk.Entry)):
            return

        try:
            content = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("오류", "클립보드가 비어있습니다.")
            return

        if not content.strip(): return

        # AI 파서 가동 (parser.py)
        parsed = parse_order_text(content)

        # 복사한 내용이 표 형태(Tab 구분이 여러개)이거나 
        # 연락처/주소가 명확히 있거나, 내용이 15글자 이상이면 "전체 주문 복사"로 간주합니다.
        is_full_order = False
        if "\t" in content and len(content.split("\t")) > 3:
             is_full_order = True
        elif parsed.get("mobile") or parsed.get("phone") or parsed.get("address"):
             is_full_order = True
        elif len(content) > 30: # 15자는 단일 상품명 치고 너무 짧을 수 있어 길이를 늘림
             is_full_order = True

        if is_full_order:
            # 전체 주문이면 기존 폼 내용을 깨끗하게 비우기
            clear_fields = ["partner", "orderer", "mid_recipient", "mobile", "phone", "address", "product_hint", "barcode", "qty", "memo"]
            for f in clear_fields:
                 self.entries[f].delete(0, 'end')
        else:
            # 상품명 조각 등만 복사했을 때는 다른 정보(이름, 주소)를 유지하고 상품/바코드 칸만 초기화
            self.entries["product_hint"].delete(0, 'end')
            self.entries["barcode"].delete(0, 'end')
            # 텍스트 전체를 상품명 힌트로만 적용
            parsed = {"product_hint": content.strip()}
        
        # 파싱된 데이터 입력창에 채우기
        if parsed.get("partner"): self.entries["partner"].insert(0, parsed.get("partner"))
            
        if parsed.get("orderer"):
            self.entries["orderer"].insert(0, parsed.get("orderer"))
            
        mid = parsed.get("mid_recipient") or parsed.get("orderer")
        if mid: self.entries["mid_recipient"].insert(0, mid)
            
        if parsed.get("mobile"): self.entries["mobile"].insert(0, parsed.get("mobile"))
        if parsed.get("address"): self.entries["address"].insert(0, parsed.get("address"))
        if parsed.get("qty"): self.entries["qty"].insert(0, parsed.get("qty"))
        if parsed.get("memo"): self.entries["memo"].insert(0, parsed.get("memo"))
        
        hint = parsed.get("product_hint")
        if hint:
            self.entries["product_hint"].insert(0, hint)
            
            # [핵심] 상품명 힌트를 바탕으로 바코드 자동 검색 (core.py)
            products = self.processor.find_barcode_by_product_name(hint)
            if products:
                # 첫번째 제품의 바코드를 넣음
                main_barcode = products[0]["barcode"]
                self.entries["barcode"].insert(0, main_barcode)
                self.entries["barcode"].configure(foreground="blue")
            else:
                self.entries["barcode"].insert(0, "[검색실패] 직접입력")
                self.entries["barcode"].configure(foreground="red")
                
        # 포커스를 바코드로 이동시켜 사용자가 최종 확인하도록 유도
        self.entries["barcode"].focus()

    def _on_barcode_manual_search(self):
        """바코드 창에서 엔터쳤을때 수동으로 제품명 검색"""
        bc = self.entries["barcode"].get()
        if bc:
             res = self.processor.lookup_product_by_barcode(bc)
             if res:
                  messagebox.showinfo("검색 완료", f"제품명: {res[0]['product_name']}")
             else:
                  messagebox.showwarning("검색 실패", "등록되지 않은 바코드입니다.")

    def add_item(self, event=None):
        """현재 입력창의 데이터를 Treeview와 내부 리스트에 추가"""
        data = {key: entry.get().strip() for key, entry in self.entries.items()}
        
        if not data["mobile"] and not data["address"]:
             return # 빈 데이터 무시
             
        if not data["barcode"] or "검색실패" in data["barcode"]:
             messagebox.showwarning("바코드 누락", "정확한 바코드를 입력해주세요.")
             self.entries["barcode"].focus()
             return
             
        # 바코드를 바탕으로 해당 제품(본품+사은품) 조회
        products = self.processor.lookup_product_by_barcode(data["barcode"])
        
        items_to_add = products if products else [{'type': '수기', 'product_name': '알수없음', 'barcode': data["barcode"]}]
             
        for item in items_to_add:
            row_data = data.copy()
            row_data["product_name"] = item["product_name"]
            row_data["barcode"] = item["barcode"]
            
            # Treeview 표시 포맷 반영 (엑셀 컬럼과 얼버무려서 화면용으로만)
            values = (
                row_data["orderer"], row_data["address"], row_data["mobile"],
                item["barcode"], item["product_name"], row_data["qty"],
                row_data["memo"], item["type"]
            )
            
            # 엑셀 헤더: [거래처명, 주문번호, 주문인, 수취인, 전화번호, 핸드폰, 우편번호, 주소, 바코드, 제품명, 사은품, 수량, 수수료, 배송비, 배송메모]
            # 내부 저장용 Data 구성
            excel_data = {
                "거래처명": row_data["partner"],
                "주문번호": "",
                "주문인": row_data["orderer"],
                "수취인": row_data["mid_recipient"],
                "전화번호": row_data["phone"],
                "핸드폰": row_data["mobile"],
                "우편번호": "",
                "주소": row_data["address"],
                "바코드": item["barcode"],
                "제품명": item["product_name"],
                "사은품": "", # 구조상 별도로 넣을수도 있지만 일단 패스
                "수량": row_data["qty"],
                "수수료": row_data["fee"],
                "배송비": row_data["ship_fee"],
                "배송메모": row_data["memo"]
            }
            
            self.order_list.append(excel_data)
            self.tree.insert("", "end", values=values)
            
        # 추가 성공 시 입력창 깨끗하게 비우기 (반복 작업 편의성)
        self.entries["barcode"].delete(0, 'end')
        self.entries["product_hint"].delete(0, 'end')

    def remove_item(self, event=None):
        """Treeview 선택 삭제"""
        selected = self.tree.selection()
        if not selected: return
        
        # 뒤에서부터 삭제하여 인덱스 꼬임 방지
        for s in reversed(selected):
            idx = self.tree.index(s)
            del self.order_list[idx]
            self.tree.delete(s)

    def export_to_excel(self):
        answer = messagebox.askyesno("저장", f"{len(self.order_list)}건의 데이터를 엑셀 제일 아래에 추가합니다.\n진행하시겠습니까?")
        if not answer: return
        
        success, msg = self.processor.append_orders_to_excel(self.order_list)
        if success:
             messagebox.showinfo("저장 성공", msg)
             # 리스트 초기화 여부?
             self.tree.delete(*self.tree.get_children())
             self.order_list.clear()
        else:
             messagebox.showerror("저장 오류", msg)

# --- 독립 실행 테스트용 ---
if __name__ == "__main__":
    root = tk.Tk()
    app = OrderApp(root, "2026통합발주서_영업_연습.xlsb")
    root.mainloop()
