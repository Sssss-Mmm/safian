import pandas as pd
import os
from datetime import datetime

class OrderProcessor:
    def __init__(self, master_file_path):
        self.master_file_path = master_file_path
        self.products_df = None
        self.addresses_df = None
        self._load_master_data()

    def _log(self, msg):
        """Logs message to a file for debugging frozen apps."""
        with open("debug.log", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now()}: {msg}\n")

    def _load_master_data(self):
        """Loads product and address data from the master xlsb file."""
        if not os.path.exists(self.master_file_path):
            self._log(f"Master file not found: {self.master_file_path}")
            raise FileNotFoundError(f"Master file not found: {self.master_file_path}")

        self._log(f"Loading master data from: {self.master_file_path}")
        try:
            # Load '코드' sheet. 
            # Analysis showed proper headers at Row 0 (Excel Row 1) containing '품번', '제품명'.
            self.products_df = pd.read_excel(self.master_file_path, sheet_name='코드', engine='pyxlsb', header=0)
            
            # Normalize column names
            self.products_df.columns = [str(c).strip() for c in self.products_df.columns]
            self._log(f"Product Columns: {self.products_df.columns.tolist()}")
            
            # Load '25.12 주소' sheet
            self.addresses_df = pd.read_excel(self.master_file_path, sheet_name='25.12 주소', engine='pyxlsb', header=1)
            self.addresses_df.columns = [str(c).strip() for c in self.addresses_df.columns]
            self._log(f"Address Columns: {self.addresses_df.columns.tolist()}")
            
            print("Master data loaded successfully.")
            
        except Exception as e:
            self._log(f"Failed to load master data: {e}")
            raise Exception(f"Failed to load master data: {e}")

    def lookup_product(self, barcode):
        """Searches for a product by barcode (품번) and returns main product + gifts."""
        if self.products_df is None:
            return []

        # clean input
        barcode = str(barcode).strip()
        self._log(f"Looking up barcode: '{barcode}'")
        
        target_col = '품번'
        if target_col not in self.products_df.columns:
            self._log(f"Column '{target_col}' not found in {self.products_df.columns.tolist()}")
            return []

        # Filter
        df_clean = self.products_df[self.products_df[target_col].notna()].copy()
        df_clean[target_col] = df_clean[target_col].astype(str).str.strip()
        
        result = df_clean[df_clean[target_col] == barcode]
        
        if result.empty:
            self._log("No match found")
            return []
        
        # Return the first match
        row = result.iloc[0]
        
        items = []
        
        # 1. Main Product
        main_item = {
            'type': '본품',
            'product_name': row.get('제품명', ''),
            'barcode': barcode # Keep barcode for main item
        }
        items.append(main_item)
        
        # 2. Gifts
        for i in range(1, 6):
            gift_col = f'사은품 {i}'
            if gift_col in row and pd.notna(row[gift_col]):
                gift_barcode = str(row[gift_col]).strip()
                if gift_barcode and gift_barcode != 'nan':
                    # Look up gift name using the same dataframe
                    gift_name = ""
                    gift_match = df_clean[df_clean[target_col] == gift_barcode]
                    if not gift_match.empty:
                        gift_name = gift_match.iloc[0].get('제품명', '')
                    
                    items.append({
                        'type': '사은품',
                        'product_name': gift_name,
                        'barcode': gift_barcode 
                    })
        
        self._log(f"Match found: {items}")     
        return items

    def generate_order_file(self, order_data_list, output_path, overwrite=False):
        """
        Generates order file. If output_path ends with .xlsb, uses Excel Automation.
        Otherwise uses pandas for .xlsx.
        """
        if not order_data_list:
            return False, "데이터가 없습니다."
            
        # Define output columns order
        columns = [
            "거래처명", "주문인", "수취인", 
            "전화번호", "핸드폰", "주소", 
            "바코드", "제품명", "사은품", "수량", 
            "수수료", "배송비", "배송메모"
        ]

        # Check file extension
        if output_path.lower().endswith('.xlsb'):
            return self._save_to_xlsb_win32(order_data_list, output_path, columns, overwrite)
        
        # Standard XLSX export
        # overwrite is implicit for xlsx as we create new file, 
        # but if we wanted to append to xlsx it would be different.
        # For now, xlsx logic creates new file always as per original design.
        
        df = pd.DataFrame(order_data_list)
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        df = df[columns]
        
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='발주서')
                # Auto-adjust column width
                worksheet = writer.sheets['발주서']
                from openpyxl.utils import get_column_letter
                for idx, col in enumerate(df.columns):
                    max_len = len(str(col)) + 2
                    for val in df[col].head(10): 
                        if val:
                            max_len = max(max_len, len(str(val)) + 2)
                    col_letter = get_column_letter(idx + 1)
                    worksheet.column_dimensions[col_letter].width = min(max_len, 50)

            print(f"Order file created at: {output_path}")
            return True, "성공"
        except Exception as e:
            error_msg = f"파일 저장 실패: {str(e)}"
            print(error_msg)
            return False, error_msg

    def _save_to_xlsb_win32(self, order_data_list, output_path, columns, overwrite=False):
        """Modified existing XLSB using Excel Application (Windows Only)."""
        import platform
        if platform.system() != 'Windows':
            return False, "XLSB 수정은 윈도우 환경에서만 가능합니다."
            
        try:
            import win32com.client
            import pythoncom
        except ImportError:
            return False, "pywin32 라이브러리가 필요합니다. (pip install pywin32)"

        excel = None
        workbook = None
        try:
            pythoncom.CoInitialize() # Needed for some environments
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False # Run in background
            # excel.DisplayAlerts = False # Be careful with this, might overwrite without warning

            abs_path = os.path.abspath(output_path)
            
            if not os.path.exists(abs_path):
                 return False, "대상 XLSB 파일이 존재하지 않습니다."

            workbook = excel.Workbooks.Open(abs_path)
            
            # Find or Create '발주내역' sheet
            target_sheet_name = '발주내역'
            sheet = None
            try:
                sheet = workbook.Sheets(target_sheet_name)
            except:
                # Create new sheet if not exists
                sheet = workbook.Sheets.Add(After=workbook.Sheets(workbook.Sheets.Count))
                sheet.Name = target_sheet_name
                # Write Headers
                for col_idx, col_name in enumerate(columns):
                    sheet.Cells(1, col_idx + 1).Value = col_name
            
            start_row = 0
            
            if overwrite:
                # Find last used row to clear
                last_used_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
                if last_used_row > 1:
                    # Clear from row 2 to last used row
                    # Use Range object
                    range_to_clear = sheet.Range(f"A2:M{last_used_row}") # Assuming M is last column (13th)
                    range_to_clear.ClearContents()
                
                start_row = 2
            else:
                # Find last row to append
                last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row # xlUp
                if last_row == 1 and sheet.Cells(1, 1).Value is None:
                    last_row = 0 # Empty sheet
                start_row = last_row + 1
            
            # Write data row by row
            for i, data in enumerate(order_data_list):
                current_row = start_row + i
                for col_idx, col_name in enumerate(columns):
                    val = data.get(col_name, "")
                    sheet.Cells(current_row, col_idx + 1).Value = str(val)
            
            workbook.Save()
            return True, "XLSB 파일에 저장 완료"
            
        except Exception as e:
            return False, f"Excel 자동화 오류: {e}"
        finally:
            if workbook:
                workbook.Close()
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()

# Example usage for testing
if __name__ == "__main__":
    try:
        processor = OrderProcessor('2026통합발주서_영업_연습.xlsb')
        # Test lookup
        # Based on analysis: B2504240301 -> DUALFIXPRO-티크
        info = processor.lookup_product('B2504240301')
        print(f"Lookup Result: {info}")
    except Exception as e:
        print(e)
