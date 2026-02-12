import pandas as pd
import os
from datetime import datetime

class OrderProcessor:
    def __init__(self, master_file_path):
        self.master_file_path = master_file_path
        self.products_df = None
        self.addresses_df = None
        self._load_master_data()

    def _load_master_data(self):
        """Loads product and address data from the master xlsb file."""
        if not os.path.exists(self.master_file_path):
            raise FileNotFoundError(f"Master file not found: {self.master_file_path}")

        print("Loading master data...")
        try:
            # Load '코드' sheet for products. Skip top rows to find header if needed.
            # Based on analysis: '품번을 입력하세요:' is in col 0.
            # Headers seems to be on row 4 (0-indexed) based on preview: 
            # Columns: ['품번을 입력하세요:', 'B2504240301', 'DUALFIXPRO-티크', 'Unnamed: 3', 'Unnamed: 4', '품번', '제품명', '사은품 1', '사은품 2', '사은품 3', '사은품 4', '사은품 5', 'Unnamed: 12', 'Unnamed: 13']
            # Actually, looking at analysis_result.txt:
            # Row 4 (index 3? or 4?): "품번", "제품명", "사은품 1"...
            # Let's try reading with header=4 first.
            
            self.products_df = pd.read_excel(self.master_file_path, sheet_name='코드', engine='pyxlsb', header=4)
            # Normalize column names just in case
            self.products_df.columns = [str(c).strip() for c in self.products_df.columns]
            
            # Load '25.12 주소' sheet for addresses
            # Columns: ['Unnamed: 0', 'Unnamed: 1', ... 'Unnamed: 5']
            # Row 0: NaN, "매장명", ..., "주소"
            # It seems header is at row 0 or 1. Let's try header=1 based on preview.
            self.addresses_df = pd.read_excel(self.master_file_path, sheet_name='25.12 주소', engine='pyxlsb', header=1)
            self.addresses_df.columns = [str(c).strip() for c in self.addresses_df.columns]
            
            print("Master data loaded successfully.")
            
        except Exception as e:
            raise Exception(f"Failed to load master data: {e}")

    def lookup_product(self, barcode):
        """Searches for a product by barcode (품번)."""
        if self.products_df is None:
            return None

        # clean input
        barcode = str(barcode).strip()
        
        # '품번' column seems to be the one we need.
        # Let's clean the dataframe '품번' column as well
        # Assuming '품번' column exists. If not, we might need to adjust based on real data.
        
        target_col = '품번'
        if target_col not in self.products_df.columns:
            # Fallback or error
            print(f"Warning: '{target_col}' column not found in products sheet. Available: {self.products_df.columns}")
            return None

        # Filter
        # Convert column to string for comparison
        result = self.products_df[self.products_df[target_col].astype(str).str.strip() == barcode]
        
        if result.empty:
            return None
        
        # Return the first match as a dict
        row = result.iloc[0]
        
        # Construct product info
        product_info = {
            'product_name': row.get('제품명', ''),
            'gifts': []
        }
        
        # Collect gifts
        for i in range(1, 6):
            gift_col = f'사은품 {i}'
            if gift_col in row and pd.notna(row[gift_col]):
                product_info['gifts'].append(row[gift_col])
                
        return product_info

    def generate_order_file(self, order_data_list, output_path):
        """
        Generates an Excel order file from the provided list of order data.
        order_data_list: List of dictionaries containing order details.
        """
        if not order_data_list:
            return False
            
        # Define output columns order
        columns = [
            "거래처명", "주문인", "수취인", 
            "전화번호", "핸드폰", "주소", 
            "바코드", "제품명", "사은품", "수량", 
            "수수료", "배송비", "배송메모"
        ]
        
        # Create DataFrame
        df = pd.DataFrame(order_data_list)
        
        # Ensure all columns exist (fill missing with empty string)
        for col in columns:
            if col not in df.columns:
                df[col] = ""
                
        # Reorder columns
        df = df[columns]
        
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='발주서')
                
                # Auto-adjust column width (basic)
                worksheet = writer.sheets['발주서']
                for idx, col in enumerate(df.columns):
                    # headers are at row 1
                    max_len = len(str(col)) + 2
                    # check content length (sample first 10 rows)
                    for val in df[col].head(10): 
                        if val:
                            max_len = max(max_len, len(str(val)) + 2)
                    
                    col_letter = chr(65 + idx) if idx < 26 else 'A' + chr(65 + (idx - 26)) # Basic A-Z logic, unsafe for > Z columns but enough here
                    # proper utils.get_column_letter is better but trying to keep imports minimal or use openpyxl utils
                    from openpyxl.utils import get_column_letter
                    col_letter = get_column_letter(idx + 1)
                    worksheet.column_dimensions[col_letter].width = min(max_len, 50)

            print(f"Order file created at: {output_path}")
            return True
        except Exception as e:
            print(f"Failed to create order file: {e}")
            return False

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
