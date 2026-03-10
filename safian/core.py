import pandas as pd
import os
import platform
import random
import string
from datetime import datetime

class OrderProcessor:
    def __init__(self, master_file_path):
        self.master_file_path = master_file_path
        self.products_df = None
        
        # Load Code map immediately upon initialization
        self._load_products()

    def _log(self, msg):
        """디버깅용 로그 저장"""
        try:
            with open("debug.log", "a", encoding="utf-8") as f:
                f.write(f"{datetime.now()}: {msg}\n")
        except:
             pass

    def _load_products(self):
        """마스터 엑셀에서 제품명 <-> 바코드 매핑 데이터 로드"""
        if not os.path.exists(self.master_file_path):
            self._log(f"엑셀 파일을 찾을 수 없습니다: {self.master_file_path}")
            return
            
        try:
            # pyxlsb를 이용하여 '코드' 시트 읽기 (발주서 연습 파일 기준)
            df = pd.read_excel(self.master_file_path, sheet_name='코드', engine='pyxlsb', header=0)
            
            # 컬럼명 정제 (공백 등 제거)
            df.columns = [str(c).strip() for c in df.columns]
            
            # None/NaN 데이터 제거 (품번, 제품명이 있는 행만)
            if '품번' in df.columns and '제품명' in df.columns:
                 self.products_df = df[df['품번'].notna() & df['제품명'].notna()].copy()
                 self.products_df['품번'] = self.products_df['품번'].astype(str).str.strip()
                 self.products_df['제품명'] = self.products_df['제품명'].astype(str).str.strip()
                 self._log("상품/바코드 목록 로드 성공 완료")
            else:
                 self._log(f"오류: '코드' 시트에 '품번'이나 '제품명' 열이 없습니다. (현재열: {df.columns.tolist()})")
                 
        except Exception as e:
             self._log(f"상품 목록 로드 실패: {e}")

    def find_barcode_by_product_name(self, hint):
        """
        고객이 쓴 상품명(힌트)을 바탕으로 바코드를 찾아냅니다. (본품과 사은품 모두)
        예: hint="나주배" -> 바코드: A001, 제품명: 명품 나주배, 사은품: 배즙
        """
        result = []
        
        if self.products_df is None or not hint:
             return result
             
        # 공백 제거하여 매칭 확률 증가
        hint_clean = hint.replace(" ", "")
        
        best_match_row = None
        
        # 1. '코드' 시트에서 제품명에 힌트가 포함되어 있거나 힌트에 제품명이 포함되어 있는지 검색
        for idx, row in self.products_df.iterrows():
            prod_name = row['제품명']
            prod_name_clean = prod_name.replace(" ", "")
            
            # 둘 중 하나라도 포함관계면 매칭 성공
            if prod_name_clean in hint_clean or hint_clean in prod_name_clean:
                best_match_row = row
                break # 첫 번째 일치 항목 반환
                
        if best_match_row is not None:
             barcode = best_match_row['품번']
             
             # 본품 추가
             result.append({
                 'type': '본품',
                 'product_name': best_match_row['제품명'],
                 'barcode': barcode
             })
             
             # 사은품 추가 (코드 시트에 사은품1 ~ 5 까지 있다고 가정)
             for i in range(1, 6):
                  gift_col = f'사은품 {i}'
                  if gift_col in best_match_row and pd.notna(best_match_row[gift_col]):
                       gift_barcode = str(best_match_row[gift_col]).strip()
                       if gift_barcode and gift_barcode != 'nan':
                            # 사은품 바코드로 제품명 조회
                            gift_match = self.products_df[self.products_df['품번'] == gift_barcode]
                            gift_name = gift_match.iloc[0]['제품명'] if not gift_match.empty else ""
                            result.append({
                                'type': '사은품',
                                'product_name': gift_name,
                                'barcode': gift_barcode
                            })
                            
             self._log(f"힌트 '{hint}' -> 바코드 '{barcode}' 매칭 성공")
                            
        return result

    def lookup_product_by_barcode(self, barcode):
         """수기로 바코드를 쳤을때 (기존 로직)"""
         result = []
         if self.products_df is None or not barcode: return result
         
         match = self.products_df[self.products_df['품번'] == barcode]
         if not match.empty:
             row = match.iloc[0]
             result.append({'type': '본품', 'product_name': row['제품명'], 'barcode': barcode})
             
             for i in range(1, 6):
                  gift_col = f'사은품 {i}'
                  if gift_col in row and pd.notna(row[gift_col]):
                       gift_barcode = str(row[gift_col]).strip()
                       if gift_barcode and gift_barcode != 'nan':
                            gm = self.products_df[self.products_df['품번'] == gift_barcode]
                            g_name = gm.iloc[0]['제품명'] if not gm.empty else ""
                            result.append({'type': '사은품', 'product_name': g_name, 'barcode': gift_barcode})
         return result

    def append_orders_to_excel(self, order_list):
        """
        윈도우 환경에서만 동작 O
        기존 엑셀파일 밑에 레코드 통째로 붙여넣어줍니다.
        임시 파일이 아닌 COM 객체를 직접 핸들링하며, 파일이 이미 열려있는 경우를 방어합니다.
        """
        if not order_list:
             return False, "저장할 데이터가 없습니다."
             
        if platform.system() != 'Windows':
             return False, "엑셀 자동 저장은 윈도우 환경에서만 지원됩니다."
             
        try:
            import win32com.client
            import pythoncom
            import os
        except ImportError:
            return False, "pywin32 패키지가 필요합니다 (pip install pywin32)"
            
        excel = None
        workbook = None
        was_open_by_user = False
        
        try:
             pythoncom.CoInitialize()
             
             # 절대 경로 변환 및 wsl 등 이상 경로 보정
             abs_path = os.path.abspath(self.master_file_path)
             if abs_path.startswith("\\\\wsl.localhost") or abs_path.startswith("\\\\wsl$"):
                 pass
             elif abs_path.startswith("\\wsl"):
                 abs_path = "\\" + abs_path
                 
             if not os.path.exists(abs_path):
                 return False, f"엑셀 파일 접근 불가: {abs_path}"

             # 열려있는 엑셀 인스턴스가 있는지 먼저 확인 (없으면 Dispatch로 새로 생성)
             try:
                 excel = win32com.client.GetActiveObject("Excel.Application")
             except:
                 excel = win32com.client.Dispatch("Excel.Application")
                 excel.DisplayAlerts = False
                 excel.Visible = False

             # 이미 열려있는 동일 파일이 있는지 확인
             for wb in excel.Workbooks:
                 if wb.FullName.lower() == abs_path.lower():
                     workbook = wb
                     was_open_by_user = True
                     break
                     
             if not workbook:
                 try:
                     # UpdateLinks=0, ReadOnly=False 강제
                     workbook = excel.Workbooks.Open(abs_path, 0, False)
                 except Exception as open_err:
                     err_msg = str(open_err)
                     if hasattr(open_err, 'excepinfo') and open_err.excepinfo:
                         err_msg += f"\n상세: {open_err.excepinfo}"
                     return False, f"엑셀 파일을 여는 중 오류 발생 (다른 프로그램에서 사용중일 수 있습니다):\n{err_msg}"

             if workbook.ReadOnly and not was_open_by_user:
                 # ReadOnly로 열린 경우 (다른 사람이 열고 있거나 권한 문제)
                 workbook.Close(SaveChanges=False)
                 return False, "엑셀 파일이 읽기 전용 상태입니다. 편집 중인 엑셀 창을 모두 닫고 다시 시도해주세요."
                 
             # 발주내역 시트를 찾거나 생성
             sheet_name = '발주내역'
             try:
                 sheet = workbook.Sheets(sheet_name)
             except:
                 sheet = workbook.Sheets.Add(After=workbook.Sheets(workbook.Sheets.Count))
                 sheet.Name = sheet_name
                 # 헤더
                 headers = ["거래처명", "주문번호", "주문인", "수취인", "전화번호", "핸드폰", "우편번호", "주소", "바코드", "제품명", "사은품", "수량", "수수료", "배송비", "배송메모"]
                 for c, h in enumerate(headers):
                      sheet.Cells(1, c + 1).Value = h
                      
             # 마지막 행 찾기 (xlUp의 값은 -4162)
             last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
             if last_row == 1 and sheet.Cells(1, 1).Value is None:
                 last_row = 0
                 
             start_row = last_row + 1
             cols = ["거래처명", "주문번호", "주문인", "수취인", "전화번호", "핸드폰", "우편번호", "주소", "바코드", "제품명", "사은품", "수량", "수수료", "배송비", "배송메모"]
             
             # 데이터 채우기
             for i, order in enumerate(order_list):
                  for c, key in enumerate(cols):
                       # 빈 문자열 대신 nan이나 None이 들어가면 에러 나므로 확인
                       val = order.get(key, "")
                       sheet.Cells(start_row + i, c + 1).Value = "" if val is None else str(val)
                       
             # 저장 및 정리
             if was_open_by_user:
                 # 사용자가 이미 열어두고 있었으면 덮어쓰기 저장만 하고 닫지는 않음
                 workbook.Save()
                 workbook = None  # finally 구문에서 close 시키지 않기 위함
             else:
                 # 백그라운드에서 우리가 열었으면 저장 후 닫기
                 workbook.Close(SaveChanges=True)
                 workbook = None

             return True, "엑셀(발주내역 시트)에 자동 저장을 완료했습니다."
             
        except Exception as e:
             return False, f"엑셀 저장 오류: {e}"
        finally:
             if workbook and not was_open_by_user: 
                 try: workbook.Close(SaveChanges=False)
                 except: pass
                 
             if excel and not was_open_by_user:
                 # 활성화된 워크북 갯수 점검 후 종료
                 try:
                     if excel.Workbooks.Count == 0:
                         excel.Quit()
                 except: pass
                 
             try: pythoncom.CoUninitialize()
             except: pass
