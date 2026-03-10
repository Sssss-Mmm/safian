import pandas as pd
df = pd.read_excel('2026통합발주서_영업_연습.xlsb', sheet_name='Sheet1 (2)', engine='pyxlsb', header=None, nrows=20)
print(df.to_string())
