import pandas as pd
df = pd.read_excel('mẫu from nhập liệu.xlsx')
# drop all completely empty rows in STT or Tên Công trình
subset = df[['STT', 'Tên Công trình']].dropna(how='all')
with open('items.txt', 'w', encoding='utf-8') as f:
    for idx, row in subset.head(30).iterrows():
        f.write(f"STT: {row['STT']}, TEN: {row['Tên Công trình']}\n")
