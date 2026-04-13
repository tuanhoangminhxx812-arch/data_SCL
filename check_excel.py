import sys
try:
    import pandas as pd
    df = pd.read_excel('mẫu from nhập liệu.xlsx')
    with open('excel_info.txt', 'w', encoding='utf-8') as f:
        f.write("COLUMNS:\n")
        f.write(str(df.columns.tolist()) + "\n\n")
        f.write("FIRST 5 ROWS:\n")
        for record in df.head(5).to_dict('records'):
            f.write(str(record) + "\n")
    print("Successfully wrote to excel_info.txt")
except Exception as e:
    print("Error:", e)
