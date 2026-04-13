import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1
content = content.replace(
    "sub_items_columns = ['STT', 'Tên Hạng mục', 'Giá trị Dự toán', 'Giá trị Q.định phê duyệt QT công trình', 'Ghi chú']",
    "sub_items_columns = ['STT', 'Tên Hạng mục', 'Giá trị Dự toán', 'Giá trị quyết toán', 'Chênh lệch']"
)

# 2
target2 = """                sub_rows = sub_rows[['STT', 'Tên Công trình', 'Giá trị Dự toán', 'Giá trị Q.định phê duyệt QT công trình', 'Ghi chú']]
                sub_rows = sub_rows.rename(columns={'Tên Công trình': 'Tên Hạng mục'})
                sub_rows['Giá trị Dự toán'] = sub_rows['Giá trị Dự toán'].fillna(0).astype(int)
                sub_rows['Giá trị Q.định phê duyệt QT công trình'] = sub_rows['Giá trị Q.định phê duyệt QT công trình'].fillna(0).astype(int)
                for col in ['STT', 'Tên Hạng mục', 'Ghi chú']:"""
repl2 = """                sub_rows = sub_rows[['STT', 'Tên Công trình', 'Giá trị Dự toán', 'Giá trị Q.định phê duyệt QT công trình', 'Ghi chú']]
                sub_rows = sub_rows.rename(columns={'Tên Công trình': 'Tên Hạng mục', 'Giá trị Q.định phê duyệt QT công trình': 'Giá trị quyết toán', 'Ghi chú': 'Chênh lệch'})
                sub_rows['Giá trị Dự toán'] = sub_rows['Giá trị Dự toán'].fillna(0).astype(int)
                sub_rows['Giá trị quyết toán'] = sub_rows['Giá trị quyết toán'].fillna(0).astype(int)
                for col in ['STT', 'Tên Hạng mục', 'Chênh lệch']:"""
content = content.replace(target2, repl2)

# 3
target3 = """"Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""}"""
repl3 = """"Giá trị quyết toán": 0, "Chênh lệch": 0}"""
content = content.replace(target3, repl3)

# 4
target4 = """    col_config = {
        "Giá trị Dự toán": st.column_config.NumberColumn(format="%,d", step=1),
        "Giá trị Q.định phê duyệt QT công trình": st.column_config.NumberColumn(format="%,d", step=1)
    }"""
repl4 = """    col_config = {
        "Giá trị Dự toán": st.column_config.NumberColumn(format="%,d", step=1),
        "Giá trị quyết toán": st.column_config.NumberColumn(format="%,d", step=1),
        "Chênh lệch": st.column_config.NumberColumn(format="%,d", disabled=True)
    }"""
content = content.replace(target4, repl4)

# 5
target5 = """    # Ép kiểu an toàn để tính toán
    for col in ['Giá trị Dự toán', 'Giá trị Q.định phê duyệt QT công trình']:
        calculated_df[col] = pd.to_numeric(calculated_df[col], errors='coerce').fillna(0).astype(int)

    changed = False"""
repl5 = """    # Ép kiểu an toàn để tính toán
    for col in ['Giá trị Dự toán', 'Giá trị quyết toán']:
        calculated_df[col] = pd.to_numeric(calculated_df[col], errors='coerce').fillna(0).astype(int)

    calculated_df['Chênh lệch'] = calculated_df['Giá trị Dự toán'] - calculated_df['Giá trị quyết toán']

    changed = False"""
content = content.replace(target5, repl5)

# 6
target6 = """            if int(calculated_df.at[idx, 'Giá trị Q.định phê duyệt QT công trình']) != int(val_qt):
                calculated_df.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = int(val_qt)
                changed = True"""
repl6 = """            if int(calculated_df.at[idx, 'Giá trị quyết toán']) != int(val_qt):
                calculated_df.at[idx, 'Giá trị quyết toán'] = int(val_qt)
                changed = True"""
content = content.replace(target6, repl6)

# 7
target7 = """        if mask.any():
            idx = mask.idxmax()
            return int(calculated_df.at[idx, 'Giá trị Dự toán']), int(calculated_df.at[idx, 'Giá trị Q.định phê duyệt QT công trình'])
        return 0, 0"""
repl7 = """        if mask.any():
            idx = mask.idxmax()
            return int(calculated_df.at[idx, 'Giá trị Dự toán']), int(calculated_df.at[idx, 'Giá trị quyết toán'])
        return 0, 0"""
content = content.replace(target7, repl7)

# 8
target8 = """                        sub_row['Giá trị Q.định phê duyệt QT công trình'] = row.get('Giá trị Q.định phê duyệt QT công trình')
                        sub_row['Ghi chú'] = row.get('Ghi chú')"""
repl8 = """                        sub_row['Giá trị Q.định phê duyệt QT công trình'] = row.get('Giá trị quyết toán')
                        sub_row['Ghi chú'] = row.get('Chênh lệch')"""
content = content.replace(target8, repl8)

# 9
target9 = """                        ws.title = safe_sheet_name
                        ws['C8'] = ct # Luôn lấy Tên công trình"""
repl9 = """                        ws.title = safe_sheet_name
                        ws['D8'] = f"Công trình: {ct}" # Luôn lấy Tên công trình"""
content = content.replace(target9, repl9)

# 10
target10 = """                    ws = wb.active
                    ws.title = 'Báo cáo'
                    ws['C8'] = selected_ct"""
repl10 = """                    ws = wb.active
                    ws.title = 'Báo cáo'
                    ws['D8'] = f"Công trình: {selected_ct}" """
content = content.replace(target10, repl10)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
