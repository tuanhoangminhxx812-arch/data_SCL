import pandas as pd
import shutil
import os

DB_FILE = 'database_cong_trinh.xlsx'
if not os.path.exists(DB_FILE):
    print("No DB found")
    exit()

shutil.copy(DB_FILE, 'database_cong_trinh_backup.xlsx')

df = pd.read_excel(DB_FILE)

# Cấu trúc Mẫu 04
sub_items_columns = ['STT', 'Tên Hạng mục', 'Giá trị Dự toán', 'Giá trị Q.định phê duyệt QT công trình', 'Ghi chú']
initial_data = [
    {"STT": "A", "Tên Hạng mục": "CHI PHÍ VẬT TƯ, THIẾT BỊ (sau thuế)", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "A.1", "Tên Hạng mục": "Chi phí thiết bị", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "A.1.1", "Tên Hạng mục": "Thiết bị nhập khẩu", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "A.1.2", "Tên Hạng mục": "VT A cấp", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "A.1.3", "Tên Hạng mục": "Chi phí tháo dỡ, lắp đặt", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "A.1.4", "Tên Hạng mục": "Chi phí thí nghiệm, hiệu chỉnh", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "A.2", "Tên Hạng mục": "Chi phí vật tư", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "A.3", "Tên Hạng mục": "Thuế GTGT", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B", "Tên Hạng mục": "CHI PHÍ SỬA CHỮA", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.1", "Tên Hạng mục": "Chi phí vật liệu", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.1.1", "Tên Hạng mục": "Vật liệu phần không áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.1.2", "Tên Hạng mục": "Vật liệu phần áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.1.3", "Tên Hạng mục": "Chênh lệch giá vật liệu phần áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.1.4", "Tên Hạng mục": "Vật liệu phụ trong SCL thiết bị", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.2", "Tên Hạng mục": "Chi phí nhân công", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.2.1", "Tên Hạng mục": "Chi phí nhân công phần không áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.2.2", "Tên Hạng mục": "Chi phí nhân công phần áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.3", "Tên Hạng mục": "Chi phí máy thi công", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.3.1", "Tên Hạng mục": "Chi phí máy thi công phần không áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.3.2", "Tên Hạng mục": "Chi phí máy thi công phần áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.4", "Tên Hạng mục": "Chi phí làm đêm, làm thêm giờ", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.5", "Tên Hạng mục": "Chi phí chung", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.6", "Tên Hạng mục": "Thu nhập chịu thuế tính trước", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.7", "Tên Hạng mục": "Giá trị sửa chữa trước thuế", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "B.8", "Tên Hạng mục": "Thuế GTGT", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "C", "Tên Hạng mục": "CHI PHÍ KHÁC (sau thuế)", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "C.1", "Tên Hạng mục": "Chi phí giám sát thi công xây dựng", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "C.2", "Tên Hạng mục": "Chi phí giám sát lắp đặt thiết bị", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "C.3", "Tên Hạng mục": "Chi phí bảo hiểm công trình", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "C.4", "Tên Hạng mục": "Chi phí thẩm tra - phê duyệt quyết toán", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "C.5", "Tên Hạng mục": "Vận chuyển VTTB A cấp đến công trường", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "C.6", "Tên Hạng mục": "Thuế GTGT", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "D", "Tên Hạng mục": "CHI PHÍ DỰ PHÒNG", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "E", "Tên Hạng mục": "Tổng giá trị sau thuế", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "E.1", "Tên Hạng mục": "Tổng giá trị trước thuế", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "E.2", "Tên Hạng mục": "Thuế GTGT", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "F", "Tên Hạng mục": "GIÁ TRỊ VẬT TƯ THU HỒI", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""},
    {"STT": "SCL", "Tên Hạng mục": "CHI PHÍ SCL", "Giá trị Dự toán": 0, "Giá trị Q.định phê duyệt QT công trình": 0, "Ghi chú": ""}
]

new_rows = []
skip_until_next_roman = False

for index, row in df.iterrows():
    stt = str(row['STT']).strip().upper()
    if stt in ['I', 'II', 'III', 'IV', 'V']:
        skip_until_next_roman = False

        # Check if children use old format
        children = []
        for i in range(index + 1, len(df)):
            c_stt = str(df.at[i, 'STT']).strip().upper()
            if c_stt in ['I', 'II', 'III', 'IV', 'V', 'VI']:
                break
            children.append(df.iloc[i])
        
        c_stts = [str(c['STT']).strip() for c in children]
        is_old = any(c in c_stts for c in ['1', '1.1', '2', '3']) and not any(c in c_stts for c in ['A.1', 'B.1.1'])
        
        new_rows.append(row.to_dict())
        
        if is_old:
            print(f"Converting Project {stt}")
            skip_until_next_roman = True
            
            mapped_vals = {'A.1.2': 0, 'A.1.4': 0, 'A.2': 0, 'B.1.1': 0, 'C.1': 0, 'D': 0, 'F': 0}
            mapped_vals_qt = {'A.1.2': 0, 'A.1.4': 0, 'A.2': 0, 'B.1.1': 0, 'C.1': 0, 'D': 0, 'F': 0}
            
            for c in children:
                c_stt = str(c['STT']).strip()
                dt = c['Giá trị Dự toán'] if pd.notna(c['Giá trị Dự toán']) else 0
                qt = c['Giá trị Q.định phê duyệt QT công trình'] if pd.notna(c['Giá trị Q.định phê duyệt QT công trình']) else 0
                
                if c_stt == '1.1': 
                    mapped_vals['A.1.2'] = dt; mapped_vals_qt['A.1.2'] = qt
                elif c_stt == '1.2':
                    mapped_vals['A.1.4'] = dt; mapped_vals_qt['A.1.4'] = qt
                elif c_stt == '2':
                    mapped_vals['A.2'] = dt; mapped_vals_qt['A.2'] = qt
                elif c_stt == '3':
                    mapped_vals['B.1.1'] = dt; mapped_vals_qt['B.1.1'] = qt
                elif c_stt == '4':
                    mapped_vals['C.1'] = dt; mapped_vals_qt['C.1'] = qt
                elif c_stt == '5':
                    mapped_vals['D'] = dt; mapped_vals_qt['D'] = qt
                elif c_stt == '6':
                    mapped_vals['F'] = dt; mapped_vals_qt['F'] = qt
                    
            import copy
            for t_row in initial_data:
                nr = copy.deepcopy(row.to_dict())
                for k in nr.keys():
                    if k not in ['STT', 'Tên Công trình', 'Mã CT', 'Kế hoạch', 'Số Phương án', 'Ngày Phương án', 'Giá trị Phương án', 'Số Dự toán', 'Ngày Dự toán', 'Giá trị Dự toán', 'Số Hợp đồng thiết kế', 'Ngày Hợp đồng thiết kế', 'Giá trị Hợp đồng thiết kế', 'Số Hợp đồng giám sát', 'Ngày Hợp đồng giám sát', 'Giá trị Hợp đồng giám sát', 'Số Hợp đồng xây lắp', 'Ngày Hợp đồng xây lắp', 'Giá trị Hợp đồng xây lắp', 'Giá trị phát sinh', 'Giá trị VT thừa', 'Giá trị VTTH', 'Số Q.định phê duyệt QT công trình', 'Ngày Q.định phê duyệt QT công trình', 'Giá trị Q.định phê duyệt QT công trình', 'Số tiền bằng chữ', 'Ghi chú', 'Đơn vị QL']: 
                        continue # just safe key check
                    nr[k] = None
                nr['STT'] = t_row['STT']
                nr['Tên Công trình'] = t_row['Tên Hạng mục']
                if t_row['STT'] in mapped_vals:
                    nr['Giá trị Dự toán'] = mapped_vals[t_row['STT']]
                    nr['Giá trị Q.định phê duyệt QT công trình'] = mapped_vals_qt[t_row['STT']]
                else:
                    nr['Giá trị Dự toán'] = 0
                    nr['Giá trị Q.định phê duyệt QT công trình'] = 0
                new_rows.append(nr)
                
        else:
            skip_until_next_roman = False
            
    elif not skip_until_next_roman:
        new_rows.append(row.to_dict())

new_df = pd.DataFrame(new_rows)
new_df.to_excel(DB_FILE, index=False)
print("Transformation Complete")
