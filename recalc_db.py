import pandas as pd
import numpy as np

DB_FILE = 'database_cong_trinh.xlsx'
df = pd.read_excel(DB_FILE)

def recalculate(df_part):
    def set_v(stt, dt, qt):
        mask = df_part['STT'].astype(str).str.strip() == stt
        if mask.any():
            idx = mask.idxmax()
            df_part.at[idx, 'Giá trị Dự toán'] = int(dt)
            df_part.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = int(qt)
            
    def get_v(stt):
        mask = df_part['STT'].astype(str).str.strip() == stt
        if mask.any():
            idx = mask.idxmax()
            dt = df_part.at[idx, 'Giá trị Dự toán']
            qt = df_part.at[idx, 'Giá trị Q.định phê duyệt QT công trình']
            return int(dt) if pd.notna(dt) else 0, int(qt) if pd.notna(qt) else 0
        return 0, 0
        
    def get_sum(stt_list):
        s_dt, s_qt = 0, 0
        for stt in stt_list:
            dt, qt = get_v(stt)
            s_dt += dt
            s_qt += qt
        return s_dt, s_qt

    val_A1_dt, val_A1_qt = get_sum(['A.1.1', 'A.1.2', 'A.1.3', 'A.1.4'])
    set_v('A.1', val_A1_dt, val_A1_qt)
    val_A2_dt, val_A2_qt = get_v('A.2')
    val_A3_dt = int(round((val_A1_dt + val_A2_dt) * 0.10))
    val_A3_qt = int(round((val_A1_qt + val_A2_qt) * 0.10))
    set_v('A.3', val_A3_dt, val_A3_qt)
    val_A_dt = val_A1_dt + val_A2_dt + val_A3_dt
    val_A_qt = val_A1_qt + val_A2_qt + val_A3_qt
    set_v('A', val_A_dt, val_A_qt)

    val_B1_dt, val_B1_qt = get_sum(['B.1.1', 'B.1.2', 'B.1.3', 'B.1.4'])
    set_v('B.1', val_B1_dt, val_B1_qt)
    val_B2_dt, val_B2_qt = get_sum(['B.2.1', 'B.2.2'])
    set_v('B.2', val_B2_dt, val_B2_qt)
    val_B3_dt, val_B3_qt = get_sum(['B.3.1', 'B.3.2'])
    set_v('B.3', val_B3_dt, val_B3_qt)
    val_B4_dt, val_B4_qt = get_v('B.4')
    val_B5_dt, val_B5_qt = get_v('B.5')
    val_B6_dt, val_B6_qt = get_v('B.6')

    val_B7_dt = sum([val_B1_dt, val_B2_dt, val_B3_dt, val_B4_dt, val_B5_dt, val_B6_dt])
    val_B7_qt = sum([val_B1_qt, val_B2_qt, val_B3_qt, val_B4_qt, val_B5_qt, val_B6_qt])
    set_v('B.7', val_B7_dt, val_B7_qt)
    
    val_B8_dt = int(round(val_B7_dt * 0.08))
    val_B8_qt = int(round(val_B7_qt * 0.08))
    set_v('B.8', val_B8_dt, val_B8_qt)
    
    val_B_dt = val_B7_dt + val_B8_dt
    val_B_qt = val_B7_qt + val_B8_qt
    set_v('B', val_B_dt, val_B_qt)

    val_C1_dt, val_C1_qt = get_v('C.1')
    val_C2_dt, val_C2_qt = get_v('C.2')
    val_C3_dt, val_C3_qt = get_v('C.3')
    val_C4_dt, val_C4_qt = get_v('C.4')
    val_C5_dt, val_C5_qt = get_v('C.5')
    sum_C_1_5_dt = val_C1_dt + val_C2_dt + val_C3_dt + val_C4_dt + val_C5_dt
    sum_C_1_5_qt = val_C1_qt + val_C2_qt + val_C3_qt + val_C4_qt + val_C5_qt
    
    val_C6_dt = int(round(sum_C_1_5_dt * 0.08))
    val_C6_qt = int(round(sum_C_1_5_qt * 0.08))
    set_v('C.6', val_C6_dt, val_C6_qt)
    
    val_C_dt = sum_C_1_5_dt + val_C6_dt
    val_C_qt = sum_C_1_5_qt + val_C6_qt
    set_v('C', val_C_dt, val_C_qt)

    val_D_dt, val_D_qt = get_v('D')

    val_E1_dt = val_A1_dt + val_A2_dt + val_B7_dt + sum_C_1_5_dt + val_D_dt
    val_E1_qt = val_A1_qt + val_A2_qt + val_B7_qt + sum_C_1_5_qt + val_D_qt
    set_v('E.1', val_E1_dt, val_E1_qt)
    
    val_E2_dt = val_A3_dt + val_B8_dt + val_C6_dt
    val_E2_qt = val_A3_qt + val_B8_qt + val_C6_qt
    set_v('E.2', val_E2_dt, val_E2_qt)
    
    val_E_dt = val_E1_dt + val_E2_dt
    val_E_qt = val_E1_qt + val_E2_qt
    set_v('E', val_E_dt, val_E_qt)

    val_F_dt, val_F_qt = get_v('F')
    
    val_SCL_dt = val_E1_dt - val_F_dt
    val_SCL_qt = val_E1_qt - val_F_qt
    set_v('SCL', val_SCL_dt, val_SCL_qt)

main_mask = df['Kế hoạch'].notna()
list_cong_trinh = df.loc[main_mask, 'Tên Công trình'].dropna().unique().tolist()

for ct in list_cong_trinh:
    start_indices = df.index[df['Tên Công trình'] == ct].tolist()
    if start_indices:
        start_idx = start_indices[0]
        end_idx = len(df)
        for i in range(start_idx + 1, len(df)):
            val = str(df.at[i, 'STT']).strip().upper()
            if val in ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']:
                end_idx = i
                break
        
        recalculate(df.iloc[start_idx:end_idx])

df.to_excel(DB_FILE, index=False)
print("Recalculate complete.")
