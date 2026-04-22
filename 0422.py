import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

# 設定網頁標題
st.set_page_config(page_title="SRM 進料戰情看板", layout="wide")

@st.cache_data
def load_data():
    # 找到最新的 Excel 檔
    file_pattern = "訂單資訊*.xls*" 
    files = glob.glob(file_pattern)
    if not files:
        return pd.DataFrame()
    
    latest_file = max(files, key=os.path.getmtime)
    
    # 關鍵修改：明確指定 engine='openpyxl'
   df = pd.read_excel(latest_file, engine='openpyxl') 
    return df

    # 4. 【套用 VBA 邏輯】刪除已作廢
    # 對應 VBA: If Trim(wsSrc.Cells(i, "C").Value) = "已作廢" Then wsSrc.Rows(i).Delete
    if '狀態' in df.columns:
        df = df[df['狀態'] != '已作廢']
    
    # 5. 【套用 VBA 邏輯】計算 Q 欄 (加總項目)
    # 對應 VBA: Case "已發貨" -> Q=N(發貨量); Case "全部/部分收貨" -> Q=O(收貨量)
    def calculate_vba_q(row):
        status = str(row['狀態']).strip()
        if status == "已發貨":
            return row.get('發貨量', 0)
        elif status in ["全部收貨", "部分收貨"]:
            return row.get('收貨量', 0)
        else:
            return 0

    df['Q_加總項目'] = df.apply(calculate_vba_q, axis=1)

    # 6. 基礎看板運算
    df['預計交貨日期'] = pd.to_datetime(df['預計交貨日期'])
    today = pd.to_datetime(datetime.now().date())
    df['剩餘天數'] = (df['預計交貨日期'] - today).dt.days
    
    return df

# --- 網頁呈現介面 ---
st.title("📊 SRM 進料自動化監控 (VBA 整合版)")

try:
    df = load_data()
    
    if not df.empty:
        # 顯示統計指標 (包含 VBA 邏輯計算出的 Q 總數)
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 逾期未交", len(df[df['剩餘天數'] < 0]))
        c2.metric("📦 Q欄總計 (已發/收)", f"{df['Q_加總項目'].sum():,.0f}")
        c3.metric("📑 總訂單筆數", len(df))

        # 全物料查詢搜尋框
        search_query = st.text_input("🔍 全物料關鍵字查詢 (輸入單號、物料或供應商)", "")
        
        if search_query:
            mask = df.apply(lambda r: r.astype(str).str.contains(search_query, case=False).any(), axis=1)
            display_df = df[mask]
        else:
            display_df = df

        # 顯示資料表格
        st.dataframe(display_df, use_container_width=True, hide_index=True)
except Exception as e:
    st.error(f"程式執行發生錯誤：{e}")
