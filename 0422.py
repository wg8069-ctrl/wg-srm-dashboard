import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

# 設定網頁標題
st.set_page_config(page_title="SRM 進料戰情看板", layout="wide")

@st.cache_data
def load_data():
    # 1. 自動偵測檔案
    file_pattern = "訂單資訊*.xls*" 
    files = glob.glob(file_pattern)
    
    if not files:
        return pd.DataFrame()

    # 2. 抓取最新檔案
    latest_file = max(files, key=os.path.getmtime)
    
    # 3. 讀取資料 (這裡就是原本報錯的地方，現在已校正縮進)
    df = pd.read_excel(latest_file, engine='openpyxl')

    # 4. 套用 VBA 邏輯：刪除已作廢
    if '狀態' in df.columns:
        df = df[df['狀態'] != '已作廢']
    
    # 5. 計算 Q 欄
    def calculate_vba_q(row):
        status = str(row.get('狀態', '')).strip()
        if status == "已發貨":
            return row.get('發貨量', 0)
        elif status in ["全部收貨", "部分收貨"]:
            return row.get('收貨量', 0)
        else:
            return 0

    df['Q_加總項目'] = df.apply(calculate_vba_q, axis=1)

    # 6. 日期運算
    df['預計交貨日期'] = pd.to_datetime(df['預計交貨日期'])
    today = pd.to_datetime(datetime.now().date())
    df['剩餘天數'] = (df['預計交貨日期'] - today).dt.days
    
    return df

# --- 網頁呈現介面 ---
st.title("📊 SRM 進料自動化監控 (VBA 整合版)")

try:
    df = load_data()
    
    if not df.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 逾期未交", len(df[df['剩餘天數'] < 0]))
        c2.metric("📦 Q欄總計", f"{df['Q_加總項目'].sum():,.0f}")
        c3.metric("📑 總訂單筆數", len(df))

        search_query = st.text_input("🔍 全物料關鍵字查詢 (單號/供應商/物料)", "")
        
        if search_query:
            mask = df.apply(lambda r: r.astype(str).str.contains(search_query, case=False).any(), axis=1)
            display_df = df[mask]
        else:
            display_df = df

        st.dataframe(display_df, use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ 資料夾中找不到『訂單資訊』開頭的 Excel 檔案，請上傳檔案至 GitHub。")

except Exception as e:
    st.error(f"程式執行發生錯誤：{e}")
