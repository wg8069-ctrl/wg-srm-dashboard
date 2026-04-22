import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

# 1. 網頁基本設定
st.set_page_config(page_title="SRM 進料戰情看板", layout="wide")

@st.cache_data
def load_data():
    # 自動搜尋 GitHub 資料夾內最新的 Excel
    files = glob.glob("訂單資訊*.xls*")
    if not files:
        return pd.DataFrame()
    
    latest_file = max(files, key=os.path.getmtime)
    
    # 讀取資料並套用 VBA 邏輯
    df = pd.read_excel(latest_file, engine='openpyxl')
    
    # 刪除已作廢
    if '狀態' in df.columns:
        df = df[df['狀態'] != '已作廢']
    
    # 計算 Q 欄
    def calculate_vba_q(row):
        status = str(row.get('狀態', '')).strip()
        if status == "已發貨":
            return row.get('發貨量', 0)
        elif status in ["全部收貨", "部分收貨"]:
            return row.get('收貨量', 0)
        return 0

    df['Q_加總項目'] = df.apply(calculate_vba_q, axis=1)
    
    # 日期處理
    df['預計交貨日期'] = pd.to_datetime(df['預計交貨日期'])
    today = pd.to_datetime(datetime.now().date())
    df['剩餘天數'] = (df['預計交貨日期'] - today).dt.days
    
    return df

# 2. 介面呈現
st.title("📊 3003 生活館 - 進料監控 (外網版)")

try:
    df = load_data()
    if not df.empty:
        # 儀表板數據
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 逾期件數", len(df[df['剩餘天數'] < 0]))
        c2.metric("📦 Q欄總計", f"{df['Q_加總項目'].sum():,.0f}")
        c3.metric("📑 總訂單筆數", len(df))

        # 搜尋功能
        search_q = st.text_input("🔍 全物料搜尋 (輸入供應商或料號)", "")
        if search_q:
            mask = df.apply(lambda r: r.astype(str).str.contains(search_q, case=False).any(), axis=1)
            df = df[mask]
        
        # 顯示表格
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.warning("請上傳『訂單資訊』開頭的 Excel 檔案到 GitHub。")
except Exception as e:
    st.error(f"程式執行出錯：{e}")
