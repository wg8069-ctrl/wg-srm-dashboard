import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

st.set_page_config(page_title="SRM進料x生產排程監控", layout="wide")

# 偵錯功能：列出目前目錄所有檔案
def list_files():
    st.sidebar.write("### 📂 目前系統偵測到的檔案：")
    all_files = os.listdir(".")
    for f in all_files:
        if "xlsx" in f or "xls" in f:
            st.sidebar.text(f"✅ {f}")
    return all_files

@st.cache_data
def load_data():
    # 1. 搜尋檔案
    srm_files = glob.glob("*訂單資訊*.xls*")
    plan_files = glob.glob("*排程資料庫*.xls*")
    
    # 如果找不到，拋出更詳細的錯誤
    if not srm_files or not plan_files:
        error_msg = ""
        if not srm_files: error_msg += "【找不到訂單資訊檔案】 "
        if not plan_files: error_msg += "【找不到排程資料庫檔案】"
        st.error(f"❌ 檔案讀取失敗：{error_msg}")
        return pd.DataFrame()

    # 2. 讀取最新檔案
    latest_srm = max(srm_files, key=os.path.getmtime)
    latest_plan = max(plan_files, key=os.path.getmtime)

    # 3. 處理排程資料 (基準)
    df_plan_raw = pd.read_excel(latest_plan, engine='openpyxl')
    if len(df_plan_raw.columns) < 13:
        st.error(f"❌ 排程檔案格式不對，總欄位數只有 {len(df_plan_raw.columns)} 欄")
        return pd.DataFrame()

    df_main = df_plan_raw.iloc[:, [10, 6, 12]].copy()
    df_main.columns = ['料件編號[*]', '訂單量', '生產上線日']
    df_main['生產上線日'] = pd.to_datetime(df_main['生產上線日'], errors='coerce')
    df_main = df_main.groupby('料件編號[*]').agg({'訂單量': 'sum', '生產上線日': 'min'}).reset_index()

    # 4. 處理進料資料
    df_srm = pd.read_excel(latest_srm, engine='openpyxl')
    
    def calc_delivered(row):
        status = str(row.get('出貨狀態', '')).strip()
        send_val = row.get('發貨量[*]', 0)
        recv_val = row.get('收貨量[*]', 0)
        if status in ["已發貨", "全部發貨"]: return send_val
        elif status in ["部分收貨", "全部收貨"]: return recv_val
        return 0
    
    df_srm['計算後已交'] = df_srm.apply(calc_delivered, axis=1)
    df_srm_agg = df_srm.groupby('料件編號[*]').agg({
        '訂單編號[*]': 'first',
        '供應商名稱[*]': 'first',
        '物料名稱[*]': 'first',
        '規格[*]': 'first',
        '計算後已交': 'sum',
        '發貨日期[*]': 'max'
    }).reset_index()
    df_srm_agg.rename(columns={'發貨日期[*]': '完工日'}, inplace=True)

    # 5. 合併與計算
    df = pd.merge(df_main, df_srm_agg, on='料件編號[*]', how='left')
    df['已交量'] = df['計算後已交'].fillna(0)
    df['未交數量'] = df['訂單量'] - df['已交量']
    
    today = pd.to_datetime(datetime.now().date())
    def check_status(row):
        if row['未交數量'] <= 0: return "✅ 已結案"
        if pd.notnull(row.get('生產上線日')) and pd.notnull(row.get('完工日')):
            if row['完工日'] > row['生產上線日']: return "❌ 警報：晚於生產日"
        days_diff = (pd.to_datetime(row.get('完工日')) - today).days if pd.notnull(row.get('完工日')) else 999
        if days_diff < 0: return "🔴 已逾期"
        return "🟢 正常待料"

    df['即時狀態'] = df.apply(check_status, axis=1)
    return df[['即時狀態', '生產上線日', '完工日', '料件編號[*]', '物料名稱[*]', '規格[*]', '訂單量', '已交量', '未交數量']]

# --- 介面 ---
st.title("📊 3003 生活館 - 進料監控診斷模式")
list_files() # 在側邊欄列出檔案供核對

try:
    df = load_data()
    if not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True)
except Exception as e:
    st.error(f"程式運行中斷：{e}")
