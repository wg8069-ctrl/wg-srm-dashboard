import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

st.set_page_config(page_title="SRM進料x生產排程監控", layout="wide")

@st.cache_data
def load_data():
    # --- 1. 讀取 SRM 進料檔 (ASN 導出檔) ---
    srm_files = glob.glob("訂單資訊*.xls*")
    df_srm = pd.DataFrame()
    if srm_files:
        latest_srm = max(srm_files, key=os.path.getmtime)
        df_srm = pd.read_excel(latest_srm, engine='openpyxl')
        
    # --- 2. 讀取 生產排程資料庫 ---
    # 搜尋包含 "排程資料庫" 字眼的檔案
    plan_files = glob.glob("*排程資料庫*.xls*")
    df_plan = pd.DataFrame()
    if plan_files:
        latest_plan = max(plan_files, key=os.path.getmtime)
        # 讀取排程表，強制讀取前 15 欄確保包含 M 欄
        df_plan_raw = pd.read_excel(latest_plan, engine='openpyxl')
        
        # 根據截圖：K 欄(Index 10)是料號，M 欄(Index 12)是上線日期
        if len(df_plan_raw.columns) >= 13:
            df_plan = df_plan_raw.iloc[:, [10, 12]].copy()
            df_plan.columns = ['料件編號[*]', '生產上線日']
            # 清理日期格式
            df_plan['生產上線日'] = pd.to_datetime(df_plan['生產上線日'], errors='coerce')
            # 同一個料號可能有多個排程，抓最早的一個
            df_plan = df_plan.groupby('料件編號[*]')['生產上線日'].min().reset_index()

    if df_srm.empty:
        return pd.DataFrame()

    # --- 3. 處理進料邏輯 ---
    def calc_delivered(row):
        status = str(row.get('出貨狀態', '')).strip()
        if status in ["已發貨", "全部發貨"]: return row.get('發貨量[*]', 0)
        elif status in ["部分收貨", "全部收貨"]: return row.get('收貨量[*]', 0)
        return 0

    df_srm['已交量'] = df_srm.apply(calc_delivered, axis=1)
    df_srm['發貨日期[*]'] = pd.to_datetime(df_srm['發貨日期[*]'], errors='coerce')

    # --- 4. 合併排程資料 ---
    if not df_plan.empty:
        df_srm = pd.merge(df_srm, df_plan, on='料件編號[*]', how='left')

    # --- 5. 判斷「即時狀態」 (核心警示邏輯) ---
    today = pd.to_datetime(datetime.now().date())
    
    def check_status(row):
        if str(row.get('出貨狀態')) == "全部收貨": return "✅ 已結案"
        
        # 優先判斷是否影響生產
        if pd.notnull(row.get('生產上線日')):
            if row['發貨日期[*]'] > row['生產上線日']:
                return "❌ 警報：晚於上線日"
        
        days_diff = (row['發貨日期[*]'] - today).days
        if days_diff < 0: return "🔴 已逾期"
        if days_diff <= 3: return "🟡 三日內到貨"
        return "🟢 正常待料"

    df_srm['即時狀態'] = df_srm.apply(check_status, axis=1)
    
    # 整理最後顯示欄位
    display_cols = [
        '發貨日期[*]', '生產上線日', '即時狀態', '訂單編號[*]', 
        '供應商名稱[*]', '料件編號[*]', '物料名稱[*]', '規格[*]', 
        '出貨狀態', '發貨量[*]', '收貨量[*]', '已交量'
    ]
    return df_srm[[c for c in display_cols if c in df_srm.columns]]

# --- 網頁介面 ---
st.title("📊 3003 生活館 - 進料 vs 生產排程同步看板")

try:
    df = load_data()
    if not df.empty:
        # 頂部指標
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("❌ 影響生產", len(df[df['即時狀態'] == "❌ 警報：晚於上線日"]))
        c2.metric("🔴 逾期未到", len(df[df['即時狀態'] == "🔴 已逾期"]))
        c3.metric("🟡 三日急件", len(df[df['即時狀態'] == "🟡 三日內到貨"]))
        c4.metric("📦 追蹤總項次", len(df))

        # 搜尋
        q = st.text_input("🔍 搜尋 (供應商/料號/狀態)", "")
        if q:
            mask = df.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)
            df = df[mask]
        
        # 排序：影響生產的放最上面
        df = df.sort_values(by="即時狀態", ascending=False)
        
        # 表格顯示
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.warning("請在 GitHub 上傳『訂單資訊』與『排程資料庫』Excel 檔案。")
except Exception as e:
    st.error(f"程式執行錯誤：{e}")
