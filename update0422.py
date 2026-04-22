import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

st.set_page_config(page_title="SRM進料x生產排程監控", layout="wide")

@st.cache_data
def load_data():
    # --- 1. 讀取 生產排程資料庫 (作為基準檔) ---
    plan_files = glob.glob("*排程資料庫*.xls*")
    df_main = pd.DataFrame()
    if plan_files:
        latest_plan = max(plan_files, key=os.path.getmtime)
        df_plan_raw = pd.read_excel(latest_plan, engine='openpyxl')
        
        # 根據截圖與需求：
        # G 欄 (Index 6) -> 訂單量
        # K 欄 (Index 10) -> 料件編號
        # M 欄 (Index 12) -> 上線日期
        if len(df_plan_raw.columns) >= 13:
            df_main = df_plan_raw.iloc[:, [10, 6, 12]].copy()
            df_main.columns = ['料件編號[*]', '訂單量', '生產上線日']
            df_main['生產上線日'] = pd.to_datetime(df_main['生產上線日'], errors='coerce')
            # 同料號可能有不同單據，加總訂單量並取最早日期
            df_main = df_main.groupby('料件編號[*]').agg({'訂單量': 'sum', '生產上線日': 'min'}).reset_index()

    # --- 2. 讀取 SRM 進料檔 (ASN 導出檔) ---
    srm_files = glob.glob("訂單資訊*.xls*")
    df_srm = pd.DataFrame()
    if srm_files:
        latest_srm = max(srm_files, key=os.path.getmtime)
        df_srm = pd.read_excel(latest_srm, engine='openpyxl')
        
        # 計算已交量邏輯
        def calc_delivered(row):
            status = str(row.get('出貨狀態', '')).strip()
            if status in ["已發貨", "全部發貨"]: return row.get('發貨量[*]', 0)
            elif status in ["部分收貨", "全部收貨"]: return row.get('收貨量[*]', 0)
            return 0
        
        df_srm['已交量'] = df_srm.apply(calc_delivered, axis=1)
        
        # 彙總進料檔資訊 (以料號為單位，合併單號、備註、已交量、發貨日期)
        # 注意：發貨日期改名為「完工日」供比對
        df_srm_agg = df_srm.groupby('料件編號[*]').agg({
            '訂單編號[*]': 'first',
            '供應商名稱[*]': 'first',
            '物料名稱[*]': 'first',
            '規格[*]': 'first',
            '已交量': 'sum',
            '發貨日期[*]': 'max',
            '單身備註': 'first',
            '出貨狀態': 'first'
        }).reset_index()
        df_srm_agg.rename(columns={'發貨日期[*]': '完工日'}, inplace=True)

    if df_main.empty:
        return pd.DataFrame()

    # --- 3. 合併資料 (以排程檔為核心) ---
    df = pd.merge(df_main, df_srm_agg, on='料件編號[*]', how='left')

    # --- 4. 計算未交數量與狀態 ---
    df['未交數量'] = df['訂單量'] - df['已交量'].fillna(0)
    
    today = pd.to_datetime(datetime.now().date())
    def check_status(row):
        if row['未交數量'] <= 0: return "✅ 已結案"
        
        # 比對完工日與生產上線日
        if pd.notnull(row.get('生產上線日')) and pd.notnull(row.get('完工日')):
            if row['完工日'] > row['生產上線日']:
                return "❌ 警報：晚於生產日"
        
        days_diff = (pd.to_datetime(row.get('完工日')) - today).days if pd.notnull(row.get('完工日')) else 999
        if days_diff < 0: return "🔴 已逾期"
        return "🟢 正常待料"

    df['即時狀態'] = df.apply(check_status, axis=1)

    # --- 5. 只保留指定欄位 ---
    keep_cols = [
        '即時狀態', '生產上線日', '完工日', '訂單編號[*]', '供應商名稱[*]', 
        '料件編號[*]', '物料名稱[*]', '規格[*]', '訂單量', '已交量', 
        '未交數量', '單身備註'
    ]
    
    return df[[c for c in keep_cols if c in df.columns]]

# --- 網頁介面 ---
st.title("📊 3003 生活館 - 跨檔案進料監控看板")

try:
    df = load_data()
    if not df.empty:
        # 儀表指標
        c1, c2, c3 = st.columns(3)
        c1.metric("❌ 影響生產", len(df[df['即時狀態'] == "❌ 警報：晚於生產日"]))
        c2.metric("🔴 逾期件數", len(df[df['即時狀態'] == "🔴 已逾期"]))
        c3.metric("📑 待交總項次", len(df[df['未交數量'] > 0]))

        q = st.text_input("🔍 快速搜尋 (供應商/料號/單號)", "")
        if q:
            mask = df.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)
            df = df[mask]
        
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.warning("請確保 GitHub 上同時存有『訂單資訊』與『排程資料庫』。")
except Exception as e:
    st.error(f"程式執行錯誤：{e}")
