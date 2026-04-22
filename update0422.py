import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

st.set_page_config(page_title="SRM進料x生產排程監控", layout="wide")

@st.cache_data
def load_data():
    # --- 1. 讀取 SRM 進料檔 ---
    srm_files = glob.glob("訂單資訊*.xls*")
    df_srm = pd.DataFrame()
    if srm_files:
        latest_srm = max(srm_files, key=os.path.getmtime)
        df_srm = pd.read_excel(latest_srm, engine='openpyxl')
        
    # --- 2. 讀取 生產排程資料庫 (M 欄上線日期) ---
    plan_files = glob.glob("*排程資料庫*.xls*")
    df_plan = pd.DataFrame()
    if plan_files:
        latest_plan = max(plan_files, key=os.path.getmtime)
        df_plan_raw = pd.read_excel(latest_plan, engine='openpyxl')
        if len(df_plan_raw.columns) >= 13:
            df_plan = df_plan_raw.iloc[:, [10, 12]].copy()
            df_plan.columns = ['料件編號[*]', '生產上線日']
            df_plan['生產上線日'] = pd.to_datetime(df_plan['生產上線日'], errors='coerce')
            df_plan = df_plan.groupby('料件編號[*]')['生產上線日'].min().reset_index()

    if df_srm.empty:
        return pd.DataFrame()

    # --- 3. 處理「完工日」 (從原本的發貨日期往後移，改抓完工日欄位) ---
    # 這裡會優先找名稱包含 "完工" 的欄位，如果沒有則維持原日期
    finish_col = next((c for c in df_srm.columns if '完工' in c), '發貨日期[*]')
    df_srm['完工日'] = pd.to_datetime(df_srm[finish_col], errors='coerce')

    # --- 4. 合併排程日期 ---
    if not df_plan.empty:
        df_srm = pd.merge(df_srm, df_plan, on='料件編號[*]', how='left')

    # --- 5. 計算已交量與狀態 (改以完工日比對) ---
    def calc_delivered(row):
        status = str(row.get('出貨狀態', '')).strip()
        if status in ["已發貨", "全部發貨"]: return row.get('發貨量[*]', 0)
        elif status in ["部分收貨", "全部收貨"]: return row.get('收貨量[*]', 0)
        return 0

    df_srm['已交量'] = df_srm.apply(calc_delivered, axis=1)
    today = pd.to_datetime(datetime.now().date())
    
    def check_status(row):
        if str(row.get('出貨狀態')) == "全部收貨": return "✅ 已結案"
        
        # 核心比對：完工日 vs 生產上線日
        if pd.notnull(row.get('生產上線日')) and pd.notnull(row.get('完工日')):
            if row['完工日'] > row['生產上線日']:
                return "❌ 警報：晚於生產日"
        
        days_diff = (row['完工日'] - today).days if pd.notnull(row['完工日']) else 999
        if days_diff < 0: return "🔴 已逾期"
        if days_diff <= 3: return "🟡 三日內完工"
        return "🟢 正常待料"

    df_srm['即時狀態'] = df_srm.apply(check_status, axis=1)
    
    # 欄位排序：讓「完工日」與「生產上線日」並列在最前面
    display_cols = [
        '即時狀態', '完工日', '生產上線日', '料件編號[*]', 
        '物料名稱[*]', '規格[*]', '供應商名稱[*]', '出貨狀態', '發貨量[*]', '已交量'
    ]
    
    return df_srm[[c for c in display_cols if c in df_srm.columns]]

# --- 介面呈現 ---
st.title("📊 3003 生活館 - 完工 vs 生產同步對照表")

try:
    df = load_data()
    if not df.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("❌ 影響生產(趕不及)", len(df[df['即時狀態'] == "❌ 警報：晚於生產日"]))
        c2.metric("🔴 已逾期", len(df[df['即時狀態'] == "🔴 已逾期"]))
        c3.metric("📑 追蹤總項次", len(df))

        q = st.text_input("🔍 搜尋 (供應商/料號/單號)", "")
        if q:
            mask = df.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)
            df = df[mask]
        
        df = df.sort_values(by="即時狀態", ascending=False)
        st.dataframe(df, use_container_width=True, hide_index=True)
        
        st.info("💡 說明：看板已將比對基準改為『完工日』，若完工日晚於排程的上線日期，會自動觸發❌警報。")
    else:
        st.warning("請在 GitHub 上傳『訂單資訊』與『排程資料庫』。")
except Exception as e:
    st.error(f"程式執行出錯：{e}")
