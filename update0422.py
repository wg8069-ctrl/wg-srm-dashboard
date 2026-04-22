import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

st.set_page_config(page_title="SRM進料x生產排程監控", layout="wide")

@st.cache_data
def load_data():
    # --- 1. 讀取 生產排程資料庫 (作為訂單量與上線日期的基準) ---
    plan_files = glob.glob("*排程資料庫*.xls*")
    df_main = pd.DataFrame()
    if plan_files:
        latest_plan = max(plan_files, key=os.path.getmtime)
        df_plan_raw = pd.read_excel(latest_plan, engine='openpyxl')
        
        # 根據排程表：G欄(Index 6)訂單量, K欄(Index 10)料號, M欄(Index 12)上線日期
        if len(df_plan_raw.columns) >= 13:
            df_main = df_plan_raw.iloc[:, [10, 6, 12]].copy()
            df_main.columns = ['料件編號[*]', '訂單量', '生產上線日']
            df_main['生產上線日'] = pd.to_datetime(df_main['生產上線日'], errors='coerce')
            # 同料號加總訂單量，取最早生產日期
            df_main = df_main.groupby('料件編號[*]').agg({'訂單量': 'sum', '生產上線日': 'min'}).reset_index()

    # --- 2. 讀取 SRM 進料檔 (計算已交量) ---
    srm_files = glob.glob("訂單資訊*.xls*")
    df_srm_agg = pd.DataFrame()
    if srm_files:
        latest_srm = max(srm_files, key=os.path.getmtime)
        df_srm = pd.read_excel(latest_srm, engine='openpyxl')
        
        # 已交量邏輯：已發貨/全部發貨抓發貨量，部分收貨抓收貨量
        def calc_delivered(row):
            status = str(row.get('出貨狀態', '')).strip()
            send_val = row.get('發貨量[*]', 0)
            recv_val = row.get('收貨量[*]', 0)
            if status in ["已發貨", "全部發貨"]: return send_val
            elif status in ["部分收貨", "全部收貨"]: return recv_val
            return 0
        
        df_srm['計算後已交'] = df_srm.apply(calc_delivered, axis=1)
        
        # 以料號彙總進料資訊
        df_srm_agg = df_srm.groupby('料件編號[*]').agg({
            '訂單編號[*]': 'first',
            '供應商名稱[*]': 'first',
            '物料名稱[*]': 'first',
            '規格[*]': 'first',
            '計算後已交': 'sum',
            '發貨日期[*]': 'max'
        }).reset_index()
        df_srm_agg.rename(columns={'發貨日期[*]': '完工日'}, inplace=True)

    if df_main.empty:
        return pd.DataFrame()

    # --- 3. 合併資料 (以排程檔為核心) ---
    df = pd.merge(df_main, df_srm_agg, on='料件編號[*]', how='left')

    # --- 4. 計算未交數量與狀態 ---
    df['已交量'] = df['計算後已交'].fillna(0)
    df['未交數量'] = df['訂單量'] - df['已交量']
    
    today = pd.to_datetime(datetime.now().date())
    def check_status(row):
        if row['未交數量'] <= 0: return "✅ 已結案"
        
        # 比對完工日(SRM)與上線日(排程)
        if pd.notnull(row.get('生產上線日')) and pd.notnull(row.get('完工日')):
            if row['完工日'] > row['生產上線日']:
                return "❌ 警報：晚於生產日"
        
        days_diff = (pd.to_datetime(row.get('完工日')) - today).days if pd.notnull(row.get('完工日')) else 999
        if days_diff < 0: return "🔴 已逾期"
        return "🟢 正常待料"

    df['即時狀態'] = df.apply(check_status, axis=1)

    # --- 5. 最終保留欄位 (已移除單身備註) ---
    keep_cols = [
        '即時狀態', '生產上線日', '完工日', '訂單編號[*]', '供應商名稱[*]', 
        '料件編號[*]', '物料名稱[*]', '規格[*]', '訂單量', '已交量', '未交數量'
    ]
    
    return df[[c for c in keep_cols if c in df.columns]]

# --- 網頁介面 ---
st.title("📊 3003 生活館 - 跨檔案進料監控看板")

try:
    df = load_data()
    if not df.empty:
        # 指標顯示
        c1, c2, c3 = st.columns(3)
        c1.metric("❌ 影響生產", len(df[df['即時狀態'] == "❌ 警報：晚於生產日"]))
        c2.metric("🔴 逾期件數", len(df[df['即時狀態'] == "🔴 已逾期"]))
        c3.metric("📑 待交總項次", len(df[df['未交數量'] > 0]))

        # 搜尋功能
        q = st.text_input("🔍 快速搜尋 (供應商/料號/單號)", "")
        if q:
            mask = df.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)
            df = df[mask]
        
        # 依照狀態排序 (警報最優先)
        df = df.sort_values(by="即時狀態", ascending=False)
        
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("請確保 GitHub 上同時存有『訂單資訊』與『排程資料庫』Excel 檔案。")
except Exception as e:
    st.error(f"程式執行錯誤：{e}")
