import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

st.set_page_config(page_title="SRM進料x生產排程監控", layout="wide")

@st.cache_data
def load_data():
    # 1. 寬鬆搜尋檔案 (只要檔名包含 '訂單' 或 '排程' 且是 excel 就抓)
    all_excel = glob.glob("*.xls*")
    
    srm_file = None
    plan_file = None
    
    for f in all_excel:
        if "訂單" in f:
            srm_file = f
        if "排程" in f:
            plan_file = f

    # 偵錯：如果還是找不到，直接列出清單
    if not srm_file or not plan_file:
        st.error(f"❌ 讀取失敗！目前目錄下的 Excel 有：{all_excel}")
        st.info("請檢查 GitHub 上的檔案名稱是否包含 '訂單' 或 '排程' 關鍵字。")
        return pd.DataFrame()

    # 2. 讀取排程 (基準)
    df_plan_raw = pd.read_excel(plan_file, engine='openpyxl')
    
    # 根據您的排程表位置：K欄(10)料號, G欄(6)訂單量, M欄(12)上線日期
    df_main = df_plan_raw.iloc[:, [10, 6, 12]].copy()
    df_main.columns = ['料件編號[*]', '訂單量', '生產上線日']
    df_main['生產上線日'] = pd.to_datetime(df_main['生產上線日'], errors='coerce')
    df_main = df_main.groupby('料件編號[*]').agg({'訂單量': 'sum', '生產上線日': 'min'}).reset_index()

    # 3. 讀取進料 (ASN)
    df_srm = pd.read_excel(srm_file, engine='openpyxl')
    
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

    # 4. 合併與計算
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
    
    # 最終顯示欄位
    keep = ['即時狀態', '生產上線日', '完工日', '料件編號[*]', '物料名稱[*]', '規格[*]', '訂單量', '已交量', '未交數量']
    return df[keep]

# --- 介面 ---
st.title("📊 3003 生活館 - 進料 x 排程整合看板")

try:
    df = load_data()
    if not df.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("❌ 影響生產", len(df[df['即時狀態'] == "❌ 警報：晚於生產日"]))
        c2.metric("🔴 逾期未到", len(df[df['即時狀態'] == "🔴 已逾期"]))
        c3.metric("📑 追蹤項次", len(df))

        q = st.text_input("🔍 搜尋 (供應商/料號/單號)", "")
        if q:
            mask = df.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)
            df = df[mask]
        
        st.dataframe(df.sort_values("即時狀態", ascending=False), use_container_width=True, hide_index=True)
except Exception as e:
    st.error(f"系統運行錯誤：{e}")
