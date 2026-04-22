import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

st.set_page_config(page_title="SRM進料x生產排程整合", layout="wide")

@st.cache_data
def load_data():
    # 1. 搜尋檔案
    all_excel = glob.glob("*.xls*")
    srm_file = next((f for f in all_excel if "訂單" in f), None)
    plan_file = next((f for f in all_excel if "排程" in f), None)

    if not srm_file or not plan_file:
        st.error(f"❌ 讀取失敗！目錄下的檔案：{all_excel}")
        return pd.DataFrame()

    # 2. 讀取排程 (改用關鍵字找欄位，避免 index out-of-bounds)
    df_p = pd.read_excel(plan_file, engine='openpyxl')
    
    # 自動尋找關鍵欄位名稱
    col_map_p = {
        '料號': next((c for c in df_p.columns if '料號' in str(c) or '料件編號' in str(c)), None),
        '訂單量': next((c for c in df_p.columns if '訂單量' in str(c) or '採購數量' in str(c) or '數量' in str(c)), None),
        '上線日': next((c for c in df_p.columns if '日期' in str(c) or '上線' in str(c) or '開工' in str(c)), None)
    }

    # 如果關鍵字找不到，改用預設位置 (G=6, K=10, M=12)
    col_no = col_map_p['料號'] if col_map_p['料號'] else df_p.columns[10]
    col_qty = col_map_p['訂單量'] if col_map_p['訂單量'] else df_p.columns[6]
    col_date = col_map_p['上線日'] if col_map_p['上線日'] else df_p.columns[12]

    df_main = df_p[[col_no, col_qty, col_date]].copy()
    df_main.columns = ['料件編號[*]', '訂單量', '生產上線日']
    df_main['生產上線日'] = pd.to_datetime(df_main['生產上線日'], errors='coerce')
    df_main = df_main.groupby('料件編號[*]').agg({'訂單量': 'sum', '生產上線日': 'min'}).reset_index()

    # 3. 讀取進料 (ASN)
    df_s = pd.read_excel(srm_file, engine='openpyxl')
    
    def calc_delivered(row):
        status = str(row.get('出貨狀態', '')).strip()
        send_val = row.get('發貨量[*]', 0)
        recv_val = row.get('收貨量[*]', 0)
        if status in ["已發貨", "全部發貨"]: return send_val
        elif status in ["部分收貨", "全部收貨"]: return recv_val
        return 0
    
    df_s['計算後已交'] = df_s.apply(calc_delivered, axis=1)
    df_s_agg = df_s.groupby('料件編號[*]').agg({
        '訂單編號[*]': 'first',
        '供應商名稱[*]': 'first',
        '物料名稱[*]': 'first',
        '規格[*]': 'first',
        '計算後已交': 'sum',
        '發貨日期[*]': 'max'
    }).reset_index()
    df_s_agg.rename(columns={'發貨日期[*]': '完工日'}, inplace=True)

    # 4. 合併與計算
    df = pd.merge(df_main, df_s_agg, on='料件編號[*]', how='left')
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
st.title("📊 3003 生活館 - 跨檔案整合看板")

try:
    data = load_data()
    if not data.empty:
        st.dataframe(data.sort_values("即時狀態", ascending=False), use_container_width=True, hide_index=True)
except Exception as e:
    st.error(f"系統運行錯誤：{e}")
