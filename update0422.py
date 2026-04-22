import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

st.set_page_config(page_title="SRM進料監控看板", layout="wide")

@st.cache_data
def load_data():
    # 1. 搜尋檔案
    all_excel = glob.glob("*.xls*")
    srm_file = next((f for f in all_excel if "訂單" in f), None)
    plan_file = next((f for f in all_excel if "排程" in f), None)

    if not srm_file or not plan_file:
        st.error("❌ GitHub 目錄下缺少『訂單資訊』或『排程資料庫』檔案")
        return pd.DataFrame()

    # 2. 讀取「排程資料庫」
    df_p = pd.read_excel(plan_file, engine='openpyxl')
    
    # 自動定位標題列 (針對截圖修正)
    if '產品品號' not in df_p.columns:
        for i in range(min(len(df_p), 10)):
            if '產品品號' in df_p.iloc[i].values:
                df_p.columns = df_p.iloc[i]
                df_p = df_p.iloc[i+1:].reset_index(drop=True)
                break

    # 準備排程資料
    try:
        df_plan = pd.DataFrame({
            '料號Key': df_p['產品品號'].astype(str).str.strip().str.upper(),
            '訂單量': pd.to_numeric(df_p['預計產量'], errors='coerce').fillna(0),
            '生產上線日': pd.to_datetime(df_p['上限日'], errors='coerce')
        })
        # 排除掉標題列本身 (如果它被誤抓進來)
        df_plan = df_plan[df_plan['料號Key'] != '產品品號']
        df_plan = df_plan[df_plan['料號Key'] != 'NAN']
        
        # 彙總
        df_plan = df_plan.groupby('料號Key').agg({'訂單量': 'sum', '生產上線日': 'min'}).reset_index()
    except Exception as e:
        st.error(f"❌ 解析排程表失敗：{e}")
        return pd.DataFrame()

    # 3. 讀取「訂單資訊」 (ASN 檔)
    df_s = pd.read_excel(srm_file, engine='openpyxl')
    
    # 統一料號格式
    df_s['料號Key'] = df_s['料件編號[*]'].astype(str).str.strip().str.upper()

    def calc_delivered(row):
        status = str(row.get('出貨狀態', '')).strip()
        send_val = row.get('發貨量[*]', 0)
        recv_val = row.get('收貨量[*]', 0)
        if status in ["已發貨", "全部發貨"]: return send_val
        elif status in ["部分收貨", "全部收貨"]: return recv_val
        return 0
    
    df_s['已交'] = df_s.apply(calc_delivered, axis=1)
    
    # 彙總進料資訊
    df_s_agg = df_s.groupby('料號Key').agg({
        '訂單編號[*]': 'first',
        '供應商名稱[*]': 'first',
        '物料名稱[*]': 'first',
        '規格[*]': 'first',
        '已交': 'sum',
        '發貨日期[*]': 'max'
    }).reset_index()
    df_s_agg.rename(columns={'發貨日期[*]': '實際完工日'}, inplace=True)

    # 4. 合併資料 (以排程為基準)
    df = pd.merge(df_plan, df_s_agg, on='料號Key', how='left')
    df['已交量'] = df['已交'].fillna(0)
    df['未交數量'] = df['訂單量'] - df['已交量']
    
    # 5. 狀態判斷
    today = pd.to_datetime(datetime.now().date())
    def check_status(row):
        if row['未交數量'] <= 0: return "✅ 已結案"
        
        actual_date = pd.to_datetime(row.get('實際完工日'), errors='coerce')
        plan_date = pd.to_datetime(row.get('生產上線日'), errors='coerce')
        
        if pd.notnull(actual_date) and pd.notnull(plan_date):
            if actual_date > plan_date: return "❌ 警報：晚於生產日"
        
        target = actual_date if pd.notnull(actual_date) else plan_date
        if pd.notnull(target) and target < today: return "🔴 已逾期"
        return "🟢 正常待料"

    df['即時狀態'] = df.apply(check_status, axis=1)
    
    # 欄位整理
    df.rename(columns={'料號Key': '料件編號[*]'}, inplace=True)
    keep = ['即時狀態', '生產上線日', '實際完工日', '料件編號[*]', '物料名稱[*]', '規格[*]', '訂單量', '已交量', '未交數量', '供應商名稱[*]']
    
    # 過濾：只留有訂單量的
    df = df[df['訂單量'] > 0]
    
    return df[[c for c in keep if c in df.columns]].sort_values("即時狀態", ascending=False)

# --- 介面 ---
st.title("📊 3003 生活館 - 進料 x 生產同步看板")

try:
    data = load_data()
    if not data.empty:
        st.dataframe(data, use_container_width=True, hide_index=True)
        st.success(f"✅ 讀取成功！共追蹤 {len(data)} 項次。")
    else:
        st.warning("⚠️ 檔案讀取成功，但找不到有效的『產品品號』數據。")
except Exception as e:
    st.error(f"系統運行錯誤：{e}")
