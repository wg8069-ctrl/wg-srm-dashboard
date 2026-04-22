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
        st.error(f"❌ 讀取失敗！目錄下缺少檔案。")
        return pd.DataFrame()

    # 2. 讀取排程 (處理無標題且有空白列的情況)
    # 不指定 header，讀進來後再用 iloc 強制抓取位置
    df_p = pd.read_excel(plan_file, engine='openpyxl', header=None)
    
    try:
        # 強制抓取：K欄(10)=料號, G欄(6)=訂單量, M欄(12)=上線日
        # iloc[:, 10] 代表第 11 欄
        df_plan = pd.DataFrame({
            '料件編號[*]': df_p.iloc[:, 10],
            '訂單量': df_p.iloc[:, 6],
            '生產上線日': df_p.iloc[:, 12]
        })
        
        # 清理資料：轉為數字與日期，並移除無法辨識的列
        df_plan['訂單量'] = pd.to_numeric(df_plan['訂單量'], errors='coerce').fillna(0)
        df_plan['生產上線日'] = pd.to_datetime(df_plan['生產上線日'], errors='coerce')
        
        # 關鍵：料號必須是字串且不能為空，過濾掉 Excel 開頭的標題或空白行
        df_plan = df_plan[df_plan['料件編號[*]'].notnull()]
        df_plan['料件編號[*]'] = df_plan['料件編號[*]'].astype(str).str.strip()
        
        # 彙總
        df_plan = df_plan.groupby('料件編號[*]').agg({'訂單量': 'sum', '生產上線日': 'min'}).reset_index()
    except Exception as e:
        st.error(f"❌ 排程表解析失敗：{e}")
        return pd.DataFrame()

    # 3. 讀取進料 (ASN)
    df_s = pd.read_excel(srm_file, engine='openpyxl')
    
    # 確保料號格式一致
    if '料件編號[*]' in df_s.columns:
        df_s['料件編號[*]'] = df_s['料件編號[*]'].astype(str).str.strip()

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

    # 4. 合併資料 (以排程檔為主)
    df = pd.merge(df_plan, df_s_agg, on='料件編號[*]', how='left')
    
    # 5. 計算狀態
    df['已交量'] = df['計算後已交'].fillna(0)
    df['未交數量'] = df['訂單量'] - df['已交量']
    
    today = pd.to_datetime(datetime.now().date())
    def check_status(row):
        if row['未交數量'] <= 0: return "✅ 已結案"
        if pd.notnull(row.get('生產上線日')) and pd.notnull(row.get('完工日')):
            if row['完工日'] > row['生產上線日']: return "❌ 警報：晚於生產日"
        
        # 逾期判斷
        comp_date = row.get('完工日') if pd.notnull(row.get('完工日')) else row.get('生產上線日')
        if pd.notnull(comp_date) and comp_date < today: return "🔴 已逾期"
        return "🟢 正常待料"

    df['即時狀態'] = df.apply(check_status, axis=1)
    
    # 只顯示有意義的追蹤項次
    df = df[df['訂單量'] > 0]
    
    keep = ['即時狀態', '生產上線日', '完工日', '料件編號[*]', '物料名稱[*]', '規格[*]', '訂單量', '已交量', '未交數量', '供應商名稱[*]']
    return df[[c for c in keep if c in df.columns]]

# --- 介面 ---
st.title("📊 3003 生活館 - 進料 x 生產同步看板")

try:
    data = load_data()
    if not data.empty:
        # 排序：警報優先
        st.dataframe(data.sort_values("即時狀態", ascending=False), use_container_width=True, hide_index=True)
    else:
        st.info("📊 系統已就緒，等待正確的 Excel 資料讀入...")
except Exception as e:
    st.error(f"系統運行錯誤：{e}")
