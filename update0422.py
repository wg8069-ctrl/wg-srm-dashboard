import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

# 網頁配置
st.set_page_config(page_title="SRM 進料戰情看板", layout="wide")

@st.cache_data
def load_data():
    files = glob.glob("訂單資訊*.xls*")
    if not files:
        return pd.DataFrame()
    
    latest_file = max(files, key=os.path.getmtime)
    # 讀取 Excel
    df = pd.read_excel(latest_file, engine='openpyxl')

    # 1. 處理已交量計算 (根據您的公式)
    # 欄位對應：發貨量[*] -> N欄 / 收貨量[*] -> O欄
    def calculate_delivered(row):
        status = str(row.get('出貨狀態', '')).strip()
        # 抓發貨量
        send_val = row.get('發貨量[*]', 0)
        # 抓收貨量
        receive_val = row.get('收貨量[*]', 0)
        
        if status in ["已發貨", "全部發貨"]:
            return send_val
        elif status in ["部分收貨", "全部收貨"]:
            return receive_val
        return 0

    df['已交量'] = df.apply(calculate_delivered, axis=1)

    # 2. 處理未交數量 (由於 ASN 檔通常不帶原始訂單總量，我們先以發貨量為基準，或請確認您的檔案是否有總量欄位)
    # 如果檔案中沒有「訂單總量」，這裡暫時設為 0 或邏輯處理
    df['未交數量'] = 0 # 根據 ASN 檔性質，此處需確認是否有原始採購總量欄位

    # 3. 處理日期與即時狀態
    # 欄位對應：預計交貨日期 -> 發貨日期[*] (或是您的檔案中的日期欄)
    date_col = '發貨日期[*]' 
    df['預計交貨日期'] = pd.to_datetime(df[date_col], errors='coerce')
    today = pd.to_datetime(datetime.now().date())
    df['剩餘天數'] = (df['預計交貨日期'] - today).dt.days

    def get_realtime_status(row):
        status = str(row.get('出貨狀態', '')).strip()
        if status in ["全部收貨"]: return "✅ 已結案"
        if row['剩餘天數'] < 0: return "🔴 已逾期"
        if row['剩餘天數'] <= 3: return "🟡 三日內到貨"
        return "🟢 正常待料"
    
    df['即時狀態'] = df.apply(get_realtime_status, axis=1)

    # 4. 保留您指定的 15 個欄位 (校對後名稱)
    keep_cols = [
        '發貨日期[*]', '訂單編號[*]', '供應商名稱[*]', 
        '料件編號[*]', '物料名稱[*]', '規格[*]', '出貨狀態', 
        '發貨量[*]', '收貨量[*]', '已交量', '未交數量', '即時狀態'
    ]
    
    # 過濾掉不存在的欄位並更名為您想要的標題
    existing_cols = [c for c in keep_cols if c in df.columns or c in ['已交量', '未交數量', '即時狀態']]
    df = df[existing_cols]
    
    # 排除已作廢
    if '出貨狀態' in df.columns:
        df = df[df['出貨狀態'] != '已作廢']

    return df

# --- 網頁介面 ---
st.title("📊 3003 生活館 - SRM 進料戰情看板")

try:
    df = load_data()
    if not df.empty:
        # 指標摘要
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 逾期件數", len(df[df['即時狀態'] == "🔴 已逾期"]))
        c2.metric("🟡 近期急件", len(df[df['即時狀態'] == "🟡 三日內到貨"]))
        c3.metric("📑 總處理項次", len(df))

        # 搜尋功能
        search_q = st.text_input("🔍 關鍵字搜尋 (供應商、料號、單號)", "")
        if search_q:
            mask = df.apply(lambda r: r.astype(str).str.contains(search_q, case=False).any(), axis=1)
            df = df[mask]
        
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.warning("請確保 GitHub 上有『訂單資訊』開頭的 Excel 檔案。")
except Exception as e:
    st.error(f"程式執行出錯，原因通常是 Excel 欄位名稱對不上：{e}")
