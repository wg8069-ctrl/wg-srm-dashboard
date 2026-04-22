import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

# 網頁配置
st.set_page_config(page_title="SRM 進料監控看板", layout="wide")

@st.cache_data
def load_data():
    files = glob.glob("訂單資訊*.xls*")
    if not files:
        return pd.DataFrame()
    
    latest_file = max(files, key=os.path.getmtime)
    df = pd.read_excel(latest_file, engine='openpyxl')

    # 1. 處理「已交量」計算邏輯 (照您的公式)
    def calculate_delivered(row):
        status = str(row.get('狀態', '')).strip()
        if status == "已發貨" or status == "全部發貨":
            return row.get('發貨量', 0)
        elif status == "部分收貨":
            return row.get('收貨量', 0)
        elif status == "全部收貨": # 增加一個保險，全部收貨通常也算收貨量
            return row.get('收貨量', 0)
        return 0

    df['已交量'] = df.apply(calculate_delivered, axis=1)

    # 2. 處理「未交數量」計算邏輯 (訂單量 - 已交量)
    # 注意：請確認您的 Excel 欄位是叫 '採購數量' 還是 '訂單量'，這裡我先幫您做相容
    order_col = '採購數量' if '採購數量' in df.columns else '訂單量'
    df['未交數量'] = df[order_col] - df['已交量']

    # 3. 處理「即時狀態」與「日期」
    date_col = '預計交貨日期'
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    today = pd.to_datetime(datetime.now().date())
    df['剩餘天數'] = (df[date_col] - today).dt.days

    def get_realtime_status(row):
        if row['未交數量'] <= 0: return "✅ 已結案"
        if row['剩餘天數'] < 0: return "🔴 已逾期"
        if row['剩餘天數'] <= 3: return "🟡 三日內到貨"
        return "🟢 正常待料"
    
    df['即時狀態'] = df.apply(get_realtime_status, axis=1)

    # 4. 只保留您指定的欄位 (精確過濾)
    keep_cols = [
        '預計交貨日期', '訂單編號', '採購日期', '供應商編號', '供應商名稱', 
        '物料編碼', '物料名稱', '規格', '倉庫名稱', '狀態', 
        order_col, '已交量', '未交數量', '單身備註', '即時狀態'
    ]
    
    # 過濾掉不存在的欄位避免報錯
    existing_cols = [c for c in keep_cols if c in df.columns]
    df = df[existing_cols]
    
    # 排除已作廢
    if '狀態' in df.columns:
        df = df[df['狀態'] != '已作廢']

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
        c3.metric("📑 待交總項次", len(df[df['未交數量'] > 0]))

        # 搜尋功能
        search_q = st.text_input("🔍 關鍵字搜尋 (支援供應商、料號、單號)", "")
        if search_q:
            mask = df.apply(lambda r: r.astype(str).str.contains(search_q, case=False).any(), axis=1)
            df = df[mask]
        
        # 顯示表格 (使用高亮顯示)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.warning("請確保 GitHub 上有『訂單資訊』開頭的 Excel 檔案。")
except Exception as e:
    st.error(f"程式執行出錯：{e}")
