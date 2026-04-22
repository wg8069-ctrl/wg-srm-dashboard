import streamlit as st
import pandas as pd
import glob
import os
from datetime import datetime

st.set_page_config(page_title="SRM進料x生產排程監控", layout="wide")

@st.cache_data
def load_data():
    # 1. 搜尋檔案
    all_excel = glob.glob("*.xls*")
    srm_file = next((f for f in all_excel if "訂單" in f), None)
    plan_file = next((f for f in all_excel if "排程" in f), None)

    if not srm_file or not plan_file:
        st.error(f"❌ 讀取失敗！目錄下缺少必要檔案。")
        return pd.DataFrame()

    # 2. 讀取「排程資料庫」
    df_p = pd.read_excel(plan_file, engine='openpyxl')
    
    try:
        # 對應 CSV 欄位：產品品號、預計產量、上限日
        df_plan = pd.DataFrame({
            '料件編號[*]': df_p['產品品號'].astype(str).str.strip(),
            '訂單量': pd.to_numeric(df_p['預計產量'], errors='coerce').fillna(0),
            '生產上線日': pd.to_datetime(df_p['上限日'], errors='coerce')
        })
        # 彙總：同料號加總產量，取最早日期
        df_plan = df_plan.groupby('料件編號[*]').agg({'訂單量': 'sum', '生產上線日': 'min'}).reset_index()
    except Exception as e:
        st.error(f"❌ 排程表解析失敗，請確認欄位名稱：{e}")
        return pd.DataFrame()

    # 3. 讀取「訂單資訊」 (ASN 檔)
    df_s = pd.read_excel(srm_file, engine='openpyxl')
    df_s['料件編號[*]'] = df_s['料件編號[*]'].astype(str).str.strip()

    def calc_delivered(row):
        status = str(row.get('出貨狀態', '')).strip()
        send_val = row.get('發貨量[*]', 0)
        recv_val = row.get('收貨量[*]', 0)
        if status in ["已發貨", "全部發貨"]: return send_val
        elif status in ["部分收貨", "全部收貨"]: return recv_val
        return 0
    
    df_s['計算後已交'] = df_s.apply(calc_delivered, axis=1)
    
    # 彙總進料資訊
    df_s_agg = df_s.groupby('料件編號[*]').agg({
        '訂單編號[*]': 'first',
        '供應商名稱[*]': 'first',
        '物料名稱[*]': 'first',
        '規格[*]': 'first',
        '計算後已交': 'sum',
        '發貨日期[*]': 'max'
    }).reset_index()
    df_s_agg.rename(columns={'發貨日期[*]': '實際完工日'}, inplace=True)

    # 4. 合併資料
    df = pd.merge(df_plan, df_s_agg, on='料件編號[*]', how='left')
    df['已交量'] = df['計算後已交'].fillna(0)
    df['未交數量'] = df['訂單量'] - df['已交量']
    
    # 強制轉換日期格式，避免比較錯誤 (關鍵修正處)
    df['實際完工日'] = pd.to_datetime(df['實際完工日'], errors='coerce')
    df['生產上線日'] = pd.to_datetime(df['生產上線日'], errors='coerce')
    
    # 5. 即時狀態判斷
    today = pd.to_datetime(datetime.now().date())
    def check_status(row):
        if row['未交數量'] <= 0: return "✅ 已結案"
        
        # 警報：實際完工日 > 生產上線日 (上限日)
        if pd.notnull(row['實際完工日']) and pd.notnull(row['生產上線日']):
            if row['實際完工日'] > row['生產上線日']:
                return "❌ 警報：晚於生產日"
        
        # 逾期判斷
        comp_date = row['實際完工日'] if pd.notnull(row['實際完工日']) else row['生產上線日']
        if pd.notnull(comp_date) and comp_date < today:
            return "🔴 已逾期"
        return "🟢 正常待料"

    df['即時狀態'] = df.apply(check_status, axis=1)
    
    # 6. 保留課長要求的精確欄位
    keep_cols = [
        '訂單編號[*]', '供應商名稱[*]', '料件編號[*]', '物料名稱[*]', '規格[*]', 
        '訂單量', '已交量', '未交數量', '生產上線日', '實際完工日', '即時狀態'
    ]
    
    # 剔除不存在的欄位並過濾無效列
    df = df[df['訂單量'] > 0]
    existing = [c for c in keep_cols if c in df.columns]
    
    # 排序：警報與逾期排最前面
    return df[existing].sort_values("即時狀態", ascending=False)

# --- 網頁介面 ---
st.title("📊 3003 生活館 - 進料 x 排程整合看板")

try:
    data = load_data()
    if not data.empty:
        # 頂部戰情指標
        c1, c2, c3 = st.columns(3)
        c1.metric("❌ 影響生產", len(data[data['即時狀態'] == "❌ 警報：晚於生產日"]))
        c2.metric("🔴 逾期未到", len(data[data['即時狀態'] == "🔴 已逾期"]))
        c3.metric("📑 追蹤項次", len(data))

        # 搜尋功能
        q = st.text_input("🔍 關鍵字搜尋 (料號/供應商/單號)", "")
        if q:
            mask = data.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)
            data = data[mask]
        
        # 顯示表格
        st.dataframe(data, use_container_width=True, hide_index=True)
    else:
        st.info("📊 系統運作中，請確保 Excel 檔案內容正確...")
except Exception as e:
    st.error(f"系統運行錯誤：{e}")
