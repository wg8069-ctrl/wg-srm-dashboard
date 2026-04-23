import streamlit as st
import pandas as pd
import glob
import os
import pymssql 
from datetime import datetime

# --- 黃課長請填入您的 ERP 連線資訊 ---
DB_CONFIG = {
    'server': '192.168.1.50:1433',
    'user': 'sa',
    'password': 'dsc@80690891',
    'database': 'WG'
}

st.set_page_config(page_title="ERP 製令進度核對看板", layout="wide")

@st.cache_data(ttl=300)
def fetch_erp_actual_data():
    try:
        conn = pymssql.connect(**DB_CONFIG)
        # TA001-TA002=製令單號, TA006=品號, TA015=預計產量, TA009=上限日, TA011=狀態
        query = """
        SELECT 
            TA001 + '-' + TA002 AS [製令單號],
            TA006 AS [料件編號], 
            TA015 AS [預計產量], 
            TA009 AS [生產上限日],
            TA011 AS [狀態代碼],
            TA034 AS [品名],
            TA035 AS [規格]
        FROM MOCTA 
        WHERE TA011 IN ('1', '2', '3')
        """
        df_erp = pd.read_sql(query, conn)
        conn.close()
        
        status_map = {'1': '1-未領料', '2': '2-已領料', '3': '3-生產中'}
        df_erp['製令狀態'] = df_erp['狀態代碼'].map(status_map)
        df_erp['料件編號'] = df_erp['料件編號'].astype(str).str.strip().str.upper()
        return df_erp
    except Exception as e:
        st.error(f"❌ ERP 連線失敗，請檢查網路或帳密：{e}")
        return pd.DataFrame()

@st.cache_data
def load_srm_excel():
    files = glob.glob("*訂單資訊*.xls*")
    if not files:
        st.error("❌ 資料夾內找不到『訂單資訊』Excel 檔案！")
        return pd.DataFrame()
    
    latest = max(files, key=os.path.getmtime)
    df_s = pd.read_excel(latest, engine='openpyxl')

    # --- 關鍵修正：自動尋找料號欄位 ---
    # 不管它叫 '料件編號[*]' 還是 '產品品號' 或 '料號'，只要有這些關鍵字就抓
    col_map = {
        '料號': next((c for c in df_s.columns if '料' in str(c) and '編' in str(c) or '品號' in str(c)), None),
        '發貨量': next((c for c in df_s.columns if '發貨量' in str(c)), None),
        '收貨量': next((c for c in df_s.columns if '收貨量' in str(c)), None),
        '狀態': next((c for c in df_s.columns if '狀態' in str(c)), None)
    }

    if not col_map['料號']:
        st.error(f"❌ 在 Excel 中找不到『料號』相關欄位。目前的欄位有：{list(df_s.columns)}")
        return pd.DataFrame()

    # 統一欄位名稱
    df_s['料件編號'] = df_s[col_map['料號']].astype(str).str.strip().str.upper()

    # 已交量計算邏輯 (使用模糊搜尋到的欄位)
    def calc_delivered(row):
        status = str(row.get(col_map['狀態'], '')).strip()
        send_val = row.get(col_map['發貨量'], 0)
        recv_val = row.get(col_map['收貨量'], 0)
        if "已發貨" in status or "全部發貨" in status: return send_val
        elif "部分收貨" in status or "全部收貨" in status: return recv_val
        return 0
    
    df_s['已交量'] = df_s.apply(calc_delivered, axis=1)
    
    # 彙總 SRM 資訊
    return df_s.groupby('料件編號').agg({
        '供應商名稱[*]': 'first' if '供應商名稱[*]' in df_s.columns else 'first', # 此處可視情況微調
        '已交量': 'sum'
    }).reset_index()

# --- 介面呈現 ---
st.title("進料核對系統")

df_erp = fetch_erp_actual_data()
df_srm = load_srm_excel()

if not df_erp.empty:
    # 兩邊都用「料件編號」合併
    df = pd.merge(df_erp, df_srm, on='料件編號', how='left')
    
    df['已交量'] = df['已交量'].fillna(0)
    df['未交缺口'] = df['預計產量'] - df['已交量']
    
    # 判定警報
    today = datetime.now()
    def check_alarm(row):
        if row['未交缺口'] <= 0: return "✅ 已到齊"
        limit_date = pd.to_datetime(row['生產上限日'], errors='coerce')
        if pd.notnull(limit_date) and limit_date < today: return "🔴 逾期生產"
        return f"🟢 {row['製令狀態']}"

    df['監控狀態'] = df.apply(check_alarm, axis=1)

    # 顯示指標
    c1, c2, c3 = st.columns(3)
    c1.metric("🔴 逾期生產", len(df[df['監控狀態'] == "🔴 逾期生產"]))
    c2.metric("⚙️ 產線中(TA011=3)", len(df[df['狀態代碼'] == '3']))
    c3.metric("📦 總缺料件數", len(df[df['未交缺口'] > 0]))

    # 排序：讓最急的(逾期)排在最前面
    df = df.sort_values(by=["監控狀態", "生產上限日"], ascending=[True, True])
    
    st.dataframe(df[['監控狀態', '生產上限日', '料件編號', '品名', '預計產量', '已交量', '未交缺口', '製令單號']], 
                 use_container_width=True, hide_index=True)
else:
    st.info("💡 請確認 ERP 連線資訊，或檢查本地 Excel 是否正確。")
