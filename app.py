import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# 1. 頁面基礎設定
st.set_page_config(page_title="儲互社分析系統 v5.0", layout="wide", page_icon="🏦")

# --- 自定義 CSS (保持旗艦質感) ---
st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    .stat-card {
        background: white; border-radius: 12px; border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); margin-bottom: 1.5rem;
        height: 240px; display: flex; flex-direction: column; overflow: hidden;
    }
    .card-header {
        padding: 10px; color: white; font-weight: 700; font-size: 1.05rem;
        text-align: center; display: flex; align-items: center; justify-content: center; gap: 10px;
    }
    .header-red { background: linear-gradient(135deg, #EF4444, #991B1B); }
    .header-orange { background: linear-gradient(135deg, #F59E0B, #92400E); }
    .header-blue { background: linear-gradient(135deg, #3B82F6, #1E40AF); }
    .header-green { background: linear-gradient(135deg, #10B981, #065F46); }
    .card-body { padding: 15px; overflow-y: auto; flex-grow: 1; background: #FFFFFF; }
    .name-tag {
        display: inline-block; background: #F1F5F9; color: #334155; padding: 4px 10px;
        border-radius: 6px; margin: 4px; font-size: 0.9rem; border: 1px solid #CBD5E1; font-weight: 500;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 民國年月轉換函式 ---
def convert_minguo_date(val):
    try:
        s = str(int(val)).strip()
        if len(s) == 5:
            year, month = int(s[:3]) + 1911, int(s[3:])
        elif len(s) == 4:
            year, month = int(s[:2]) + 1911, int(s[2:])
        else: return pd.NaT
        return pd.to_datetime(f"{year}-{month}-01")
    except: return pd.NaT

# --- 【核心優化】多格式檔案處理引擎 ---
@st.cache_data(show_spinner="正在解析檔案並計算數據...")
def process_any_files(uploaded_files):
    df_main_raw, df_loan_raw = None, None
    
    for file in uploaded_files:
        # 如果是 Excel，一次讀兩個 Sheet
        if file.name.endswith('.xlsx'):
            df_main_raw = pd.read_excel(file, sheet_name="社務及資金運用情形", dtype={'社號': str, '年月': str})
            df_loan_raw = pd.read_excel(file, sheet_name="放款及逾期放款", dtype={'社號': str, '年月': str})
            break # Excel 優先，讀完就跳出
        
        # 如果是 CSV，根據檔名關鍵字判斷
        elif file.name.endswith('.csv'):
            # 使用 utf-8-sig 以支援 Excel 產出的中文 CSV
            temp_df = pd.read_csv(file, dtype={'社號': str, '年月': str}, encoding='utf-8-sig')
            if "社務" in file.name:
                df_main_raw = temp_df
            elif "放款" in file.name:
                df_loan_raw = temp_df

    if df_main_raw is None or df_loan_raw is None:
        return None, None, None

    # 資料清洗邏輯 (與先前一致)
    df_main_raw['年月'] = df_main_raw['年月'].apply(convert_minguo_date)
    df_loan_raw['年月'] = df_loan_raw['年月'].apply(convert_minguo_date)
    
    for col in ['社員數', '股金', '貸放比']: 
        if col in df_main_raw.columns:
            df_main_raw[col] = pd.to_numeric(df_main_raw[col], errors='coerce').fillna(0)
    
    df_main_raw['儲蓄率'] = pd.to_numeric(df_main_raw['儲蓄率'], errors='coerce').fillna(0) / 100
    df_loan_raw['逾放比'] = pd.to_numeric(df_loan_raw['逾放比'], errors='coerce').fillna(0)
    df_loan_raw['提撥率'] = pd.to_numeric(df_loan_raw['提撥率'], errors='coerce').fillna(0) / 100
    df_loan_raw['收支比'] = pd.to_numeric(df_loan_raw['收支比'], errors='coerce').fillna(0) / 100

    df_m = df_main_raw.dropna(subset=['年月']).sort_values(by=['社號', '年月'])
    df_l = df_loan_raw.dropna(subset=['年月']).sort_values(by=['社號', '年月'])

    # 12個月成長彙整邏輯
    max_date = df_m['年月'].max()
    date_12m_ago = max_date - pd.DateOffset(months=12)
    societies = df_m['社號'].unique()
    rows = []
    
    for s_no in societies:
        m_sub = df_m[df_m['社號'] == s_no]
        l_sub = df_l[df_l['社號'] == s_no]
        s_name = m_sub['社名'].iloc[0]
        
        def get_v(g, c, lat=True):
            if g.empty: return 0
            if lat:
                v = g[g['年月'] == max_date][c].values
                return v[0] if len(v)>0 else g.iloc[-1][c]
            else:
                v = g[g['年月'] <= date_12m_ago].tail(1)[c].values
                return v[0] if len(v)>0 else g.iloc[0][c]

        eM, sM = get_v(m_sub, '社員數', True), get_v(m_sub, '社員數', False)
        eS, sS = get_v(m_sub, '股金', True), get_v(m_sub, '股金', False)
        
        rows.append({
            '社號': s_no, '社名': s_name, '現有社員': eM, '社員成長率(12M)': (eM-sM)/sM if sM!=0 else 0,
            '現有股金': eS, '股金成長率(12M)': (eS-sS)/sS if sS!=0 else 0,
            '貸放比': get_v(m_sub, '貸放比', True), '儲蓄率': get_v(m_sub, '儲蓄率', True),
            '逾放比(初)': l_sub.iloc[0]['逾放比'] if not l_sub.empty else 0,
            '逾放比(末)': get_v(l_sub, '逾放比', True),
            '提撥率': get_v(l_sub, '提撥率', True), '收支比': get_v(l_sub, '收支比', True),
            'sM_total': sM, 'sS_total': sS
        })
    return pd.DataFrame(rows), df_m, df_l

# --- 主程式 ---
st.title("🏦 儲互社分析儀表板 v5.0 (雙格式相容)")
st.sidebar.header("📁 數據來源")
# 關鍵：設定 accept_multiple_files=True
uploaded_files = st.sidebar.file_uploader("上傳 Excel 或 CSV (可多選)", type=["xlsx", "csv"], accept_multiple_files=True)

if uploaded_files:
    data, df_m, df_l = process_any_files(uploaded_files)
    
    if data is None:
        st.error("⚠️ 檔案讀取失敗！請確保上傳了包含「社務」與「放款」關鍵字的檔案。")
    else:
        tab1, tab2, tab3 = st.tabs(["📊 經營總覽", "📋 詳細數據", "📈 歷史趨勢"])

        with tab1:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("全體社員數", f"{int(data['現有社員'].sum()):,}", f"{(data['現有社員'].sum()-data['sM_total'].sum())/data['sM_total'].sum():.2%}")
            m2.metric("全體股金水位", f"${data['現有股金'].sum()/1e8:.2f} 億", f"{(data['現有股金'].sum()-data['sS_total'].sum())/data['sS_total'].sum():.2%}")
            m3.metric("平均收支比", f"{data['收支比'].mean():.2%}")
            m4.metric("平均逾放比", f"{data['逾放比(末)'].mean():.2%}")

            st.markdown("<br>", unsafe_allow_html=True)
            hr = data[(data['逾放比(末)'] > data['逾放比(初)']) & (data['逾放比(末)'] > 0.1)]['社名'].tolist()
            liq = data[(data['貸放比'] > 0.9) & (data['股金成長率(12M)'] < 0)]['社名'].tolist()
            idl = data[(data['貸放比'] < 0.3) & (data['逾放比(末)'] < 0.02)]['社名'].tolist()
            std = data[(data['社員成長率(12M)'] > 0) & (data['股金成長率(12M)'] > 0) & (data['貸放比'] > 0.4) & (data['貸放比'] < 0.8) & (data['逾放比(末)'] < 0.02)]['社名'].tolist()

            c1, c2, c3, c4 = st.columns(4)
            def dc(t, i, ns, cl):
                tags = "".join([f'<span class="name-tag">{n}</span>' for n in ns]) if ns else '<div style="color:#94A3B8; text-align:center; margin-top:20px;">無標的</div>'
                st.markdown(f'<div class="stat-card"><div class="card-header {cl}">{i} {t}</div><div class="card-body">{tags}</div></div>', unsafe_allow_html=True)
            with c1: dc("高風險列管", "🚨", hr, "header-red")
            with c2: dc("流動性緊繃", "⚠️", liq, "header-orange")
            with c3: dc("資金閒置", "💤", idl, "header-blue")
            with c4: dc("穩健模範", "✅", std, "header-green")

        with tab2:
            st.dataframe(data.drop(columns=['sM_total', 'sS_total']).style.format({
                '社員成長率(12M)': '{:.2%}', '股金成長率(12M)': '{:.2%}', '貸放比': '{:.1%}', '儲蓄率': '{:.2%}', 
                '逾放比(初)': '{:.2%}', '逾放比(末)': '{:.2%}', '提撥率': '{:.2%}', '收支比': '{:.2%}', '現有社員': '{:,}', '現有股金': '${:,.0f}'
            }), use_container_width=True, height=500)

        with tab3:
            show_avg = st.checkbox("顯示全體平均線", value=True)
            df_all = pd.merge(df_m, df_l[['年月', '社號', '逾放比', '收支比']], on=['年月', '社號'], how='left')
            selected = st.multiselect("選擇比較的社：", options=data['社名'].unique(), default=data['社名'].iloc[0])
            if selected:
                avg_df = df_all.groupby('年月').mean(numeric_only=True).reset_index()
                avg_df['社名'] = '—— 全體平均 ——'
                plot_data = pd.concat([df_all[df_all['社名'].isin(selected)], avg_df]) if show_avg else df_all[df_all['社名'].isin(selected)]
                
                def dr(col, tit, pct=True):
                    fig = px.line(plot_data, x='年月', y=col, color='社名', title=tit, markers=True, color_discrete_map={'—— 全體平均 ——': '#333333'})
                    fig.for_each_trace(lambda t: t.update(line=dict(dash='dash', width=3)) if t.name == '—— 全體平均 ——' else ())
                    if pct: fig.update_layout(yaxis_tickformat='.1%')
                    st.plotly_chart(fig, use_container_width=True)

                c1, c2 = st.columns(2)
                with c1: dr('現有社員', "社員走勢", False) if '現有社員' in plot_data else dr('社員數', "社員走勢", False)
                with c2: dr('貸放比', "貸放走勢")
                c3, c4 = st.columns(2)
                with c3: dr('儲蓄率', "儲蓄走勢")
                with c4: dr('逾放比', "逾放走勢")
                dr('收支比', "收支走勢")
else:
    st.info("👋 請上傳包含「社務」與「放款」字樣的 Excel 或 CSV 檔案開始分析。")