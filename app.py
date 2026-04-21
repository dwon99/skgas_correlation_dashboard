import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from itertools import combinations
from io import BytesIO
from datetime import datetime, date
import re
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Fundamental Driver Correlation Dashboard", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    div[data-testid="metric-container"] { background: linear-gradient(135deg, #f0f9f0 0%, #e8f5e9 100%); border: 1px solid #c8e6c9; border-radius: 10px; padding: 12px 16px; }
    .stDownloadButton > button { background-color: #006D36; color: white; border: none; border-radius: 8px; }
    .stDownloadButton > button:hover { background-color: #2D8C3C; color: white; }
</style>
""", unsafe_allow_html=True)

# ─── Helpers ───
def parse_flexible_date(date_val):
    if pd.isna(date_val): return pd.NaT
    if isinstance(date_val, (datetime, pd.Timestamp)): return pd.Timestamp(date_val)
    s = str(date_val).strip()
    for fmt in ['%Y-%m-%d','%Y/%m/%d','%m/%d/%Y','%d/%m/%Y','%Y-%m-%d %H:%M:%S','%m/%d/%Y %H:%M:%S','%d-%b-%y','%d-%b-%Y','%b-%d-%Y','%b %d, %Y','%Y%m%d','%d.%m.%Y','%Y.%m.%d']:
        try: return pd.Timestamp(datetime.strptime(s, fmt))
        except: continue
    try: return pd.Timestamp(pd.to_datetime(s, dayfirst=False))
    except: pass
    try: return pd.Timestamp(pd.to_datetime(s, dayfirst=True))
    except: return pd.NaT

def parse_month_str(m):
    try: return pd.Timestamp(str(m).strip()[:7] + '-01')
    except: return pd.NaT

def resolve_contract_month(base_year, y_offset, month):
    return f"{base_year + y_offset}-{month:02d}"

def safe_sheet_name(name):
    return re.sub(r'[\[\]:*?/\\]', '', str(name))[:31]

def expand_to_full_months(start, end):
    """윈도우 시작/종료를 해당 월의 1일/말일로 확장.
    월별 데이터(매월 1일 기록)가 윈도우에 겹치면 포함되도록."""
    start_expanded = pd.Timestamp(start.year, start.month, 1)
    end_expanded = pd.Timestamp(end.year, end.month, 1) + pd.offsets.MonthEnd(0)
    return start_expanded, end_expanded

def pearson_corr(x, y):
    mask = ~(np.isnan(x) | np.isnan(y))
    x, y = x[mask], y[mask]
    if len(x) < 3: return np.nan
    if np.std(x) == 0 or np.std(y) == 0: return np.nan
    return np.corrcoef(x, y)[0, 1]

def get_spread_curve(idx_df, idx1_name, month1, idx2_name, month2, start_date, end_date):
    df1 = idx_df[(idx_df['Index명']==idx1_name)&(idx_df['월물']==month1)]
    df2 = idx_df[(idx_df['Index명']==idx2_name)&(idx_df['월물']==month2)]
    df1 = df1[(df1['기준일자']>=start_date)&(df1['기준일자']<=end_date)]
    df2 = df2[(df2['기준일자']>=start_date)&(df2['기준일자']<=end_date)]
    merged = pd.merge(df1[['기준일자','Value']].rename(columns={'Value':'V1'}),
                      df2[['기준일자','Value']].rename(columns={'Value':'V2'}), on='기준일자', how='inner')
    merged['Spread'] = merged['V1'] - merged['V2']
    return merged[['기준일자','Spread']].sort_values('기준일자')

def get_fund_monthly(fund_df, index_name, driver_name, start_ym, end_ym):
    sub = fund_df[(fund_df['Index']==index_name)&(fund_df['Fundamental Driver']==driver_name)]
    sub = sub[(sub['YearMonth']>=start_ym)&(sub['YearMonth']<=end_ym)]
    monthly = sub.groupby('YearMonth')['Value'].mean().reset_index()
    monthly['Date'] = monthly['YearMonth'].dt.to_timestamp()
    return monthly.sort_values('Date')

def make_excel(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in sheets_dict.items():
            sn = safe_sheet_name(name)
            df.to_excel(writer, sheet_name=sn, index=False)
            ws = writer.sheets[sn]
            for i, col in enumerate(df.columns, 1):
                mx = max(len(str(col)), df[col].astype(str).str.len().max() if len(df)>0 else 0)
                ws.column_dimensions[chr(64+min(i,26))].width = min(mx+4, 40)
    output.seek(0)
    return output

# ─── Data Loading ───
@st.cache_data
def load_index_data(files):
    dfs = []
    for f in files:
        try: df = pd.read_csv(f, encoding='utf-8-sig')
        except:
            f.seek(0)
            try: df = pd.read_excel(f)
            except: f.seek(0); df = pd.read_csv(f, encoding='cp949')
        dfs.append(df)
    combined = pd.concat(dfs, ignore_index=True)
    col_map = {}
    for c in combined.columns:
        cl = c.strip().lower().replace(' ','_')
        if 'index_id' in cl: col_map[c]='INDEX_ID'
        elif 'index' in cl and ('명' in c or 'name' in cl): col_map[c]='Index명'
        elif '기준일자' in c or 'date' in cl or '일자' in c: col_map[c]='기준일자'
        elif '월물' in c or 'contract' in cl or 'month' in cl: col_map[c]='월물'
        elif '휴일' in c or 'holiday' in cl: col_map[c]='휴일여부'
        elif 'value' in cl or '값' in c: col_map[c]='Value'
    combined = combined.rename(columns=col_map)
    combined['기준일자'] = combined['기준일자'].apply(parse_flexible_date)
    combined['월물_dt'] = combined['월물'].apply(parse_month_str)
    combined['Value'] = pd.to_numeric(combined['Value'], errors='coerce')
    return combined.dropna(subset=['기준일자','Value'])

@st.cache_data(show_spinner="Fundamental 데이터 로딩 중...")
def load_fund_data(files):
    dfs = []
    for f in files:
        try: df = pd.read_csv(f, encoding='utf-8-sig')
        except:
            f.seek(0)
            try: df = pd.read_excel(f)
            except: f.seek(0); df = pd.read_csv(f, encoding='cp949')
        dfs.append(df)
    combined = pd.concat(dfs, ignore_index=True)
    # Strip whitespace from column names
    combined.columns = [c.strip() for c in combined.columns]
    col_map = {}
    for c in combined.columns:
        cl = c.strip().lower().replace(' ','_')
        if 'index' in cl and 'id' not in cl: col_map[c]='Index'
        elif 'driver' in cl or 'fundamental' in cl: col_map[c]='Fundamental Driver'
        elif 'date' in cl or '일자' in cl or '날짜' in cl: col_map[c]='Date'
        elif 'value' in cl or '값' in cl: col_map[c]='Value'
    combined = combined.rename(columns=col_map)
    # Fallback: if required columns missing, try exact match
    required = ['Index', 'Fundamental Driver', 'Date', 'Value']
    for req in required:
        if req not in combined.columns:
            for c in combined.columns:
                if c.strip().lower().replace(' ','_') == req.lower().replace(' ','_'):
                    combined = combined.rename(columns={c: req})
                    break
    missing = [r for r in required if r not in combined.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing}. 현재 컬럼: {list(combined.columns)}")
    combined['Date'] = combined['Date'].apply(parse_flexible_date)
    combined['Value'] = combined['Value'].apply(lambda x: pd.to_numeric(str(x).replace(',','').strip(), errors='coerce') if pd.notna(x) else np.nan)
    combined = combined.dropna(subset=['Date','Value'])
    combined['YearMonth'] = combined['Date'].dt.to_period('M')
    return combined

# ─── UI Components ───
def driver_multiselect(fund_df, key_prefix, label="Fundamental Driver 선택"):
    all_indices = sorted(fund_df['Index'].unique())
    sel_indices = st.multiselect("Index 선택 (Fundamental)", all_indices, default=all_indices[:1] if all_indices else [], key=f'{key_prefix}_fidx')
    if not sel_indices: return []
    opts = []
    for idx in sel_indices:
        for d in sorted(fund_df[fund_df['Index']==idx]['Fundamental Driver'].unique()):
            opts.append((f"[{idx}] {d}", idx, d))
    disp = [o[0] for o in opts]
    sel = st.multiselect(label, disp, default=disp[:1] if disp else [], key=f'{key_prefix}_fdrv')
    return [(idx,drv) for ds,idx,drv in opts if ds in sel]

def contract_month_selector(key_prefix, label="월물"):
    c1, c2 = st.columns(2)
    with c1: y_off = st.selectbox(f"{label} Y+?년", [0,1,2,3], format_func=lambda x:f"Y+{x}", key=f'{key_prefix}_yoff')
    with c2: mon = st.selectbox(f"{label} 월", list(range(1,13)), format_func=lambda x:f"{x}월", key=f'{key_prefix}_mon')
    return y_off, mon

def year_selector(key_prefix, fund_df=None, idx_df=None):
    ay = set()
    if fund_df is not None: ay |= set(fund_df['Date'].dt.year.unique())
    if idx_df is not None: ay |= set(idx_df['기준일자'].dt.year.unique())
    ay = sorted(ay)
    return st.multiselect("분석 연도", ay, default=ay[-3:] if len(ay)>=3 else ay, key=f'{key_prefix}_yrs')

def window_selector_simple(key_prefix, fund_df=None, idx_df=None):
    mode = st.radio("윈도우 유형", ["연도별 (커스텀 월-일)","자유 윈도우 (연-월-일)"], horizontal=True, key=f'{key_prefix}_mode')
    windows = []
    if mode == "연도별 (커스텀 월-일)":
        ay = set()
        if fund_df is not None: ay |= set(fund_df['Date'].dt.year.unique())
        if idx_df is not None: ay |= set(idx_df['기준일자'].dt.year.unique())
        ay = sorted(ay)
        sy = st.multiselect("연도 선택", ay, default=ay[-3:] if len(ay)>=3 else ay, key=f'{key_prefix}_yrs')
        c1,c2 = st.columns(2)
        with c1: smd = st.date_input("시작 월-일", value=date(2024,1,1), key=f'{key_prefix}_smd')
        with c2: emd = st.date_input("종료 월-일", value=date(2024,12,31), key=f'{key_prefix}_emd')
        for yr in sy:
            ws = pd.Timestamp(f'{yr}-{smd.month:02d}-{smd.day:02d}')
            we = pd.Timestamp(f'{yr}-{emd.month:02d}-{emd.day:02d}')
            windows.append((f"{yr} ({smd.month}/{smd.day}~{emd.month}/{emd.day})", ws, we, yr))
    else:
        nw = st.number_input("윈도우 개수", 1, 10, 2, key=f'{key_prefix}_nw')
        for i in range(int(nw)):
            c1,c2 = st.columns(2)
            with c1: ws = st.date_input(f"윈도우 {i+1} 시작", key=f'{key_prefix}_ws_{i}')
            with c2: we = st.date_input(f"윈도우 {i+1} 종료", key=f'{key_prefix}_we_{i}')
            windows.append((f"{ws}~{we}", pd.Timestamp(ws), pd.Timestamp(we), ws.year))
    return windows

COLORS = ['#006D36','#2D8C3C','#3498DB','#E67E22','#9B59B6','#E74C8B','#1ABC9C','#E74C3C','#F39C12','#8E44AD','#2980B9','#27AE60']

SPOT_COL_MAP = {
    0: 'Date', 1: 'Brent', 2: 'WTI',
    3: 'CP_C3', 4: 'CP_C4', 5: 'FEI_C3', 6: 'FEI_C4',
    7: 'Purity_C2', 8: 'MB_NonTET_C3', 9: 'MB_NonTET_nC4', 10: 'MB_NonTET_iC4',
    11: 'MB_TET_C3', 12: 'MB_TET_nC4', 13: 'MB_TET_iC4',
    14: 'ARA_C3', 15: 'ARA_C4', 16: 'Baltic_Freight', 17: 'RIM_Freight',
    18: 'MOPJ', 19: 'MOPS_FO_380', 20: 'JKM'
}

@st.cache_data(show_spinner="Spot 데이터 로딩 중...")
def load_spot_data(f):
    df = pd.read_excel(f, header=None, skiprows=2)
    # Keep only mapped columns
    cols_to_keep = {i: name for i, name in SPOT_COL_MAP.items() if i < df.shape[1]}
    df = df[[i for i in cols_to_keep.keys()]].copy()
    df.columns = [cols_to_keep[i] for i in cols_to_keep.keys()]
    # Parse date
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])
    # Convert numeric, replace 0 with NaN, interpolate
    for c in df.columns:
        if c == 'Date': continue
        df[c] = pd.to_numeric(df[c], errors='coerce')
        df.loc[df[c] == 0, c] = np.nan
    df = df.sort_values('Date').reset_index(drop=True)
    df = df.interpolate(method='linear', limit_direction='forward')
    return df

# ═══ MAIN ═══
st.title("📊 Fundamental Driver Correlation Dashboard")
st.caption("SK Gas Trading AX — Position Play AI — v3.0")

with st.sidebar:
    st.header("📁 데이터 업로드")
    st.subheader("1️⃣ 인덱스 가격 데이터")
    idx_files = st.file_uploader("인덱스 파일", type=['csv','xlsx','xls'], accept_multiple_files=True, key='idx_upload')
    st.subheader("2️⃣ Fundamental Driver 데이터")
    fund_files = st.file_uploader("Fundamental 파일", type=['csv','xlsx','xls'], accept_multiple_files=True, key='fund_upload')
    st.subheader("3️⃣ Daily Spot 가격 데이터")
    st.caption("Daily Spot Price (멀티 헤더 Excel)")
    spot_file = st.file_uploader("Daily Spot 파일", type=['xlsx','xls'], key='spot_upload')
    if st.button("🔄 캐시 초기화 (파일 재업로드 시)", key='clear_cache'):
        st.cache_data.clear()
        st.rerun()

idx_df = fund_df = spot_df = None
if idx_files:
    try:
        idx_df = load_index_data(idx_files)
        st.sidebar.success(f"✅ 인덱스: {len(idx_df):,}행, {idx_df['Index명'].nunique()}개")
    except Exception as e: st.sidebar.error(f"인덱스 오류: {e}")
if fund_files:
    try:
        fund_df = load_fund_data(fund_files)
        st.sidebar.success(f"✅ Fundamental: {len(fund_df):,}행, {fund_df['Fundamental Driver'].nunique()}개 Driver")
        with st.sidebar.expander("📋 인덱스별 Driver"):
            for idx in sorted(fund_df['Index'].unique()):
                st.caption(f"**{idx}**: {fund_df[fund_df['Index']==idx]['Fundamental Driver'].nunique()}개")
    except Exception as e: st.sidebar.error(f"Fundamental 오류: {e}")
if spot_file:
    try:
        spot_df = load_spot_data(spot_file)
        spot_cols = [c for c in spot_df.columns if c != 'Date']
        st.sidebar.success(f"✅ Spot: {len(spot_df):,}행, {len(spot_cols)}개 인덱스")
    except Exception as e: st.sidebar.error(f"Spot 오류: {e}")

if idx_df is None and fund_df is None and spot_df is None:
    st.info("👈 사이드바에서 데이터를 업로드하세요.")
    st.stop()

tab_spread, tab_fund_curve, tab1, tab2, tab3, tab_data = st.tabs([
    "📊 Spread Curve",
    "📈 Fundamental Curve",
    "📈 Output 1: Driver 평균 Correlation",
    "📉 Output 2: Spread vs Driver Rolling",
    "📊 Output 3: Index vs Driver Rolling",
    "🔍 데이터 미리보기"
])

# ═══ DATA PREVIEW ═══
with tab_data:
    c1,c2,c3 = st.columns(3)
    with c1:
        st.subheader("인덱스 가격")
        if idx_df is not None: st.dataframe(idx_df.head(100), use_container_width=True, height=400)
    with c2:
        st.subheader("Fundamental Driver")
        if fund_df is not None: st.dataframe(fund_df.head(100), use_container_width=True, height=400)
    with c3:
        st.subheader("Daily Spot")
        if spot_df is not None: st.dataframe(spot_df.head(100), use_container_width=True, height=400)

# ═══ SPREAD & INDEX CURVE ═══
with tab_spread:
    st.subheader("Spread & Index & Spot Curve 시각화")
    st.caption("Spread 커브 + 각 Index 월물 커브 + Daily Spot 커브 | Raw / Normalized 전환 가능")
    if idx_df is None:
        st.warning("인덱스 가격 데이터를 업로드하세요.")
    else:
        all_idx = sorted(idx_df['Index명'].unique())
        cs1,cs2 = st.columns(2)
        with cs1:
            st.markdown("**Index 1**"); sp_idx1 = st.selectbox("Index 1", all_idx, key='sp_i1')
            sp_y1, sp_m1 = contract_month_selector('sp_cm1', label="월물 1")
        with cs2:
            st.markdown("**Index 2**"); sp_idx2 = st.selectbox("Index 2", all_idx, key='sp_i2', index=min(1,len(all_idx)-1))
            sp_y2, sp_m2 = contract_month_selector('sp_cm2', label="월물 2")

        # Spot 인덱스 매핑
        spot_idx1_col = spot_idx2_col = None
        if spot_df is not None:
            st.markdown("---")
            st.markdown("**Daily Spot 매핑 (각 Index에 대응하는 Spot 선택)**")
            spot_available = ['(없음)'] + [c for c in spot_df.columns if c != 'Date']
            sc1, sc2 = st.columns(2)
            with sc1:
                spot_idx1_sel = st.selectbox(f"Index 1 ({sp_idx1}) → Spot", spot_available, key='sp_spot1')
                if spot_idx1_sel != '(없음)': spot_idx1_col = spot_idx1_sel
            with sc2:
                spot_idx2_sel = st.selectbox(f"Index 2 ({sp_idx2}) → Spot", spot_available, key='sp_spot2')
                if spot_idx2_sel != '(없음)': spot_idx2_col = spot_idx2_sel

        st.markdown("---")
        sp_windows = window_selector_simple('sp', idx_df=idx_df)

        # Raw / Normalized 토글
        norm_mode = st.radio("표시 모드", ["Raw (원래 값)", "Normalized (Z-score: 평균=0, 표준편차=1)"],
                             horizontal=True, key='sp_norm')
        is_normalized = norm_mode.startswith("Normalized")

        if st.button("📊 커브 그리기", key='sp_run') and sp_windows:
            sp_data = {}
            idx1_data = {}
            idx2_data = {}
            spot1_data = {}  # Index 1 Spot
            spot2_data = {}  # Index 2 Spot
            spot_spread_data = {}  # Spot1 - Spot2
            decomp_rows = []  # Variance decomposition
            direction_rows = []  # Direction analysis

            for i,(wl,ws,we,by) in enumerate(sp_windows):
                m1 = resolve_contract_month(by, sp_y1, sp_m1)
                m2 = resolve_contract_month(by, sp_y2, sp_m2)

                # Index 1 Forward curve
                df1 = idx_df[(idx_df['Index명']==sp_idx1)&(idx_df['월물']==m1)&(idx_df['기준일자']>=ws)&(idx_df['기준일자']<=we)].sort_values('기준일자')
                if len(df1)>0: idx1_data[wl] = (df1[['기준일자','Value']].copy(), m1)

                # Index 2 Forward curve
                df2 = idx_df[(idx_df['Index명']==sp_idx2)&(idx_df['월물']==m2)&(idx_df['기준일자']>=ws)&(idx_df['기준일자']<=we)].sort_values('기준일자')
                if len(df2)>0: idx2_data[wl] = (df2[['기준일자','Value']].copy(), m2)

                # Spread curve (Forward)
                sp = get_spread_curve(idx_df, sp_idx1, m1, sp_idx2, m2, ws, we)
                if len(sp)>0: sp_data[wl] = sp

                # Index 1 Spot curve
                if spot_df is not None and spot_idx1_col:
                    ss1 = spot_df[(spot_df['Date']>=ws)&(spot_df['Date']<=we)][['Date',spot_idx1_col]].dropna().sort_values('Date')
                    if len(ss1)>0:
                        spot1_data[wl] = ss1.rename(columns={spot_idx1_col:'Value','Date':'기준일자'})

                # Index 2 Spot curve
                if spot_df is not None and spot_idx2_col:
                    ss2 = spot_df[(spot_df['Date']>=ws)&(spot_df['Date']<=we)][['Date',spot_idx2_col]].dropna().sort_values('Date')
                    if len(ss2)>0:
                        spot2_data[wl] = ss2.rename(columns={spot_idx2_col:'Value','Date':'기준일자'})

                # Spot Spread (Spot1 - Spot2)
                if spot_df is not None and spot_idx1_col is not None and spot_idx2_col is not None:
                    try:
                        ss_both = spot_df[(spot_df['Date']>=ws)&(spot_df['Date']<=we)][['Date',spot_idx1_col,spot_idx2_col]].dropna().sort_values('Date').reset_index(drop=True)
                        if len(ss_both)>0:
                            ss_both = ss_both.rename(columns={'Date':'기준일자'})
                            ss_both['Spread'] = ss_both[spot_idx1_col].values - ss_both[spot_idx2_col].values
                            spot_spread_data[wl] = ss_both[['기준일자','Spread']]
                    except Exception:
                        pass

            def normalize_series(values):
                """Z-score normalize: mean=0, std=1"""
                v = np.array(values, dtype=float)
                if len(v) < 2: return v
                m, s = np.nanmean(v), np.nanstd(v)
                if s == 0: return v - m
                return (v - m) / s

            def make_overlay(data_dict, title, y_label, chart_num, is_spread=False):
                """Generic overlay chart builder"""
                st.markdown(f"### {chart_num}. {title}")
                fig = go.Figure()
                for i,(wl, item) in enumerate(data_dict.items()):
                    if is_spread:
                        df_plot = item.reset_index(drop=True)
                        y_vals = df_plot['Spread'].values
                        dates = df_plot['기준일자']
                    else:
                        if isinstance(item, tuple):
                            df_plot = item[0].reset_index(drop=True)
                        else:
                            df_plot = item.reset_index(drop=True)
                        y_vals = df_plot['Value'].values
                        dates = df_plot['기준일자']

                    if is_normalized:
                        y_vals = normalize_series(y_vals)

                    days = list(range(1, len(y_vals)+1))
                    fig.add_trace(go.Scatter(
                        x=days, y=y_vals, mode='lines',
                        name=wl, line=dict(color=COLORS[i%len(COLORS)], width=2.5),
                        hovertemplate=f"{wl}<br>Day %{{x}}<br>{y_label}: %{{y:.2f}}<br>Date: %{{customdata}}<extra></extra>",
                        customdata=dates.dt.strftime('%Y-%m-%d')
                    ))
                y_title = y_label if not is_normalized else f"{y_label} (Z-score)"
                fig.update_layout(
                    title=title, xaxis_title="Trading Day (윈도우 시작 기준)", yaxis_title=y_title,
                    height=450, template='plotly_white',
                    legend=dict(orientation='h', y=-0.15), hovermode='x unified'
                )
                if is_spread and not is_normalized:
                    fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
                if is_normalized:
                    fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.4)
                st.plotly_chart(fig, use_container_width=True)

            # ── Overlay 1: Forward Spread ──
            if sp_data:
                make_overlay(sp_data,
                             f"Forward Spread: {sp_idx1}(Y+{sp_y1} {sp_m1}월) − {sp_idx2}(Y+{sp_y2} {sp_m2}월)",
                             "Spread", "1", is_spread=True)

            # ── Overlay 2: Spot Spread ──
            if spot_spread_data:
                make_overlay(spot_spread_data,
                             f"Spot Spread: {spot_idx1_col} − {spot_idx2_col}",
                             "Spread", "2", is_spread=True)

            # ── Overlay 3: Index 1 Forward ──
            if idx1_data:
                make_overlay(idx1_data,
                             f"{sp_idx1} Forward (Y+{sp_y1} {sp_m1}월)",
                             "Price", "3")

            # ── Overlay 4: Index 1 Spot ──
            if spot1_data:
                make_overlay(spot1_data,
                             f"{sp_idx1} Spot ({spot_idx1_col})",
                             "Price", "4")

            # ── Overlay 5: Index 2 Forward ──
            if idx2_data:
                make_overlay(idx2_data,
                             f"{sp_idx2} Forward (Y+{sp_y2} {sp_m2}월)",
                             "Price", "5")

            # ── Overlay 6: Index 2 Spot ──
            if spot2_data:
                make_overlay(spot2_data,
                             f"{sp_idx2} Spot ({spot_idx2_col})",
                             "Price", "6")

            # ═══ Spread 기여도 분해 분석 ═══
            if idx1_data and idx2_data:
                st.markdown("---")
                st.markdown("### 📊 Spread 기여도 분해 (Variance Decomposition)")
                st.caption("ΔSpread = ΔIndex2 − ΔIndex1. 변동 크기 + 방향 + 연동성을 종합 분석합니다.")

                decomp_rows = []
                direction_rows = []

                for wl in sp_data.keys():
                    if wl not in idx1_data or wl not in idx2_data:
                        continue

                    df1 = idx1_data[wl][0].copy().reset_index(drop=True)
                    df2 = idx2_data[wl][0].copy().reset_index(drop=True)

                    merged = pd.merge(df1.rename(columns={'Value':'V1'}),
                                       df2.rename(columns={'Value':'V2'}),
                                       on='기준일자', how='inner').sort_values('기준일자')
                    if len(merged) < 5:
                        continue

                    # Daily changes
                    dV1 = merged['V1'].diff().dropna().values
                    dV2 = merged['V2'].diff().dropna().values
                    dSpread = dV2 - dV1

                    var_dV1 = np.var(dV1)
                    var_dV2 = np.var(dV2)
                    var_dSpread = np.var(dSpread)
                    std_v1 = np.std(dV1)
                    std_v2 = np.std(dV2)

                    # Price level stats
                    std_p1 = np.std(merged['V1'].values)
                    std_p2 = np.std(merged['V2'].values)
                    var_p1 = np.var(merged['V1'].values)
                    var_p2 = np.var(merged['V2'].values)
                    mean_p1 = np.mean(merged['V1'].values)
                    mean_p2 = np.mean(merged['V2'].values)

                    # Contribution ratio
                    total_var = var_dV1 + var_dV2
                    contrib_v1_pct = var_dV1 / total_var * 100 if total_var > 0 else 50.0
                    contrib_v2_pct = var_dV2 / total_var * 100 if total_var > 0 else 50.0

                    # Price level contribution
                    total_var_p = var_p1 + var_p2
                    contrib_p1_pct = var_p1 / total_var_p * 100 if total_var_p > 0 else 50.0
                    contrib_p2_pct = var_p2 / total_var_p * 100 if total_var_p > 0 else 50.0

                    # |ΔV1| > |ΔV2| ratio
                    abs_v1_bigger = np.sum(np.abs(dV1) > np.abs(dV2))
                    abs_v1_ratio = abs_v1_bigger / len(dV1) * 100

                    decomp_rows.append({
                        'Window': wl,
                        'Trading Days': len(dV1),
                        f'Std(Δ{sp_idx1})': round(std_v1, 4),
                        f'Std(Δ{sp_idx2})': round(std_v2, 4),
                        f'{sp_idx1} ΔPrice 변동 비중': f"{contrib_v1_pct:.1f}%",
                        f'{sp_idx2} ΔPrice 변동 비중': f"{contrib_v2_pct:.1f}%",
                        f'|Δ{sp_idx1}|>|Δ{sp_idx2}| 비율': f"{abs_v1_ratio:.1f}%",
                        f'Std({sp_idx1})': round(std_p1, 4),
                        f'Std({sp_idx2})': round(std_p2, 4),
                        f'Var({sp_idx1})': round(var_p1, 4),
                        f'Var({sp_idx2})': round(var_p2, 4),
                        f'{sp_idx1} Price 변동 비중': f"{contrib_p1_pct:.1f}%",
                        f'{sp_idx2} Price 변동 비중': f"{contrib_p2_pct:.1f}%",
                    })

                    # ── 방향 분석 ──
                    v1_start = merged['V1'].iloc[0]
                    v1_end = merged['V1'].iloc[-1]
                    v2_start = merged['V2'].iloc[0]
                    v2_end = merged['V2'].iloc[-1]
                    total_dV1 = v1_end - v1_start
                    total_dV2 = v2_end - v2_start
                    total_dSpread = total_dV2 - total_dV1
                    spread_start = v2_start - v1_start
                    spread_end = v2_end - v1_end

                    # Spread 확대일 / 축소일 분리
                    expand_mask = dSpread > 0
                    shrink_mask = dSpread < 0
                    n_expand = np.sum(expand_mask)
                    n_shrink = np.sum(shrink_mask)
                    avg_dV1_expand = np.mean(dV1[expand_mask]) if n_expand > 0 else 0
                    avg_dV2_expand = np.mean(dV2[expand_mask]) if n_expand > 0 else 0
                    avg_dV1_shrink = np.mean(dV1[shrink_mask]) if n_shrink > 0 else 0
                    avg_dV2_shrink = np.mean(dV2[shrink_mask]) if n_shrink > 0 else 0

                    # 방향성 상관계수
                    corr_v1_spread = np.corrcoef(dV1, dSpread)[0, 1] if len(dV1) >= 3 else np.nan
                    corr_v2_spread = np.corrcoef(dV2, dSpread)[0, 1] if len(dV2) >= 3 else np.nan

                    direction_rows.append({
                        'Window': wl,
                        f'{sp_idx1} Start→End': f"{v1_start:.2f}→{v1_end:.2f} (Δ{total_dV1:+.2f})",
                        f'{sp_idx2} Start→End': f"{v2_start:.2f}→{v2_end:.2f} (Δ{total_dV2:+.2f})",
                        f'Spread Start→End': f"{spread_start:.2f}→{spread_end:.2f} (Δ{total_dSpread:+.2f})",
                        '확대일 수': int(n_expand),
                        '축소일 수': int(n_shrink),
                        f'확대일 평균Δ{sp_idx1}': round(avg_dV1_expand, 4),
                        f'확대일 평균Δ{sp_idx2}': round(avg_dV2_expand, 4),
                        f'축소일 평균Δ{sp_idx1}': round(avg_dV1_shrink, 4),
                        f'축소일 평균Δ{sp_idx2}': round(avg_dV2_shrink, 4),
                        f'Corr(Δ{sp_idx1}, ΔSpread)': round(corr_v1_spread, 4) if not np.isnan(corr_v1_spread) else None,
                        f'Corr(Δ{sp_idx2}, ΔSpread)': round(corr_v2_spread, 4) if not np.isnan(corr_v2_spread) else None,
                    })

                if decomp_rows:
                    # ── A-1. 일별 변화량(ΔPrice) 변동 크기 ──
                    st.markdown("#### A-1. 일별 변화량(ΔPrice) 변동 크기")
                    st.caption("Std(ΔPrice): 일별 가격 변화의 표준편차. 매일 얼마나 크게 움직이는가.")
                    decomp_df = pd.DataFrame(decomp_rows)
                    delta_cols = ['Window', 'Trading Days',
                                  f'Std(Δ{sp_idx1})', f'Std(Δ{sp_idx2})',
                                  f'{sp_idx1} ΔPrice 변동 비중', f'{sp_idx2} ΔPrice 변동 비중',
                                  f'|Δ{sp_idx1}|>|Δ{sp_idx2}| 비율']
                    st.dataframe(decomp_df[delta_cols], use_container_width=True)

                    fig_dc = go.Figure()
                    fig_dc.add_trace(go.Bar(
                        x=decomp_df['Window'],
                        y=[float(v.replace('%','')) for v in decomp_df[f'{sp_idx1} ΔPrice 변동 비중']],
                        name=sp_idx1, marker_color='#006D36',
                        text=decomp_df[f'{sp_idx1} ΔPrice 변동 비중'], textposition='outside'
                    ))
                    fig_dc.add_trace(go.Bar(
                        x=decomp_df['Window'],
                        y=[float(v.replace('%','')) for v in decomp_df[f'{sp_idx2} ΔPrice 변동 비중']],
                        name=sp_idx2, marker_color='#3498DB',
                        text=decomp_df[f'{sp_idx2} ΔPrice 변동 비중'], textposition='outside'
                    ))
                    fig_dc.update_layout(
                        title=f"Var(ΔPrice) 비중: {sp_idx1} vs {sp_idx2}",
                        barmode='group', height=380, template='plotly_white',
                        yaxis_title="변동 비중 (%)", legend=dict(orientation='h', y=-0.15)
                    )
                    fig_dc.add_hline(y=50, line_dash="dash", line_color="gray", opacity=0.4)
                    st.plotly_chart(fig_dc, use_container_width=True)

                    avg_v1 = np.mean([float(v.replace('%','')) for v in decomp_df[f'{sp_idx1} ΔPrice 변동 비중']])
                    avg_v2 = np.mean([float(v.replace('%','')) for v in decomp_df[f'{sp_idx2} ΔPrice 변동 비중']])
                    avg_ratio = np.mean([float(v.replace('%','')) for v in decomp_df[f'|Δ{sp_idx1}|>|Δ{sp_idx2}| 비율']])
                    mc1, mc2, mc3 = st.columns(3)
                    mc1.metric(f"{sp_idx1} 평균 ΔPrice 비중", f"{avg_v1:.1f}%")
                    mc2.metric(f"{sp_idx2} 평균 ΔPrice 비중", f"{avg_v2:.1f}%")
                    mc3.metric(f"|Δ{sp_idx1}|>|Δ{sp_idx2}| 평균", f"{avg_ratio:.1f}%")

                    # ── A-2. 가격 수준(Price) 변동 크기 ──
                    st.markdown("---")
                    st.markdown("#### A-2. 가격 수준(Price) 변동 크기")
                    st.caption("Std(Price), Var(Price): 윈도우 내 가격 수준이 얼마나 넓은 범위에서 움직였는가.")
                    price_cols = ['Window',
                                  f'Std({sp_idx1})', f'Std({sp_idx2})',
                                  f'Var({sp_idx1})', f'Var({sp_idx2})',
                                  f'{sp_idx1} Price 변동 비중', f'{sp_idx2} Price 변동 비중']
                    st.dataframe(decomp_df[price_cols], use_container_width=True)

                    fig_pv = go.Figure()
                    fig_pv.add_trace(go.Bar(
                        x=decomp_df['Window'],
                        y=[float(v.replace('%','')) for v in decomp_df[f'{sp_idx1} Price 변동 비중']],
                        name=sp_idx1, marker_color='#006D36',
                        text=decomp_df[f'{sp_idx1} Price 변동 비중'], textposition='outside'
                    ))
                    fig_pv.add_trace(go.Bar(
                        x=decomp_df['Window'],
                        y=[float(v.replace('%','')) for v in decomp_df[f'{sp_idx2} Price 변동 비중']],
                        name=sp_idx2, marker_color='#3498DB',
                        text=decomp_df[f'{sp_idx2} Price 변동 비중'], textposition='outside'
                    ))
                    fig_pv.update_layout(
                        title=f"Var(Price) 비중: {sp_idx1} vs {sp_idx2}",
                        barmode='group', height=380, template='plotly_white',
                        yaxis_title="변동 비중 (%)", legend=dict(orientation='h', y=-0.15)
                    )
                    fig_pv.add_hline(y=50, line_dash="dash", line_color="gray", opacity=0.4)
                    st.plotly_chart(fig_pv, use_container_width=True)

                    avg_p1 = np.mean([float(v.replace('%','')) for v in decomp_df[f'{sp_idx1} Price 변동 비중']])
                    avg_p2 = np.mean([float(v.replace('%','')) for v in decomp_df[f'{sp_idx2} Price 변동 비중']])
                    mc1, mc2 = st.columns(2)
                    mc1.metric(f"{sp_idx1} 평균 Price 비중", f"{avg_p1:.1f}%")
                    mc2.metric(f"{sp_idx2} 평균 Price 비중", f"{avg_p2:.1f}%")

                if direction_rows:
                    dir_df = pd.DataFrame(direction_rows)

                    # ── B. 구간 내 총 변화량 분해 ──
                    st.markdown("---")
                    st.markdown("#### B. 구간 내 총 변화량 분해")
                    st.caption("윈도우 Start→End 가격 변화. Spread 확대/축소의 실제 원인을 보여줍니다.")
                    display_cols_b = ['Window',
                                      f'{sp_idx1} Start→End', f'{sp_idx2} Start→End', f'Spread Start→End']
                    st.dataframe(dir_df[display_cols_b], use_container_width=True)

                    # ── C. Spread 확대일/축소일 분리 ──
                    st.markdown("---")
                    st.markdown("#### C. Spread 확대일/축소일 분리 분석")
                    st.caption("Spread가 확대된 날 vs 축소된 날, 각 Index의 평균 일별 변화")
                    expand_cols = ['Window', '확대일 수', '축소일 수',
                                   f'확대일 평균Δ{sp_idx1}', f'확대일 평균Δ{sp_idx2}',
                                   f'축소일 평균Δ{sp_idx1}', f'축소일 평균Δ{sp_idx2}']
                    st.dataframe(dir_df[expand_cols], use_container_width=True)

                    # Chart: expand day analysis
                    fig_exp = go.Figure()
                    fig_exp.add_trace(go.Bar(x=dir_df['Window'], y=dir_df[f'확대일 평균Δ{sp_idx1}'],
                                              name=f"확대일 Δ{sp_idx1}", marker_color='#006D36'))
                    fig_exp.add_trace(go.Bar(x=dir_df['Window'], y=dir_df[f'확대일 평균Δ{sp_idx2}'],
                                              name=f"확대일 Δ{sp_idx2}", marker_color='#3498DB'))
                    fig_exp.add_trace(go.Bar(x=dir_df['Window'], y=dir_df[f'축소일 평균Δ{sp_idx1}'],
                                              name=f"축소일 Δ{sp_idx1}", marker_color='#2D8C3C', opacity=0.6))
                    fig_exp.add_trace(go.Bar(x=dir_df['Window'], y=dir_df[f'축소일 평균Δ{sp_idx2}'],
                                              name=f"축소일 Δ{sp_idx2}", marker_color='#85C1E9', opacity=0.6))
                    fig_exp.update_layout(
                        title="Spread 확대일/축소일: 각 Index 평균 Δ",
                        barmode='group', height=400, template='plotly_white',
                        yaxis_title="평균 일별 변화", legend=dict(orientation='h', y=-0.18)
                    )
                    fig_exp.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.4)
                    st.plotly_chart(fig_exp, use_container_width=True)

                    # ── D. 방향성 상관계수 ──
                    st.markdown("---")
                    st.markdown("#### D. 방향성 상관계수")
                    st.caption("Corr(ΔIndex, ΔSpread): 음수가 클수록 해당 Index 하락 → Spread 확대 연동이 강함")
                    corr_v1_col = f'Corr(Δ{sp_idx1}, ΔSpread)'
                    corr_v2_col = f'Corr(Δ{sp_idx2}, ΔSpread)'
                    st.dataframe(dir_df[['Window', corr_v1_col, corr_v2_col]], use_container_width=True)

                    fig_corr = go.Figure()
                    fig_corr.add_trace(go.Bar(x=dir_df['Window'], y=dir_df[corr_v1_col],
                                               name=f"Corr(Δ{sp_idx1}, ΔSpread)", marker_color='#006D36',
                                               text=dir_df[corr_v1_col].apply(lambda x: f"{x:.3f}" if pd.notna(x) else ""), textposition='outside'))
                    fig_corr.add_trace(go.Bar(x=dir_df['Window'], y=dir_df[corr_v2_col],
                                               name=f"Corr(Δ{sp_idx2}, ΔSpread)", marker_color='#3498DB',
                                               text=dir_df[corr_v2_col].apply(lambda x: f"{x:.3f}" if pd.notna(x) else ""), textposition='outside'))
                    fig_corr.update_layout(
                        title="방향성 상관: ΔIndex vs ΔSpread",
                        barmode='group', height=400, template='plotly_white',
                        yaxis_title="Pearson Correlation", yaxis_range=[-1.1, 1.1],
                        legend=dict(orientation='h', y=-0.15)
                    )
                    fig_corr.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.4)
                    st.plotly_chart(fig_corr, use_container_width=True)

                    # Summary
                    avg_corr_v1 = dir_df[corr_v1_col].mean()
                    avg_corr_v2 = dir_df[corr_v2_col].mean()
                    mc1, mc2 = st.columns(2)
                    mc1.metric(f"평균 Corr(Δ{sp_idx1}, ΔSpread)", f"{avg_corr_v1:.3f}")
                    mc2.metric(f"평균 Corr(Δ{sp_idx2}, ΔSpread)", f"{avg_corr_v2:.3f}")

                    # 종합 판정
                    dominant_size = sp_idx1 if avg_v1 > avg_v2 else sp_idx2
                    dominant_dir = sp_idx1 if abs(avg_corr_v1) > abs(avg_corr_v2) else sp_idx2
                    if dominant_size == dominant_dir:
                        st.success(f"📌 **{dominant_size}**가 Spread 변동의 주요인입니다. "
                                   f"변동 크기({avg_v1:.1f}% vs {avg_v2:.1f}%)와 "
                                   f"방향성 연동(|r|={abs(avg_corr_v1 if dominant_dir==sp_idx1 else avg_corr_v2):.3f}) 모두 일관됩니다.")
                    else:
                        st.warning(f"⚠️ 크기 주도({dominant_size}: {max(avg_v1,avg_v2):.1f}%)와 "
                                   f"방향 주도({dominant_dir}: |r|={max(abs(avg_corr_v1),abs(avg_corr_v2)):.3f})가 다릅니다. "
                                   f"윈도우별로 개별 확인이 필요합니다.")

            # ── Excel 다운로드 ──
            has_data = sp_data or idx1_data or idx2_data or spot1_data or spot2_data or spot_spread_data
            if has_data:
                sheets = {}
                for wl, d in sp_data.items():
                    sheets[safe_sheet_name(f"FwdSpread_{wl}")] = d
                for wl, d in spot_spread_data.items():
                    sheets[safe_sheet_name(f"SpotSpread_{wl}")] = d
                for wl, (df,m) in idx1_data.items():
                    sheets[safe_sheet_name(f"{sp_idx1}Fwd_{wl}")] = df
                for wl, df in spot1_data.items():
                    sheets[safe_sheet_name(f"{spot_idx1_col}Spot_{wl}")] = df
                for wl, (df,m) in idx2_data.items():
                    sheets[safe_sheet_name(f"{sp_idx2}Fwd_{wl}")] = df
                for wl, df in spot2_data.items():
                    sheets[safe_sheet_name(f"{spot_idx2_col}Spot_{wl}")] = df
                if decomp_rows:
                    sheets['Variance Decomposition'] = pd.DataFrame(decomp_rows)
                if direction_rows:
                    sheets['Direction Analysis'] = pd.DataFrame(direction_rows)
                st.download_button("📥 엑셀 (전체 커브 + 분해)", data=make_excel(sheets),
                                   file_name="Spread_Index_Spot_Curves.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("해당 조합에 데이터가 없습니다.")

# ═══ FUNDAMENTAL CURVE ═══
with tab_fund_curve:
    st.subheader("Fundamental Curve 시각화")
    st.caption("선택한 Driver의 커브를 윈도우별로 오버레이하여 비교합니다")

    if fund_df is None:
        st.warning("Fundamental Driver 데이터를 업로드하세요.")
    else:
        st.markdown("**Fundamental Driver 선택 (복수)**")
        sd_fc = driver_multiselect(fund_df, key_prefix='fc', label="Driver 선택")

        st.markdown("---")
        st.markdown("**윈도우 설정 (연-월-일)**")
        w_fc = window_selector_simple('fc', fund_df=fund_df)

        # 오버레이 모드 선택
        overlay_mode = st.radio("오버레이 방식", ["윈도우별 (같은 Driver의 연도별 비교)", "Driver별 (같은 윈도우 내 Driver 비교)"],
                                horizontal=True, key='fc_overlay')

        if st.button("📈 커브 그리기", key='fc_run') and sd_fc and w_fc:

            if overlay_mode == "윈도우별 (같은 Driver의 연도별 비교)":
                # Driver별로 차트 1개씩. 각 차트 안에 윈도우별 커브 오버레이.
                for fi, dn in sd_fc:
                    sf = fund_df[(fund_df['Index']==fi)&(fund_df['Fundamental Driver']==dn)]
                    drv_label = f"[{fi}] {dn}"
                    fig_fc = go.Figure()
                    export_data = {}

                    for i, (wl, ws, we, _) in enumerate(w_fc):
                        ws_exp, we_exp = expand_to_full_months(ws, we)
                        sub = sf[(sf['Date']>=ws_exp)&(sf['Date']<=we_exp)].sort_values('Date')
                        if len(sub) == 0:
                            continue

                        # 월별 집계
                        monthly = sub.groupby(sub['Date'].dt.to_period('M'))['Value'].mean()
                        x_labels = [f"M{j+1}" for j in range(len(monthly))]
                        fig_fc.add_trace(go.Scatter(
                            x=x_labels, y=monthly.values,
                            mode='lines+markers', name=wl,
                            line=dict(color=COLORS[i % len(COLORS)], width=2.5),
                            marker=dict(size=7)
                        ))
                        export_data[wl] = pd.DataFrame({
                            'Month Position': x_labels,
                            'Period': [str(p) for p in monthly.index],
                            'Value': monthly.values
                        })

                    if len(fig_fc.data) > 0:
                        fig_fc.update_layout(
                            title=f"{drv_label} — 윈도우별 오버레이",
                            xaxis_title="Month Position (윈도우 시작 기준)",
                            yaxis_title="Value",
                            height=450, template='plotly_white',
                            legend=dict(orientation='h', y=-0.15)
                        )
                        st.plotly_chart(fig_fc, use_container_width=True)
                    else:
                        st.info(f"{drv_label}: 해당 윈도우에 데이터가 없습니다.")

                # 엑셀 다운로드 (마지막 Driver 기준)
                if export_data:
                    sheets_fc = {safe_sheet_name(k): v for k, v in export_data.items()}
                    st.download_button("📥 엑셀", data=make_excel(sheets_fc),
                                       file_name="Fundamental_Curve.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            else:
                # 윈도우별로 차트 1개씩. 각 차트 안에 Driver별 커브 비교.
                for wl, ws, we, _ in w_fc:
                    fig_fc2 = go.Figure()
                    ws_exp, we_exp = expand_to_full_months(ws, we)

                    for i, (fi, dn) in enumerate(sd_fc):
                        sf = fund_df[(fund_df['Index']==fi)&(fund_df['Fundamental Driver']==dn)]
                        sub = sf[(sf['Date']>=ws_exp)&(sf['Date']<=we_exp)].sort_values('Date')
                        if len(sub) == 0:
                            continue

                        monthly = sub.groupby(sub['Date'].dt.to_period('M'))['Value'].mean()
                        fig_fc2.add_trace(go.Scatter(
                            x=[str(p) for p in monthly.index],
                            y=monthly.values,
                            mode='lines+markers', name=f"[{fi}] {dn}",
                            line=dict(color=COLORS[i % len(COLORS)], width=2.5),
                            marker=dict(size=7),
                            yaxis='y' if i == 0 else f'y{i+1}'
                        ))

                    if len(fig_fc2.data) > 0:
                        # 다중 Y축 (Driver별 스케일이 다를 수 있으므로)
                        layout_update = dict(
                            title=f"{wl} — Driver별 비교",
                            xaxis_title="Period",
                            height=450, template='plotly_white',
                            legend=dict(orientation='h', y=-0.15)
                        )
                        if len(sd_fc) <= 2:
                            layout_update['yaxis'] = dict(title=f"[{sd_fc[0][0]}] {sd_fc[0][1]}" if sd_fc else "Value")
                            if len(sd_fc) == 2:
                                layout_update['yaxis2'] = dict(title=f"[{sd_fc[1][0]}] {sd_fc[1][1]}",
                                                                overlaying='y', side='right')
                        else:
                            layout_update['yaxis'] = dict(title="Value")
                        fig_fc2.update_layout(**layout_update)
                        st.plotly_chart(fig_fc2, use_container_width=True)
                    else:
                        st.info(f"{wl}: 해당 윈도우에 데이터가 없습니다.")

# ═══ OUTPUT 1: DRIVER AVG CORRELATION (SIMPLIFIED) ═══
with tab1:
    st.subheader("Fundamental Driver별 평균 Pearson Correlation")
    st.caption("모든 윈도우 조합의 평균 Correlation만 Driver별로 요약합니다")
    if fund_df is None:
        st.warning("Fundamental 데이터를 업로드하세요.")
    else:
        st.markdown("**Fundamental Driver 선택 (복수)**")
        sd1 = driver_multiselect(fund_df, key_prefix='t1')
        st.markdown("---"); st.markdown("**윈도우 설정**")
        w1 = window_selector_simple('t1', fund_df=fund_df)
        if st.button("🔍 분석", key='t1_run') and sd1 and w1:
            rows = []
            for fi, dn in sd1:
                sf = fund_df[(fund_df['Index']==fi)&(fund_df['Fundamental Driver']==dn)]
                curves = {}
                for wl,ws,we,_ in w1:
                    ws_exp, we_exp = expand_to_full_months(ws, we)
                    sub = sf[(sf['Date']>=ws_exp)&(sf['Date']<=we_exp)]
                    if len(sub)>=3:
                        m = sub.groupby(sub['Date'].dt.to_period('M'))['Value'].mean()
                        if len(m)>=3: curves[wl] = m
                rs = []
                if len(curves)>=2:
                    for k1,k2 in combinations(curves.keys(),2):
                        v1,v2 = curves[k1].values, curves[k2].values
                        n = min(len(v1),len(v2))
                        if n>=3:
                            r = pearson_corr(v1[:n],v2[:n])
                            if not np.isnan(r): rs.append(r)
                rows.append({'Index': fi, 'Fundamental Driver': dn,
                             'Average r': round(np.mean(rs),4) if rs else None,
                             'Windows': len(curves), 'Pairs': len(rs)})
            if rows:
                sdf = pd.DataFrame(rows)
                st.dataframe(sdf, use_container_width=True)
                fig1 = go.Figure()
                fig1.add_trace(go.Bar(x=[f"{r['Index']} - {r['Fundamental Driver']}" for _,r in sdf.iterrows()],
                                      y=sdf['Average r'], marker_color=[COLORS[i%len(COLORS)] for i in range(len(sdf))],
                                      text=sdf['Average r'].apply(lambda x: f"{x:.3f}" if pd.notna(x) else "N/A"), textposition='outside'))
                fig1.update_layout(title="Driver별 평균 Pearson Correlation", height=420, template='plotly_white', yaxis_title="Average r")
                st.plotly_chart(fig1, use_container_width=True)
                st.download_button("📥 엑셀", data=make_excel({'Average Correlation': sdf}),
                                   file_name="Output1_Avg_Correlation.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ═══ OUTPUT 2: SPREAD vs DRIVER ROLLING ═══
with tab2:
    st.subheader("Spread vs Fundamental Driver Rolling Correlation")
    st.caption("All Year + 3개월 Rolling (−1/0/+1달). 실제 일자 표시.")
    if idx_df is None or fund_df is None:
        st.warning("인덱스 + Fundamental 데이터를 모두 업로드하세요.")
    else:
        c2a,c2b = st.columns(2)
        with c2a:
            st.markdown("**Spread 정의**"); all_idx = sorted(idx_df['Index명'].unique())
            idx1 = st.selectbox("Index 1", all_idx, key='t2_i1')
            y1o,m1v = contract_month_selector('t2_cm1', label="월물 1")
            idx2 = st.selectbox("Index 2", all_idx, key='t2_i2', index=min(1,len(all_idx)-1))
            y2o,m2v = contract_month_selector('t2_cm2', label="월물 2")
        with c2b:
            st.markdown("**Fundamental Driver (복수)**")
            sd2 = driver_multiselect(fund_df, key_prefix='t2', label="Driver")
            sy2 = year_selector('t2', fund_df=fund_df, idx_df=idx_df)
        sp_lbl = f"{idx1} - {idx2} (Y+{y1o} {m1v}월/Y+{y2o} {m2v}월)"

        if st.button("🔍 분석", key='t2_run') and sy2 and sd2:
            res = []
            for yr in sy2:
                mo1 = resolve_contract_month(yr, y1o, m1v)
                mo2 = resolve_contract_month(yr, y2o, m2v)
                for fi,fd in sd2:
                    # All Year
                    ys,ye = pd.Timestamp(f'{yr}-01-01'), pd.Timestamp(f'{yr}-12-31')
                    sp_all = get_spread_curve(idx_df, idx1, mo1, idx2, mo2, ys, ye)
                    if len(sp_all)>=5:
                        sm = sp_all.set_index('기준일자').resample('ME')['Spread'].mean().dropna()
                        fm = get_fund_monthly(fund_df, fi, fd, pd.Period(f'{yr}-01',freq='M'), pd.Period(f'{yr}-12',freq='M'))
                        if len(sm)>=3 and len(fm)>=3:
                            n=min(len(sm),len(fm)); r=pearson_corr(sm.values[:n],fm['Value'].values[:n])
                            res.append({'Spread':sp_lbl,'Driver':fd,'Index (Fund)':fi,'Year':yr,
                                        'Spread Window':'All Year','Fundamental Window':'All Year',
                                        'Offset':'동일','Pearson r':round(r,4) if not np.isnan(r) else None,'N':n})
                    # Rolling 3M
                    for sm_s in range(1,11):
                        sm_e = sm_s+2
                        if sm_e>12: break
                        ws = pd.Timestamp(f'{yr}-{sm_s:02d}-01')
                        we = pd.Timestamp(f'{yr}-{sm_e:02d}-01')+pd.offsets.MonthEnd(0)
                        sp = get_spread_curve(idx_df, idx1, mo1, idx2, mo2, ws, we)
                        if len(sp)<5: continue
                        spm = sp.set_index('기준일자').resample('ME')['Spread'].mean().dropna()
                        sw_str = f"{sm_s}~{sm_e}월"
                        for off,ol in [(-1,'-1달'),(0,'동일'),(1,'+1달')]:
                            fs=sm_s+off; fe=sm_e+off; fys=yr; fye=yr
                            if fs<1: fs+=12; fys-=1
                            if fs>12: fs-=12; fys+=1
                            if fe<1: fe+=12; fye-=1
                            if fe>12: fe-=12; fye+=1
                            fsy = pd.Period(f'{fys}-{fs:02d}',freq='M')
                            fey = pd.Period(f'{fye}-{fe:02d}',freq='M')
                            fmo = get_fund_monthly(fund_df, fi, fd, fsy, fey)
                            fw_str = f"{fys}.{fs:02d}~{fye}.{fe:02d}"
                            if len(fmo)>=3 and len(spm)>=3:
                                sv = spm.values[:min(len(spm),len(fmo))]
                                fv = fmo['Value'].values[:len(sv)]
                                if len(sv)>=3:
                                    r = pearson_corr(sv,fv)
                                    res.append({'Spread':sp_lbl,'Driver':fd,'Index (Fund)':fi,'Year':yr,
                                                'Spread Window':sw_str,'Fundamental Window':fw_str,
                                                'Offset':ol,'Pearson r':round(r,4) if not np.isnan(r) else None,'N':len(sv)})
            if res:
                rdf = pd.DataFrame(res)
                st.dataframe(rdf, use_container_width=True, height=500)
                for fi,fd in sd2:
                    ds = rdf[(rdf['Driver']==fd)&(rdf['Index (Fund)']==fi)]
                    if ds.empty: continue
                    fig2 = go.Figure()
                    for ol in ['-1달','동일','+1달']:
                        sub = ds[ds['Offset']==ol]
                        fig2.add_trace(go.Bar(x=[f"{r['Year']} {r['Spread Window']}" for _,r in sub.iterrows()],
                                              y=sub['Pearson r'], name=f"Fund {ol}",
                                              marker_color={'-1달':'#3498DB','동일':'#006D36','+1달':'#E67E22'}.get(ol,'#999')))
                    fig2.update_layout(title=f"Spread vs [{fi}] {fd}", barmode='group', height=420, template='plotly_white',
                                       legend=dict(orientation='h',y=-0.15))
                    st.plotly_chart(fig2, use_container_width=True)
                st.download_button("📥 엑셀", data=make_excel({'Rolling Correlation':rdf}),
                                   file_name="Output2_Spread_vs_Driver.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else: st.warning("분석 결과가 없습니다.")

# ═══ OUTPUT 3: INDEX vs DRIVER ROLLING ═══
with tab3:
    st.subheader("Index vs Fundamental Driver Rolling Correlation")
    st.caption("All Year + 3개월 Rolling (−1/0/+1달). 실제 일자 표시.")
    if idx_df is None or fund_df is None:
        st.warning("인덱스 + Fundamental 데이터를 모두 업로드하세요.")
    else:
        c3a,c3b = st.columns(2)
        with c3a:
            st.markdown("**Index 설정**")
            si3 = st.selectbox("Index", sorted(idx_df['Index명'].unique()), key='t3_idx')
            y3o,m3v = contract_month_selector('t3_cm', label="월물")
        with c3b:
            st.markdown("**Fundamental Driver (복수)**")
            sd3 = driver_multiselect(fund_df, key_prefix='t3', label="Driver")
            sy3 = year_selector('t3', fund_df=fund_df, idx_df=idx_df)
        il3 = f"{si3} (Y+{y3o} {m3v}월)"

        if st.button("🔍 분석", key='t3_run') and sy3 and sd3:
            res3 = []
            for yr in sy3:
                sm3 = resolve_contract_month(yr, y3o, m3v)
                for fi,fd in sd3:
                    # All Year
                    ys,ye = pd.Timestamp(f'{yr}-01-01'), pd.Timestamp(f'{yr}-12-31')
                    isub = idx_df[(idx_df['Index명']==si3)&(idx_df['월물']==sm3)&(idx_df['기준일자']>=ys)&(idx_df['기준일자']<=ye)].sort_values('기준일자')
                    if len(isub)>=5:
                        im = isub.set_index('기준일자').resample('ME')['Value'].mean().dropna()
                        fm = get_fund_monthly(fund_df, fi, fd, pd.Period(f'{yr}-01',freq='M'), pd.Period(f'{yr}-12',freq='M'))
                        if len(im)>=3 and len(fm)>=3:
                            n=min(len(im),len(fm)); r=pearson_corr(im.values[:n],fm['Value'].values[:n])
                            res3.append({'인덱스':il3,'Driver':fd,'Index (Fund)':fi,'Year':yr,
                                         'Index Window':'All Year','Fundamental Window':'All Year',
                                         'Offset':'동일','Pearson r':round(r,4) if not np.isnan(r) else None,'N':n})
                    # Rolling
                    for sm_s in range(1,11):
                        sm_e = sm_s+2
                        if sm_e>12: break
                        ws = pd.Timestamp(f'{yr}-{sm_s:02d}-01')
                        we = pd.Timestamp(f'{yr}-{sm_e:02d}-01')+pd.offsets.MonthEnd(0)
                        isub2 = idx_df[(idx_df['Index명']==si3)&(idx_df['월물']==sm3)&(idx_df['기준일자']>=ws)&(idx_df['기준일자']<=we)].sort_values('기준일자')
                        if len(isub2)<5: continue
                        imm = isub2.set_index('기준일자').resample('ME')['Value'].mean().dropna()
                        iw_str = f"{sm_s}~{sm_e}월"
                        for off,ol in [(-1,'-1달'),(0,'동일'),(1,'+1달')]:
                            fs=sm_s+off; fe=sm_e+off; fys=yr; fye=yr
                            if fs<1: fs+=12; fys-=1
                            if fs>12: fs-=12; fys+=1
                            if fe<1: fe+=12; fye-=1
                            if fe>12: fe-=12; fye+=1
                            fsy = pd.Period(f'{fys}-{fs:02d}',freq='M')
                            fey = pd.Period(f'{fye}-{fe:02d}',freq='M')
                            fmo = get_fund_monthly(fund_df, fi, fd, fsy, fey)
                            fw_str = f"{fys}.{fs:02d}~{fye}.{fe:02d}"
                            if len(fmo)>=3 and len(imm)>=3:
                                n=min(len(imm),len(fmo)); r=pearson_corr(imm.values[:n],fmo['Value'].values[:n])
                                res3.append({'인덱스':il3,'Driver':fd,'Index (Fund)':fi,'Year':yr,
                                             'Index Window':iw_str,'Fundamental Window':fw_str,
                                             'Offset':ol,'Pearson r':round(r,4) if not np.isnan(r) else None,'N':n})
            if res3:
                r3df = pd.DataFrame(res3)
                st.dataframe(r3df, use_container_width=True, height=500)
                for fi,fd in sd3:
                    ds = r3df[(r3df['Driver']==fd)&(r3df['Index (Fund)']==fi)]
                    if ds.empty: continue
                    fig3 = go.Figure()
                    for ol in ['-1달','동일','+1달']:
                        sub = ds[ds['Offset']==ol]
                        fig3.add_trace(go.Bar(x=[f"{r['Year']} {r['Index Window']}" for _,r in sub.iterrows()],
                                              y=sub['Pearson r'], name=f"Fund {ol}",
                                              marker_color={'-1달':'#3498DB','동일':'#006D36','+1달':'#E67E22'}.get(ol,'#999')))
                    fig3.update_layout(title=f"{il3} vs [{fi}] {fd}", barmode='group', height=420, template='plotly_white',
                                       legend=dict(orientation='h',y=-0.15))
                    st.plotly_chart(fig3, use_container_width=True)
                st.download_button("📥 엑셀", data=make_excel({'Rolling Correlation':r3df}),
                                   file_name=f"Output3_{safe_sheet_name(si3)}_vs_drivers.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else: st.warning("분석 결과가 없습니다.")

st.markdown("---")
st.caption("Built with Streamlit | BCG × SK Gas Trading AX | v3.0")
