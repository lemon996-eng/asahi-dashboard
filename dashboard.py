import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import re
import os
import unicodedata
import json
from datetime import datetime

st.set_page_config(page_title="아사히 마시나리 대시보드", layout="wide", initial_sidebar_state="expanded")

# 🎨 [안전한 모던 클린 CSS] 스트림릿 구조를 파괴하지 않는 타겟팅
st.markdown("""
<style>
/* 📱 모바일 화면 전체 좌우 밀림 완벽 방지 */
html, body, [data-testid="stAppViewContainer"], .main {
    max-width: 100vw !important;
    overflow-x: hidden !important;
}
* {
    box-sizing: border-box !important;
}
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');

/* 1. 전체 배경: 가장 안전한 뷰 컨테이너만 타겟팅 */
[data-testid="stAppViewContainer"] { 
    background-color: #F3F4F8 !important; 
    font-family: 'Pretendard', sans-serif !important;
}

h1, h2, h3 { 
    color: #1A1A2E !important; 
    font-weight: 800 !important; 
    letter-spacing: -0.5px !important;
}

/* 2. 핵심 지표(KPI) 카드: 안전한 둥근 테두리 및 그림자 */
[data-testid="stMetric"] { 
    background-color: #FFFFFF !important; 
    border-radius: 16px !important; 
    padding: 16px 24px !important; 
    box-shadow: 0px 4px 12px rgba(30, 41, 59, 0.05) !important;
    border: 1px solid #E2E8F0 !important;
}

/* 3. 버튼 및 엑셀 다운로드 (글자 뭉개짐 해결) */
div.stButton > button, div.stDownloadButton > button { 
    background-color: #1E2772 !important; 
    border: none !important; 
    border-radius: 8px !important; 
    padding: 8px 24px !important;
    box-shadow: 0px 4px 10px rgba(30, 39, 114, 0.2) !important;
}
/* 버튼 내부의 모든 글자를 강제로 흰색 처리 */
div.stButton > button *, div.stDownloadButton > button * { 
    color: #FFFFFF !important; 
    font-weight: 700 !important; 
}
div.stButton > button:hover, div.stDownloadButton > button:hover { 
    background-color: #141A4C !important; 
}

/* 4. 입력 필드 */
div[data-testid="stForm"], div.stDateInput > div > div > input, div[data-baseweb="select"] > div { 
    border-radius: 8px !important; 
    border: 1px solid #E2E8F0 !important; 
    background-color: #FFFFFF !important;
}

/* 5. 데이터 프레임(표) 네이티브 디자인 */
[data-testid="stDataFrame"] {
    background-color: #FFFFFF !important;
    border-radius: 16px !important;
    padding: 16px !important;
    box-shadow: 0px 4px 12px rgba(30, 41, 59, 0.05) !important;
    border: 1px solid #E2E8F0 !important;
    margin-bottom: 20px !important;
}
hr { border-color: #E2E8F0 !important; margin: 2em 0 !important; }
</style>
""", unsafe_allow_html=True)

st.title("🏭 한국 아사히 마시나리 - 생산 대시보드")
st.markdown("---")

DB_FILE_PATH = "아사히_마스터_DB.csv"
SETTINGS_FILE_PATH = "대시보드_검색기록.json"
ACCOUNT_REGEX = re.compile(r'^(\d{4})[-_]*([A-Z]*)[-_]*(.*)$')
SPLIT_REGEX = re.compile(r'[~_-]+')
GEUNTAE_PATTERN = re.compile('휴가|조퇴|외출|지각|休|早退|外出|遲刻', flags=re.IGNORECASE)

# 특이 근태 판별용 개별 패턴 (compile 1회로 재사용 -> tab3 반복 계산 최적화)
PATTERN_VACATION = re.compile('휴가|休')
PATTERN_OUTING = re.compile('외출|外出')
PATTERN_LATE = re.compile('지각|遲刻')
PATTERN_EARLY = re.compile('조퇴|早退')

BRAND_COLORS = ['#1E2772', '#A3E635', '#7C3AED', '#F59E0B', '#3B82F6', '#10B981']

if "processed_file_names" not in st.session_state:
    st.session_state.processed_file_names = set()

def load_filter_settings():
    if os.path.exists(SETTINGS_FILE_PATH):
        try:
            with open(SETTINGS_FILE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except: return {}
    return {}

def save_filter_settings(settings):
    with open(SETTINGS_FILE_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)

@st.cache_data(show_spinner=False)
def load_master_db():
    if os.path.exists(DB_FILE_PATH):
        df = pd.read_csv(DB_FILE_PATH, dtype={'작업일자': str, '구좌명': str, '부서': str, '작업자': str, '근태': str})
        df['근태'] = df['근태'].fillna('')
        return df
    else: return pd.DataFrame()

def save_master_db(df):
    df.to_csv(DB_FILE_PATH, index=False, encoding='utf-8-sig')

master_db = load_master_db()

st.sidebar.header("📁 데이터 업로드")
uploaded_files = st.sidebar.file_uploader("파일 드래그 앤 드롭", type=['xlsx', 'xls'], accept_multiple_files=True)

st.sidebar.markdown("---")
st.sidebar.subheader("🗑️ DB 관리")
if st.sidebar.button("🚨 마스터 DB 초기화"):
    if os.path.exists(DB_FILE_PATH): os.remove(DB_FILE_PATH)
    if os.path.exists(SETTINGS_FILE_PATH): os.remove(SETTINGS_FILE_PATH)
    st.cache_data.clear()
    st.session_state.processed_file_names = set()
    try: st.rerun()
    except AttributeError: st.experimental_rerun()

dept_mapping = {'機械': '기계', '기계부': '기계', '電機': '전기', '전기부': '전기', '電裝': '전장', '전장부': '전장', '組立': '조립', '조립부': '조립', '設計': '설계', '설계부': '설계', '檢査': '검사', '검사부': '검사', '加工': '가공', '가공부': '가공', '制御': '제어', '제어부': '제어', '品質': '품질', '품질부': '품질'}
worker_mapping = {'金雲石': '김운석', '朴振求': '박진구', '黃斗煥': '황두환', '李東在': '이동재', '李东在': '이동재', '李東宰': '이동재', '金映德': '김영덕', '金正吉': '김정길', '李哲珉': '이철민', '李哲民': '이철민', '李喆珉': '이철민', '金泰旻': '김태민', '丁海成': '정해성', '金建佑': '김건우', '金榮勳': '김영훈', '崔仁河': '최인하', '咸同圭': '함동규', '朴두리': '박두리', '黃纘赫': '황찬혁', '韓載壽': '한재수', '金泰賢': '김태현', '安成任': '안성임'}
# 매핑 키를 미리 1회 정규화해 두어, translate_name 호출마다 반복되던 normalize 연산을 제거
NORMALIZED_WORKER_MAPPING = [(unicodedata.normalize('NFKC', hanja), korean) for hanja, korean in worker_mapping.items()]

KOR_HOLIDAYS = ['2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18', '2026-03-01', '2026-05-05', '2026-05-24', '2026-06-06', '2026-08-15', '2026-09-24', '2026-09-25', '2026-09-26', '2026-10-03', '2026-10-09', '2026-12-25', '2025-01-01', '2025-01-28', '2025-01-29', '2025-01-30', '2025-03-01', '2025-05-05', '2025-06-06', '2025-08-15', '2025-10-03', '2025-10-05', '2025-10-06', '2025-10-07', '2025-10-09', '2025-12-25']
# 반복 조회(in 연산)를 위해 set으로도 보관 (리스트 순회보다 조회 속도 개선)
KOR_HOLIDAYS_SET = frozenset(KOR_HOLIDAYS)

def format_unmanned(val):
    try:
        v = float(val)
        if v == int(v): return f"{int(v):02d}시간"
        else: return f"{v:05.2f}시간"
    except: return "00시간"

def translate_name(name_str):
    if pd.isna(name_str): return ""
    name_str = re.sub(r'\s+', '', str(name_str))
    name_str = unicodedata.normalize('NFKC', name_str)
    for norm_hanja, korean in NORMALIZED_WORKER_MAPPING:
        if norm_hanja in name_str: name_str = name_str.replace(norm_hanja, korean)
    return name_str

def parse_and_split_accounts(account_str, reg_time, over_time, unmanned_time=0):
    account_str = str(account_str).strip().upper().replace('-구좌', '').replace('(TSG)', '')
    account_str = re.sub(r'\s+', '', account_str)
    match = ACCOUNT_REGEX.match(account_str)
    if not match: return [(account_str, reg_time, over_time, unmanned_time)]
    prefix_num, prefix_alpha, rest = match.groups()
    prefix = f"{prefix_num}-{prefix_alpha}" if prefix_alpha else f"{prefix_num}-"
    parts = [p for p in SPLIT_REGEX.split(rest) if p.strip()]
    if not parts: return [(f"{prefix_num}-{prefix_alpha}" if prefix_alpha else prefix_num, reg_time, over_time, unmanned_time)]
    accounts = [f"{prefix}{p}" for p in parts]
    n = len(accounts)
    if n == 1: return [(accounts[0], reg_time, over_time, unmanned_time)]
    def distribute_time(val):
        cents = int(round(val * 100))
        dist = [cents // n] * n
        for i in range(cents % n): dist[n - 1 - i] += 1
        return [round(c / 100, 2) for c in dist]
    return list(zip(accounts, distribute_time(reg_time), distribute_time(over_time), distribute_time(unmanned_time)))

def generate_attendance_html(df):
    if df.empty: return ""
    agg = df.groupby(['작업자', '작업일자']).agg({'정규시간': 'sum', '잔업시간': 'sum'}).reset_index()
    dates = sorted(agg['작업일자'].unique())
    date_cols = []
    is_weekend = {}
    for d in dates:
        dt = pd.to_datetime(d)
        weekday_kr = ['월', '화', '수', '목', '금', '토', '일'][dt.weekday()]
        col_name = f"{dt.month}/{dt.day}{weekday_kr}"
        date_cols.append((d, col_name))
        is_weekend[d] = dt.weekday() >= 5
    workers = sorted(agg['작업자'].unique())

    # 문자열을 '+='로 반복 이어붙이면 데이터가 커질수록 느려지므로(O(n^2)),
    # 리스트에 조각을 모았다가 마지막에 한 번만 join (결과는 완전히 동일, 성능만 개선)
    html_parts = []
    html_parts.append("""
    <div style='overflow-x: auto; background-color: #FFFFFF; border-radius: 16px; padding: 24px; border: 1px solid #E2E8F0; box-shadow: 0px 4px 12px rgba(30, 41, 59, 0.05); margin-bottom: 20px;'>
    <table style='border-collapse: separate; border-spacing: 0; width: 100%; text-align: center; font-size: 14px; font-family: "Pretendard", sans-serif; white-space: nowrap;'>
        <tr>
            <th style='border-bottom: 2px solid #F1F5F9; padding: 15px 10px; color: #64748B; font-weight: 600;'>이름</th>
    """)
    for d, c in date_cols:
        color = "#EF4444" if is_weekend[d] else "#64748B"
        html_parts.append(f"<th style='border-bottom: 2px solid #F1F5F9; padding: 15px 10px; color: {color}; font-weight: 600;'>{c}</th>")

    html_parts.append("<th style='border-bottom: 2px solid #F1F5F9; padding: 15px 10px; color: #1E2772; font-weight: 700;'>합계(h)</th></tr>")

    # 작업자별 일자 데이터를 groupby로 미리 나눠두어, 반복 필터링(day_data = w_data[...]) 비용을 절감
    grouped_by_worker = {w: wdf.set_index('작업일자') for w, wdf in agg.groupby('작업자')}

    for w in workers:
        html_parts.append(f"<tr><td style='border-bottom: 1px solid #F8FAFC; padding: 12px 10px; font-weight: 700; color: #1A1A2E;'>{w}</td>")
        w_data = grouped_by_worker.get(w)
        total_sum = 0
        for d, c in date_cols:
            bg = "#FEF2F2" if is_weekend[d] else "#FFFFFF"
            if w_data is None or d not in w_data.index:
                html_parts.append(f"<td style='border-bottom: 1px solid #F8FAFC; background-color: {bg};'></td>")
            else:
                day_rows = w_data.loc[[d]]
                reg = day_rows['정규시간'].sum()
                over = day_rows['잔업시간'].sum()
                total = reg + over
                total_sum += total
                if total == 0:
                    html_parts.append(f"<td style='border-bottom: 1px solid #F8FAFC; background-color: {bg};'></td>")
                elif is_weekend[d]:
                    html_parts.append(f"<td style='border-bottom: 1px solid #F8FAFC; background-color: {bg}; font-weight: 700; color: #EF4444;'>{total:.2f}</td>")
                else:
                    reg_str = f"{reg:.2f}" if reg > 0 else "-"
                    over_str = f"{over:.2f}" if over > 0 else "-"
                    cell = f"<td style='border-bottom: 1px solid #F8FAFC; background-color: {bg};'>"
                    cell += f"<div style='font-weight: 500; color: #94A3B8; margin-bottom: 4px;'>{reg_str}</div>"
                    if over > 0 or reg > 0: cell += f"<div style='font-weight: 700; color: #1E2772;'>{over_str}</div>"
                    cell += "</td>"
                    html_parts.append(cell)
        html_parts.append(f"<td style='border-bottom: 1px solid #F8FAFC; font-weight: 800; color: #1E2772;'>{total_sum:.2f}</td></tr>")
    html_parts.append("</table></div>")
    return "".join(html_parts)

def apply_modern_chart_layout(fig, x_title, y_title, legend_title=""):
    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Pretendard, sans-serif", color="#64748B", size=13),
        margin=dict(t=30, b=40, l=40, r=20),
        xaxis=dict(title=x_title, showgrid=False, linecolor="#E2E8F0", linewidth=1, tickfont=dict(color="#94A3B8")),
        yaxis=dict(title=y_title, showgrid=True, gridcolor="#F1F5F9", gridwidth=1, griddash="solid", tickfont=dict(color="#94A3B8")),
        legend=dict(title=legend_title, orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hoverlabel=dict(bgcolor="#FFFFFF", font_size=14, font_family="Pretendard", bordercolor="#E2E8F0")
    )
    fig.update_traces(marker_line_width=0, opacity=0.9)
    return fig

def render_tab_filters(tab_id, df):
    global_settings = load_filter_settings()
    tab_settings = global_settings.get(tab_id, {})

    # 작업일자를 1회만 datetime으로 변환해 min/max 계산과 필터링에 재사용
    # (기존 코드는 동일한 변환을 두 번 수행했음 -> 결과는 동일, 계산 비용만 절감)
    dates_dt = pd.to_datetime(df['작업일자'])
    min_date = dates_dt.min().date()
    max_date = dates_dt.max().date()

    def_start_str = tab_settings.get("start_date", str(min_date))
    def_end_str = tab_settings.get("end_date", str(max_date))
    def_start = max(min_date, min(pd.to_datetime(def_start_str).date(), max_date))
    def_end = max(min_date, min(pd.to_datetime(def_end_str).date(), max_date))

    dept_options = sorted(df['부서'].dropna().unique().tolist())
    worker_options = sorted(df['작업자'].dropna().unique().tolist())
    acc_options = sorted(df['구좌명'].dropna().unique().tolist())

    def_dept = [x for x in tab_settings.get("dept", []) if x in dept_options]
    def_worker = [x for x in tab_settings.get("worker", []) if x in worker_options]
    def_acc = [x for x in tab_settings.get("acc", []) if x in acc_options]

    st.markdown(f"<div style='font-weight:700; color:#1A1A2E; margin-bottom:10px;'>🔍 {tab_id} 상세 필터</div>", unsafe_allow_html=True)
    col_d1, col_d2, col_dept, col_worker, col_acc, col_dl = st.columns([1, 1, 1.2, 1.2, 1.2, 1])

    with col_d1: start_d = st.date_input("📅 시작일", min_value=min_date, max_value=max_date, value=def_start, key=f"start_{tab_id}")
    with col_d2: end_d = st.date_input("📅 종료일", min_value=min_date, max_value=max_date, value=def_end, key=f"end_{tab_id}")
    with col_dept: dept_search = st.multiselect("🏢 부서", options=dept_options, default=def_dept, placeholder="전체", key=f"dept_{tab_id}")
    with col_worker: worker_search = st.multiselect("👷 작업자", options=worker_options, default=def_worker, placeholder="전체", key=f"worker_{tab_id}")
    with col_acc: account_search = st.multiselect("⚙️ 구좌명", options=acc_options, default=def_acc, placeholder="전체", key=f"acc_{tab_id}")

    current_state = {
        "start_date": str(start_d),
        "end_date": str(end_d),
        "dept": dept_search,
        "worker": worker_search,
        "acc": account_search
    }
    if current_state != tab_settings:
        global_settings[tab_id] = current_state
        save_filter_settings(global_settings)

    filtered_df = df.copy()
    filtered_df = filtered_df[(dates_dt.dt.date >= start_d) & (dates_dt.dt.date <= end_d)]

    if dept_search: filtered_df = filtered_df[filtered_df['부서'].isin(dept_search)]
    if worker_search: filtered_df = filtered_df[filtered_df['작업자'].isin(worker_search)]
    if account_search: filtered_df = filtered_df[filtered_df['구좌명'].isin(account_search)]

    with col_dl:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        csv_data = filtered_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(label="📥 엑셀 저장", data=csv_data, file_name=f"아사히_마시나리_{tab_id}.csv", mime="text/csv", key=f"dl_{tab_id}", use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)
    return filtered_df

if uploaded_files:
    new_files = [f for f in uploaded_files if f.name not in st.session_state.processed_file_names]

    if new_files:
        all_cleaned_data = []
        with st.spinner('🚀 데이터를 스캔 중입니다...'):
            for file in new_files:
                df_raw = pd.read_excel(file)
                df_clean = pd.DataFrame()

                if '作業日' in df_raw.columns: df_clean['작업일자'] = df_raw['作業日']
                elif '작업일자' in df_raw.columns: df_clean['작업일자'] = df_raw['작업일자']
                if '作業者' in df_raw.columns: df_clean['작업자'] = df_raw['作業者']
                elif '작업자' in df_raw.columns: df_clean['작업자'] = df_raw['작업자']
                if '口座番號' in df_raw.columns: df_clean['구좌명'] = df_raw['口座番號']
                elif '구좌번호' in df_raw.columns: df_clean['구좌명'] = df_raw['구좌번호']
                if '部署' in df_raw.columns: df_clean['부서'] = df_raw['部署'].fillna('미지정')
                elif '부서' in df_raw.columns: df_clean['부서'] = df_raw['부서'].fillna('미지정')

                if '作業' in df_raw.columns and '殘業' in df_raw.columns:
                    df_clean['정규시간'] = pd.to_numeric(df_raw['作業'], errors='coerce').fillna(0).astype(float)
                    df_clean['잔업시간'] = pd.to_numeric(df_raw['殘業'], errors='coerce').fillna(0).astype(float)
                elif '산출시간(h)' in df_raw.columns:
                    temp_total = pd.to_numeric(df_raw['산출시간(h)'], errors='coerce').fillna(0).astype(float)
                    if '잔업여부' in df_raw.columns:
                        is_over = df_raw['잔업여부'] == 'O'
                        df_clean['잔업시간'] = temp_total * is_over
                        df_clean['정규시간'] = temp_total - df_clean['잔업시간']
                    else:
                        df_clean['정규시간'] = temp_total
                        df_clean['잔업시간'] = 0
                else:
                    df_clean['정규시간'] = 0.0
                    df_clean['잔업시간'] = 0.0

                temp_dates = pd.to_datetime(df_clean['작업일자'], errors='coerce')
                mask_c = temp_dates <= pd.to_datetime('2026-05-31')
                mask_c = mask_c.fillna(False)
                df_clean.loc[mask_c, '잔업시간'] = df_clean.loc[mask_c, '잔업시간'] * 1.5
                df_clean['총시간'] = df_clean['정규시간'] + df_clean['잔업시간']

                if '무인가공(h)' in df_raw.columns:
                    df_clean['무인가공'] = pd.to_numeric(df_raw['무인가공(h)'], errors='coerce').fillna(0).astype(float)
                else:
                    df_clean['무인가공'] = 0.0

                geuntae_series = pd.Series('', index=df_raw.index)
                cols_to_check = [col for col in df_raw.columns if col not in ['作業者', '작업자', '作業日', '작업일자']]
                for col in cols_to_check:
                    col_data = df_raw[col]
                    if col_data.dtype == object or pd.api.types.is_string_dtype(col_data):
                        col_str = col_data.astype(str)
                        mask = col_str.str.contains(GEUNTAE_PATTERN, na=False)
                        if mask.any(): geuntae_series.loc[mask] = (geuntae_series.loc[mask] + " " + col_str.loc[mask]).str.strip()
                df_clean['근태'] = geuntae_series

                if not df_clean.empty and '총시간' in df_clean.columns:
                    df_clean['작업일자'] = pd.to_datetime(df_clean['작업일자'], errors='coerce').dt.strftime('%Y-%m-%d')
                    df_clean['작업자'] = df_clean['작업자'].apply(translate_name)
                    df_clean['부서'] = df_clean['부서'].astype(str).str.strip().replace(dept_mapping)

                    mask_valid = (df_clean['총시간'] > 0) | (df_clean['무인가공'] > 0) | (df_clean['근태'] != '')
                    df_valid = df_clean[mask_valid].copy()

                    if not df_valid.empty:
                        def process_row(row):
                            splits = parse_and_split_accounts(row['구좌명'], row['정규시간'], row['잔업시간'], row['무인가공'])
                            res = []
                            for acc, r, o, u in splits: res.append({'구좌명': acc, '정규시간': r, '잔업시간': o, '총시간': r + o, '무인가공': u})
                            return res

                        df_valid['splits'] = df_valid.apply(process_row, axis=1)
                        df_expanded = df_valid.explode('splits').reset_index(drop=True)
                        splits_df = pd.DataFrame(df_expanded['splits'].tolist(), index=df_expanded.index)
                        df_expanded = df_expanded.drop(columns=['구좌명', '정규시간', '잔업시간', '총시간', '무인가공', 'splits'])
                        df_expanded = pd.concat([df_expanded, splits_df], axis=1)

                        daily_agg = df_expanded.groupby(['작업일자', '작업자', '구좌명', '부서']).agg({
                            '정규시간': 'sum', '잔업시간': 'sum', '총시간': 'sum', '무인가공': 'max', '근태': 'max'
                        }).reset_index()
                        all_cleaned_data.append(daily_agg)

        if all_cleaned_data:
            new_df = pd.concat(all_cleaned_data, ignore_index=True)
            master_db = pd.concat([master_db, new_df], ignore_index=True)
            master_db['근태_있음'] = master_db['근태'] != ''
            master_db = master_db.sort_values(by=['무인가공', '근태_있음'], ascending=[False, False])
            master_db = master_db.drop_duplicates(subset=['작업일자', '작업자', '구좌명', '부서'], keep='first')
            master_db = master_db.drop(columns=['근태_있음'])
            save_master_db(master_db)
            for f in new_files: st.session_state.processed_file_names.add(f.name)
            st.success("✅ 업데이트 완료!")
            st.cache_data.clear()
            master_db = load_master_db()

if not master_db.empty:
    tab1, tab2, tab3 = st.tabs(["📊 요약 대시보드", "📅 출·퇴근 현황", "🏖️ 특이 근태"])

    with tab1:
        filtered_df_tab1 = render_tab_filters("요약대시보드", master_db)

        if not filtered_df_tab1.empty:
            st.markdown("<h3 style='margin: 15px 0;'>핵심 지표</h3>", unsafe_allow_html=True)
            unmanned_total = filtered_df_tab1.groupby('작업일자')['무인가공'].max().sum()

            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            kpi1.metric(label="총 투입 시간", value=f"{filtered_df_tab1['총시간'].sum():,.1f}h")
            kpi2.metric(label="무인가공 합계", value=format_unmanned(unmanned_total))
            kpi3.metric(label="투입 인원", value=f"{filtered_df_tab1['작업자'].nunique():,}명")
            kpi4.metric(label="진행 구좌", value=f"{filtered_df_tab1['구좌명'].nunique():,}개")
            kpi5.metric(label="작업 일수", value=f"{filtered_df_tab1['작업일자'].nunique():,}일")
            st.markdown("<br>", unsafe_allow_html=True)

            st.markdown("### 1. 구좌별 투입 시간", unsafe_allow_html=True)
            valid_accounts = filtered_df_tab1['구좌명'].unique()
            df_project_full = master_db[master_db['구좌명'].isin(valid_accounts)]
            df_project_full = df_project_full.groupby('구좌명')['총시간'].sum().reset_index().sort_values(by='구좌명', ascending=True)

            fig1 = px.bar(df_project_full, x='구좌명', y='총시간', text_auto='.1f', color_discrete_sequence=['#1E2772'])
            fig1.update_traces(textposition="outside", textfont=dict(size=12, color="#64748B"), marker_cornerradius=4)
            fig1 = apply_modern_chart_layout(fig1, "구좌명", "시간 (h)")
            st.plotly_chart(fig1, use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)

            st.markdown("### 2. 구좌별/부서별 상세 분석", unsafe_allow_html=True)
            col2_chart, col2_table = st.columns([6, 4])
            df_dept = filtered_df_tab1.groupby(['구좌명', '부서'])['총시간'].sum().reset_index().sort_values(by='구좌명', ascending=True)
            with col2_chart:
                fig2 = px.bar(df_dept, x='구좌명', y='총시간', color='부서', text_auto='.1f', color_discrete_sequence=BRAND_COLORS)
                fig2.update_traces(textposition="outside", textfont=dict(size=11, color="#64748B"))
                fig2 = apply_modern_chart_layout(fig2, "구좌명", "시간 (h)", "부서")
                st.plotly_chart(fig2, use_container_width=True)

            with col2_table:
                df_dept_pivot = filtered_df_tab1.pivot_table(index='구좌명', columns='부서', values='총시간', aggfunc='sum', fill_value=0)
                df_dept_pivot['총합계'] = df_dept_pivot.sum(axis=1)
                df_dept_pivot.loc['총계'] = df_dept_pivot.sum(axis=0)
                df_dept_pivot = df_dept_pivot.reset_index()
                st.dataframe(df_dept_pivot, use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)

            st.markdown("### 3. 작업자별 개인 생산 시간", unsafe_allow_html=True)
            filtered_df_tab1['작업구분'] = filtered_df_tab1['구좌명'].apply(lambda x: 'Z구좌 (휴식/조례)' if x == 'Z' else '일반 작업')
            df_worker_split = filtered_df_tab1.groupby(['작업자', '작업구분'])['총시간'].sum().reset_index()
            worker_total = filtered_df_tab1.groupby('작업자')['총시간'].sum().reset_index().sort_values(by='총시간', ascending=True)

            fig3 = px.bar(df_worker_split, x='총시간', y='작업자', color='작업구분', orientation='h', text_auto='.1f',
                          category_orders={"작업자": worker_total['작업자'].tolist()},
                          color_discrete_map={'일반 작업': '#1E2772', 'Z구좌 (휴식/조례)': '#A3E635'})
            fig3.update_traces(textposition="outside", textfont=dict(size=12, color="#64748B"))
            fig3 = apply_modern_chart_layout(fig3, "시간 (h)", "작업자", "구분")

            chart_col, table_col = st.columns([6, 4])
            with chart_col: st.plotly_chart(fig3, use_container_width=True)
            with table_col:
                df_pivot = filtered_df_tab1.pivot_table(index='작업자', columns='작업구분', values='총시간', aggfunc='sum', fill_value=0)
                if '일반 작업' not in df_pivot.columns: df_pivot['일반 작업'] = 0
                if 'Z구좌 (휴식/조례)' not in df_pivot.columns: df_pivot['Z구좌 (휴식/조례)'] = 0
                df_pivot['총합계'] = df_pivot['일반 작업'] + df_pivot['Z구좌 (휴식/조례)']
                df_pivot = df_pivot.sort_values(by='총합계', ascending=False).reset_index()
                st.dataframe(df_pivot, use_container_width=True)

        else:
            st.warning("데이터가 없습니다.")

    with tab2:
        filtered_df_tab2 = render_tab_filters("출퇴근및잔업", master_db)
        if not filtered_df_tab2.empty:
            html_table = generate_attendance_html(filtered_df_tab2)
            st.markdown(html_table, unsafe_allow_html=True)
        else:
            st.info("데이터가 없습니다.")

    with tab3:
        filtered_df_tab3 = render_tab_filters("특이근태현황", master_db)
        st.info("💡 **알림:** 주말(토/일) 및 법정 공휴일에 발생한 '조퇴', '지각', '외출'은 자동 제외됩니다.")

        df_geuntae = filtered_df_tab3[filtered_df_tab3['근태'].str.contains(GEUNTAE_PATTERN, na=False)][['작업일자', '부서', '작업자', '근태']]
        df_geuntae = df_geuntae.drop_duplicates(subset=['작업일자', '작업자'], keep='first')

        if not df_geuntae.empty:
            # 기존에는 행 단위 apply에서 매번 pd.to_datetime을 호출했으나,
            # 날짜 변환과 주말/공휴일 판정, 유형 분류를 모두 벡터화(np.select)하여
            # 동일한 결과를 더 적은 연산으로 산출 (elif 우선순위: 휴가 > 외출 > 지각 > 조퇴 그대로 유지)
            x_str = df_geuntae['근태'].astype(str)
            dt_series = pd.to_datetime(df_geuntae['작업일자'])
            is_weekend_holiday = (dt_series.dt.weekday >= 5) | df_geuntae['작업일자'].isin(KOR_HOLIDAYS_SET)

            has_vacation = x_str.str.contains(PATTERN_VACATION, na=False)
            has_outing = x_str.str.contains(PATTERN_OUTING, na=False)
            has_late = x_str.str.contains(PATTERN_LATE, na=False)
            has_early = x_str.str.contains(PATTERN_EARLY, na=False)

            conditions = [
                has_vacation,
                has_outing & is_weekend_holiday,
                has_outing & ~is_weekend_holiday,
                has_late & is_weekend_holiday,
                has_late & ~is_weekend_holiday,
                has_early & is_weekend_holiday,
                has_early & ~is_weekend_holiday,
            ]
            choices = ['휴가', '제외', '외출', '제외', '지각', '제외', '조퇴']
            df_geuntae = df_geuntae.copy()
            df_geuntae['구분'] = np.select(conditions, choices, default='기타')

            df_geuntae = df_geuntae[~df_geuntae['구분'].isin(['기타', '제외'])]

            if not df_geuntae.empty:
                counts = df_geuntae['구분'].value_counts()
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("🌴 휴가", f"{counts.get('휴가', 0)} 건")
                c2.metric("🏃 평일 조퇴", f"{counts.get('조퇴', 0)} 건")
                c3.metric("🚶 평일 외출", f"{counts.get('외출', 0)} 건")
                c4.metric("⏰ 평일 지각", f"{counts.get('지각', 0)} 건")

                chart_col, table_col = st.columns([4, 6])
                with chart_col:
                    fig_g = px.histogram(df_geuntae, x='부서', color='구분', text_auto=True, barmode='group', color_discrete_sequence=BRAND_COLORS)
                    fig_g = apply_modern_chart_layout(fig_g, "부서", "건수", "구분")
                    st.plotly_chart(fig_g, use_container_width=True)

                with table_col:
                    df_display = df_geuntae.drop(columns=['구분']).sort_values(by=['작업일자', '부서', '작업자']).reset_index(drop=True)
                    st.dataframe(df_display, use_container_width=True)
            else:
                st.info("특이 근태가 없습니다.")
        else:
            st.info("특이 근태 기록이 없습니다.")
else:
    st.info("👈 사이드바에서 엑셀 파일을 업로드해 주세요!")