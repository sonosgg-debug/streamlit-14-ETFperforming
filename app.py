"""
app.py
한국(K Market) 및 미국(US Market) 증시 ETF 수익률 비교 대시보드
- 시장 선택: K Market, US Market
- 배율 선택: 3X, 2X, 1X(디폴트), -1X, -2X, -3X, 전체
- 15개 필수 컬럼 및 거래대금 기준 상위 100개 순위
- 7대 기간(1W, 2W, 1M, 3M, 6M, 1Y, 3Y) 수익률 스마트 정렬 필터
- 서식 적용 엑셀(.xlsx) 다운로드
- 종목별 상세 정보 및 4대 시각화 인터랙티브 차트 (이동평균, 구간수익률, 벤치마크 비교, MDD 리스크)
"""

import sys
import io
import os
import datetime
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots

import data_loader

# ==========================================
# 1. 페이지 기본 설정
# ==========================================
st.set_page_config(
    page_title="한국 및 미국 증시 ETF 수익률 비교",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. 프리미엄 다크 테마 커스텀 CSS (39 DividendStock 일관 계승)
# ==========================================
st.markdown("""
<style>
    /* Main Background */
    .stApp {
        background-color: #0f172a;
        color: #f8fafc;
        font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Malgun Gothic", "맑은 고딕", sans-serif;
    }
    
    /* Main Content Area */
    .main .block-container,
    [data-testid="stMainBlockContainer"] {
        padding-top: 2.0rem !important;
        padding-bottom: 3.5rem !important;
        max-width: 98% !important;
    }
    
    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #1e293b !important;
        border-right: 1px solid #334155;
    }
    
    /* Main Title & Headers */
    h1, .main h1, [data-testid="stHeadingWithActionElements"] h1 {
        color: #8AB4F8 !important;
        font-weight: 800 !important;
        font-size: 2.0rem !important;
    }

    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3 {
        color: #f8fafc !important;
    }

    /* Input & Select Box styling */
    .stTextInput input, .stSelectbox select, .stMultiSelect {
        background-color: #334155 !important;
        color: #f8fafc !important;
        border: 1px solid #475569 !important;
        border-radius: 6px !important;
    }

    /* Metric Cards */
    [data-testid="stMetricValue"] {
        font-size: 1.35rem !important;
        font-weight: 700 !important;
        color: #f8fafc !important;
    }
    [data-testid="stMetricValue"] > div {
        font-size: 1.35rem !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.82rem !important;
        color: #94a3b8 !important;
        font-weight: 500 !important;
    }

    /* 6-column detail metrics */
    div[data-testid="stHorizontalBlock"]:has(div[data-testid="column"]:nth-child(6)) [data-testid="stMetricValue"],
    div[data-testid="stHorizontalBlock"]:has(div[data-testid="column"]:nth-child(6)) [data-testid="stMetricValue"] > div {
        font-size: 1.15rem !important;
        font-weight: 700 !important;
    }
    div[data-testid="stHorizontalBlock"]:has(div[data-testid="column"]:nth-child(6)) [data-testid="stMetricLabel"] {
        font-size: 0.78rem !important;
    }

    /* Section Subheaders */
    .section-header {
        font-size: 1.25rem;
        font-weight: 700;
        color: #e2e8f0;
        margin-top: 10px;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* Primary & Download Buttons */
    .stButton button[kind="primary"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: none !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        transition: all 0.2s ease !important;
    }
    .stButton button[kind="primary"]:hover {
        background-color: #1d4ed8 !important;
        box-shadow: 0 0 10px rgba(37, 99, 235, 0.4) !important;
    }

    /* 다운로드 버튼 공통 통일 스타일 */
    div[data-testid="stDownloadButton"] > button,
    .stDownloadButton > button {
        background-color: #334155 !important;
        color: #f8fafc !important;
        border: 1px solid #475569 !important;
        border-radius: 6px !important;
        font-size: 0.875rem !important;
        font-weight: 500 !important;
        height: 38px !important;
        min-height: 38px !important;
        max-height: 38px !important;
        line-height: 36px !important;
        padding: 0 16px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        transition: all 0.2s ease-in-out !important;
        box-sizing: border-box !important;
    }
    div[data-testid="stDownloadButton"] > button:hover,
    .stDownloadButton > button:hover {
        background-color: #475569 !important;
        border-color: #38bdf8 !important;
        color: #ffffff !important;
        box-shadow: 0 0 10px rgba(56, 189, 248, 0.25) !important;
    }
    div[data-testid="stDownloadButton"] > button:active,
    .stDownloadButton > button:active {
        background-color: #1e293b !important;
        border-color: #0284c7 !important;
    }
    div[data-testid="stDownloadButton"] > button p,
    div[data-testid="stDownloadButton"] > button span,
    .stDownloadButton > button p,
    .stDownloadButton > button span {
        font-size: 0.875rem !important;
        font-weight: 500 !important;
        color: inherit !important;
        line-height: inherit !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    /* Badges */
    .badge-bull {
        background-color: rgba(239, 68, 68, 0.18);
        color: #f87171;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.8rem;
        font-weight: 600;
        border: 1px solid rgba(239, 68, 68, 0.35);
    }
    .badge-bear {
        background-color: rgba(59, 130, 246, 0.18);
        color: #60a5fa;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.8rem;
        font-weight: 600;
        border: 1px solid rgba(59, 130, 246, 0.35);
    }
    .badge-neutral {
        background-color: rgba(16, 185, 129, 0.18);
        color: #34d399;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.8rem;
        font-weight: 600;
        border: 1px solid rgba(16, 185, 129, 0.35);
    }

    /* =========================================================
       사이드바 접기(<<) 및 펼치기(>>) 버튼 항상 표시 및 시인성/대비 강화
       ========================================================= */
    /* 1. 사이드바가 열려 있을 때 접기 버튼 (<<) 상시 표시 */
    [data-testid="stSidebarCollapseButton"] {
        visibility: visible !important;
        opacity: 1 !important;
        display: inline-flex !important;
    }
    
    [data-testid="stSidebarCollapseButton"] button {
        visibility: visible !important;
        opacity: 1 !important;
        background-color: #1e293b !important;       /* 진한 네이비 배경 */
        border: 1.5px solid #38bdf8 !important;     /* 선명한 스카이블루 테두리로 상자 명확화 */
        border-radius: 8px !important;
        width: 38px !important;
        height: 38px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.4), 0 0 6px rgba(56, 189, 248, 0.2) !important;
        transition: all 0.2s ease !important;
    }
    
    /* 상자 내부의 << 아이콘(Material Icon span/svg/문자)을 순백색으로 강제하여 상자와 극명한 대비 구현 */
    [data-testid="stSidebarCollapseButton"] button *,
    [data-testid="stSidebarCollapseButton"] span,
    [data-testid="stSidebarCollapseButton"] [data-testid="stIconMaterial"],
    [data-testid="stSidebarCollapseButton"] svg {
        color: #ffffff !important;
        fill: #ffffff !important;
        opacity: 1 !important;
        visibility: visible !important;
        font-size: 1.35rem !important;
        font-weight: 700 !important;
    }
    
    /* 호버(PC) 및 터치 시 반전 효과 */
    [data-testid="stSidebarCollapseButton"] button:hover {
        background-color: #38bdf8 !important;
        border-color: #38bdf8 !important;
    }
    [data-testid="stSidebarCollapseButton"] button:hover * {
        color: #0f172a !important;
        fill: #0f172a !important;
    }

    /* 2. 사이드바 헤더 영역 패딩 및 정렬 보정 */
    [data-testid="stSidebarHeader"] {
        padding-top: 0.5rem !important;
        padding-bottom: 0.5rem !important;
    }

    /* 3. 사이드바가 닫혔을 때 다시 여는 버튼 (>>) 시인성 강화 */
    [data-testid="stSidebarCollapsedControl"] {
        visibility: visible !important;
        opacity: 1 !important;
    }
    
    [data-testid="stSidebarCollapsedControl"] button {
        background-color: #1e293b !important;
        border: 1.5px solid #38bdf8 !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.4), 0 0 6px rgba(56, 189, 248, 0.2) !important;
    }
    
    [data-testid="stSidebarCollapsedControl"] button *,
    [data-testid="stSidebarCollapsedControl"] span,
    [data-testid="stSidebarCollapsedControl"] [data-testid="stIconMaterial"],
    [data-testid="stSidebarCollapsedControl"] svg {
        color: #38bdf8 !important;
        fill: #38bdf8 !important;
        opacity: 1 !important;
        visibility: visible !important;
        font-size: 1.35rem !important;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# 3. 데이터 로딩 캐시 함수
# ==========================================
@st.cache_data(ttl=3600, show_spinner=False)
def get_cached_etf_data(target_date_key: str, force_refresh: bool = False):
    return data_loader.load_etf_data(force_refresh=force_refresh)


@st.cache_data(ttl=600, show_spinner=False)
def get_cached_etf_history(ticker, market, months=12):
    try:
        return data_loader.load_etf_history(ticker, market, months=months)
    except Exception as e:
        print(f"get_cached_etf_history 오류: {e}")
        return pd.DataFrame(), pd.DataFrame(), 0.0, {}


# ==========================================
# 4. 세션 상태 초기화
# ==========================================
if 'market_selection' not in st.session_state:
    st.session_state.market_selection = "K Market"
if 'leverage_selection' not in st.session_state:
    st.session_state.leverage_selection = "1X"
if 'selected_etf_name' not in st.session_state:
    st.session_state.selected_etf_name = None
if 'force_reload' not in st.session_state:
    st.session_state.force_reload = False


# ==========================================
# 5. 왼쪽 사이드 패널 (사이드바)
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='color: #8AB4F8; font-size: 1.3rem; margin-top: 0;'>⚙️ 시장 & 배율 선택</h2>", unsafe_allow_html=True)

    # 1) 시장 선택 (한국 시장 (KRX) 디폴트, 미국 시장 (US))
    market_options = ["한국 시장 (KRX)", "미국 시장 (US)"]
    if st.session_state.market_selection == "K Market":
        st.session_state.market_selection = "한국 시장 (KRX)"
    elif st.session_state.market_selection == "US Market":
        st.session_state.market_selection = "미국 시장 (US)"
    current_mkt_idx = market_options.index(st.session_state.market_selection) if st.session_state.market_selection in market_options else 0
    market_choice = st.radio(
        "🏛️ 시장 선택",
        options=market_options,
        index=current_mkt_idx,
        horizontal=True,
        help="조회할 ETF 시장을 선택합니다. (한국거래소 상장 ETF, 미국 뉴욕/나스닥 상장 ETF)"
    )

    # 2) 배율 선택 (가로 2열 배치: 제1열 1X, 2X, 3X / 제2열 -1X, -2X, -3X)
    leverage_options = ["1X", "2X", "3X", "-1X", "-2X", "-3X"]
    if st.session_state.leverage_selection not in leverage_options:
        st.session_state.leverage_selection = "1X"
    current_lev_idx = leverage_options.index(st.session_state.leverage_selection) if st.session_state.leverage_selection in leverage_options else 0
    
    st.markdown("""
    <style>
    div[data-testid="stRadio"]:has(input[value="-1X"]) div[role="radiogroup"] {
        display: grid !important;
        grid-auto-flow: column !important;
        grid-template-rows: repeat(3, auto) !important;
        grid-template-columns: 1fr 1fr !important;
        gap: 6px 12px !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    leverage_choice = st.radio(
        "배율 선택 (Multiplier)",
        options=leverage_options,
        index=current_lev_idx,
        help="조회할 ETF의 레버리지 배율을 선택합니다. (제1열: 1X, 2X, 3X / 제2열: -1X, -2X, -3X)"
    )

    st.markdown("<hr style='border: 0; height: 1px; background-color: #334155; margin: 16px 0;'>", unsafe_allow_html=True)

    # 3) 스마트 필터 (사용자 특별 요청: 정렬 기준 1W~3Y 선택 기능)
    st.markdown("<div style='font-size: 0.95rem; font-weight: 700; color: #cbd5e1; margin-bottom: 8px;'>🎯 스마트 필터</div>", unsafe_allow_html=True)

    search_keyword = st.text_input(
        "종목명 / 티커 검색",
        placeholder="예: KODEX 200, 나스닥, SPY, SOXL, 반도체...",
        help="종목명 또는 코드/티커로 검색합니다."
    )

    sort_options = [
        "거래대금 높은순 (기본)",
        "1W(%) 수익률 높은순",
        "2W(%) 수익률 높은순",
        "1M(%) 수익률 높은순",
        "3M(%) 수익률 높은순",
        "6M(%) 수익률 높은순",
        "1Y(%) 수익률 높은순",
        "3Y(%) 수익률 높은순",
        "거래량 높은순",
        "현재가 높은순"
    ]

    selected_sort = st.selectbox(
        "정렬 기준 (스마트 정렬)",
        options=sort_options,
        index=0,
        help="원하는 정렬 기준을 선택하면 상위 100개 순위가 실시간으로 재정렬됩니다."
    )

    max_display_count = st.slider(
        "최대 표시 종목 수",
        min_value=10,
        max_value=100,
        value=100,
        step=10,
        help="데이터 테이블에 표시할 최대 ETF 개수 (최대 100개)"
    )

    st.markdown("<hr style='border: 0; height: 1px; background-color: #334155; margin: 16px 0;'>", unsafe_allow_html=True)

    # 4) Update & 조회 버튼 (다른 앱과의 레이아웃 통일: 왼쪽 Update, 오른쪽 조회)
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        btn_update = st.button("🔄 Update", use_container_width=True, help="최신 전일 종가 및 7대 기간 수익률을 다시 수집하고 캐시를 갱신합니다.")
    with col_btn2:
        btn_search = st.button("🔍 조회", type="primary", use_container_width=True, help="선택한 조건으로 대시보드를 조회합니다.")

    if btn_update:
        st.cache_data.clear()
        st.session_state.force_reload = True
        st.rerun()

    if btn_search or (market_choice != st.session_state.market_selection) or (leverage_choice != st.session_state.leverage_selection):
        st.session_state.market_selection = market_choice
        st.session_state.leverage_selection = leverage_choice
        st.session_state.selected_etf_name = None
        st.rerun()


# ==========================================
# 6. 데이터 로드 및 필터링
# ==========================================
force_refresh = st.session_state.force_reload
st.session_state.force_reload = False

target_business_date = data_loader.get_latest_business_date()
is_cached_today = data_loader.is_cache_available(target_business_date)

if not is_cached_today or force_refresh:
    spinner_msg = f"🔄 최신 영업일({target_business_date}) ETF 데이터를 자동으로 수집 및 정제하고 있습니다. 잠시만 기다려 주세요 (약 20~30초 소요)..."
else:
    spinner_msg = "한국 및 미국 증시 ETF 데이터를 불러오는 중입니다..."

with st.spinner(spinner_msg):
    df_raw, target_date, is_fallback = get_cached_etf_data(target_business_date, force_refresh=force_refresh)

if df_raw.empty:
    st.error("데이터를 불러올 수 없습니다. 인터넷 연결 및 데이터 소스 상태를 확인해 주세요.")
    st.stop()

# 1) 선택된 시장 필터링
active_market = st.session_state.market_selection
target_market = "K Market" if ("한국" in active_market or "KRX" in active_market or active_market == "K Market") else "US Market"
is_korean = (target_market == "K Market")
currency_symbol = "₩" if is_korean else "$"
currency_unit = "원" if is_korean else "달러"

df_market = df_raw[df_raw["시장"] == target_market].copy()

# 2) 선택된 배율 필터링
active_leverage = st.session_state.leverage_selection
df_filtered = df_market[df_market["배율"] == active_leverage].copy()

# 3) 한국 시장 3X / -3X 선택 시 규제 안내
is_kr_3x_warning = is_korean and (active_leverage in ["3X", "-3X"])

# 4) 스마트 필터 (키워드 검색)
if search_keyword.strip():
    kw = search_keyword.strip().lower()
    df_filtered = df_filtered[
        df_filtered["종목명"].astype(str).str.lower().str.contains(kw) |
        df_filtered["코드/티커"].astype(str).str.lower().str.contains(kw)
    ]

# 5) 정렬 기준 적용 (기본: 거래대금 높은순, 사이드바 정렬 옵션 지원)
sort_map = {
    "거래대금 높은순 (기본)": ("거래대금", False),
    "1W(%) 수익률 높은순": ("1W(%)", False),
    "2W(%) 수익률 높은순": ("2W(%)", False),
    "1M(%) 수익률 높은순": ("1M(%)", False),
    "3M(%) 수익률 높은순": ("3M(%)", False),
    "6M(%) 수익률 높은순": ("6M(%)", False),
    "1Y(%) 수익률 높은순": ("1Y(%)", False),
    "3Y(%) 수익률 높은순": ("3Y(%)", False),
    "거래량 높은순": ("거래량", False),
    "현재가 높은순": ("현재가", False),
}

sort_col, sort_asc = sort_map.get(selected_sort, ("거래대금", False))
if sort_col in df_filtered.columns:
    df_filtered = df_filtered.sort_values(by=sort_col, ascending=sort_asc, na_position='last').reset_index(drop=True)

# 상위 N개 슬라이싱 (최대 100개)
df_display_source = df_filtered.head(max_display_count).copy().reset_index(drop=True)

# 순위 재부여 (1부터 100까지 명확히 표시)
df_display_source["순위"] = list(range(1, len(df_display_source) + 1))


# ==========================================
# 7. 메인 타이틀 영역
# ==========================================
st.markdown(
    "<h1 style='text-align: center; color: #8AB4F8 !important; font-weight: 800; font-size: 2.0rem; margin-top: 0; margin-bottom: 0.3rem; letter-spacing: -0.5px;'>"
    "한국 및 미국 증시 ETF 수익률 비교"
    "</h1>",
    unsafe_allow_html=True
)

# 메타 정보 표시줄
st.markdown(
    f"<div style='text-align: center; font-size: 0.86rem; color: #94a3b8; margin-bottom: 12px;'>"
    f"기준일: <span style='color: #38bdf8; font-weight: 600;'>{target_date}</span> (전일 종가 기준) &nbsp;|&nbsp; "
    f"선택 시장: <span style='color: #f8fafc; font-weight: 700;'>{active_market}</span> &nbsp;|&nbsp; "
    f"선택 배율: <span style='color: #fbbf24; font-weight: 700;'>{active_leverage}</span> &nbsp;|&nbsp; "
    f"통화 단위: <span style='color: #34d399; font-weight: 600;'>{'원화(KRW, ₩)' if is_korean else '달러(USD, $)'}</span> &nbsp;|&nbsp; "
    f"데이터 출처: <span style='color: #cbd5e1;'>{'네이버 금융 / KRX' if is_korean else 'S&P / NASDAQ / Yahoo Finance'}</span>"
    f"</div>",
    unsafe_allow_html=True
)

# 데이터 수집 지연/Fallback 안내 배너
if is_fallback:
    st.warning(
        f"⚠️ **외부 통신 지연 안내**: 최신 영업일({target_business_date}) 데이터의 자동 수집이 지연되어 "
        f"기존 저장된 데이터({target_date} 기준)가 표시되고 있습니다. 최신 데이터로 업데이트를 원하시면 "
        f"왼쪽 사이드바의 **'🔄 Update'** 버튼을 클릭해 주세요."
    )

# 한국 3X/-3X 규제 예외 안내 배너
if is_kr_3x_warning:
    st.info(
        "ℹ️ **국내 자본시장 규제 안내**: 국내 자본시장법 및 한국거래소(KRX) 규정에 따라 한국 시장 상장 ETF는 "
        "레버리지 배율이 **최대 ±2배(2X, -2X)** 로 엄격히 제한되어 있어 **3X 및 -3X ETF가 상장되어 있지 않습니다**.\n\n"
        "👉 **3X / -3X(TQQQ, SOXL, SQQQ, SOXS 등)** 고배율 ETF 투자를 원하시면 왼쪽 사이드바에서 **'US Market'**을 선택해 주세요."
    )

st.markdown("<hr style='border: 0; height: 1px; background-color: #334155; margin: 8px 0 16px 0;'>", unsafe_allow_html=True)


# ==========================================
# 8. 핵심 요약 KPI 지표 카드 (4열)
# ==========================================
col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)

total_etf_cnt = len(df_display_source)
avg_1m_ret = df_display_source["1M(%)"].dropna().mean() if total_etf_cnt > 0 else 0.0
avg_trade_val = df_display_source["거래대금"].dropna().mean() if total_etf_cnt > 0 else 0.0

if total_etf_cnt > 0 and df_display_source["1M(%)"].notna().any():
    top_1m_idx = df_display_source["1M(%)"].idxmax()
    top_1m_row = df_display_source.loc[top_1m_idx]
    top_etf_str = f"{top_1m_row['종목명'][:12]}.. ({top_1m_row['1M(%)']:+.2f}%)"
else:
    top_etf_str = "-"

with col_kpi1:
    with st.container(border=True):
        st.metric("📊 조회된 ETF 수", f"{total_etf_cnt:,} 개", help=f"{active_market} [{active_leverage}] 조건 부합 ETF")

with col_kpi2:
    with st.container(border=True):
        st.metric("📈 평균 1개월 수익률", f"{avg_1m_ret:+.2f} %", help="조회된 ETF들의 1M 평균 수익률")

with col_kpi3:
    with st.container(border=True):
        if is_korean:
            val_str = f"{avg_trade_val / 1e8:,.1f} 억원"
        else:
            val_str = f"${avg_trade_val / 1e6:,.2f} M"
        st.metric("💰 평균 일일 거래대금", val_str, help="조회된 ETF들의 전일 평균 거래대금")

with col_kpi4:
    with st.container(border=True):
        st.metric("🏆 1개월 최고 수익률 ETF", top_etf_str, help="조회된 목록 중 1개월 수익률 1위")


# ==========================================
# 9. 메인 영역: 조회 결과 데이터 테이블 (TOP 100)
# ==========================================
col_tbl_title, col_dl = st.columns([8, 2], vertical_alignment="bottom")
with col_tbl_title:
    st.markdown(
        f"<div class='section-header'>📋 {active_market} [{active_leverage}] 수익률 비교 TOP {min(100, len(df_display_source))} 데이터 테이블</div>",
        unsafe_allow_html=True
    )
with col_dl:
    # 엑셀 다운로드 파일 생성
    excel_bytes = data_loader.create_excel_download(df_display_source, active_market, active_leverage)
    file_name = f"ETF_TOP100_{active_market.replace(' ', '_')}_{active_leverage}_{target_date.replace('-', '')}.xlsx"
    st.download_button(
        label="📥 엑셀 파일 다운로드",
        data=excel_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# 표시용 데이터프레임 가공
df_table = df_display_source.copy()

# 통화별 가격 및 거래대금 포맷 설정
if is_korean:
    price_format = "%,.0f 원"
    trade_val_format = "%,.0f 원"
else:
    price_format = "$%,.2f"
    trade_val_format = "$%,.0f"

# 15개 요구 항목 정의
table_columns = [
    "순위", "코드/티커", "종목명", "시장", "배율",
    "현재가", "거래량", "거래대금",
    "1W(%)", "2W(%)", "1M(%)", "3M(%)", "6M(%)", "1Y(%)", "3Y(%)"
]

# st.column_config 헤더 오름차순/내림차순 토글 및 서식 완벽 지원
column_config = {
    "순위": st.column_config.NumberColumn("순위", width=50, format="%d", pinned=True),
    "코드/티커": st.column_config.TextColumn("코드/티커", width=85),
    "종목명": st.column_config.TextColumn("종목명", width="medium"),
    "시장": st.column_config.TextColumn("시장", width="small"),
    "배율": st.column_config.TextColumn("배율", width="small"),
    "현재가": st.column_config.NumberColumn("현재가", format=price_format, width="small"),
    "거래량": st.column_config.NumberColumn("거래량", format="%,.0f", width="small"),
    "거래대금": st.column_config.NumberColumn("거래대금", format=trade_val_format, width="small"),
    "1W(%)": st.column_config.NumberColumn("1W(%)", format="%+.2f%%", width="small"),
    "2W(%)": st.column_config.NumberColumn("2W(%)", format="%+.2f%%", width="small"),
    "1M(%)": st.column_config.NumberColumn("1M(%)", format="%+.2f%%", width="small"),
    "3M(%)": st.column_config.NumberColumn("3M(%)", format="%+.2f%%", width="small"),
    "6M(%)": st.column_config.NumberColumn("6M(%)", format="%+.2f%%", width="small"),
    "1Y(%)": st.column_config.NumberColumn("1Y(%)", format="%+.2f%%", width="small"),
    "3Y(%)": st.column_config.NumberColumn("3Y(%)", format="%+.2f%%", width="small"),
}

# 인터랙티브 테이블 렌더링 (헤더 정렬 토글 지원, 행 클릭 선택 지원)
selection = st.dataframe(
    df_table[table_columns],
    use_container_width=True,
    height=420,
    hide_index=True,
    column_config=column_config,
    on_select="rerun",
    selection_mode="single-row"
)

# 테이블에서 선택된 행 추출
selected_from_table = None
if selection and selection.get("rows"):
    sel_idx = selection["rows"][0]
    if sel_idx < len(df_table):
        selected_from_table = df_table.iloc[sel_idx]["종목명"]
        st.session_state.selected_etf_name = selected_from_table


# ==========================================
# 10. 가로 구분선
# ==========================================
st.markdown("<hr style='border: 0; height: 1px; background-color: #334155; margin: 25px 0 20px 0;'>", unsafe_allow_html=True)


# ==========================================
# 11. 메인 영역: 선택한 종목의 상세 정보 및 시각화 차트
# ==========================================
st.markdown("<div class='section-header'>📈 선택한 ETF 상세 정보 및 시각화 차트</div>", unsafe_allow_html=True)

etf_name_options = df_display_source["종목명"].tolist()

if not etf_name_options:
    st.info("조건에 부합하는 ETF가 없습니다. 왼쪽 사이드바의 필터를 변경해 보세요.")
    st.stop()

# 디폴트 인덱스 계산 (테이블 선택 우선, 없으면 1위 종목)
if st.session_state.selected_etf_name in etf_name_options:
    default_etf_idx = etf_name_options.index(st.session_state.selected_etf_name)
elif selected_from_table and selected_from_table in etf_name_options:
    default_etf_idx = etf_name_options.index(selected_from_table)
else:
    default_etf_idx = 0

col_sel1, col_sel2 = st.columns([7, 5])
with col_sel1:
    chosen_etf_name = st.selectbox(
        "분석할 ETF 종목 선택",
        options=etf_name_options,
        index=default_etf_idx,
        help="상단 테이블에서 행을 클릭하거나 목록에서 종목을 직접 선택할 수 있습니다."
    )

target_row = df_display_source[df_display_source["종목명"] == chosen_etf_name].iloc[0]
chosen_ticker = str(target_row["코드/티커"])
chosen_mkt = target_row["시장"]
chosen_lev = target_row["배율"]

with col_sel2:
    period_label = st.radio(
        "차트 분석 기간",
        options=["3개월", "6개월", "1년", "3년"],
        index=2, # 디폴트 1년
        horizontal=True,
        help="시계열 차트 및 벤치마크 비교 기간을 설정합니다."
    )
    period_map = {"3개월": 3, "6개월": 6, "1년": 12, "3년": 36}
    chosen_months = period_map[period_label]

# 시계열 주가, 벤치마크, 기술적 지표 로드
with st.spinner(f"'{chosen_etf_name}' 시계열 차트 데이터를 불러오는 중..."):
    df_history, df_benchmark, mdd_val, stats = get_cached_etf_history(
        chosen_ticker,
        chosen_mkt,
        months=chosen_months
    )

# 6열 상세 지표 요약 카드
with st.container(border=True):
    col_m1, col_m2, col_m3, col_m4, col_m5, col_m6 = st.columns(6)
    with col_m1:
        st.metric("종목명 (코드)", f"{chosen_etf_name}", help=f"티커: {chosen_ticker} | 시장: {chosen_mkt}")
    with col_m2:
        curr_price_str = f"{target_row['현재가']:,.0f}원" if is_korean else f"${target_row['현재가']:,.2f}"
        st.metric("현재가 (전일종가)", curr_price_str)
    with col_m3:
        if is_korean:
            tval_str = f"{target_row['거래대금'] / 1e8:,.1f} 억원"
        else:
            tval_str = f"${target_row['거래대금'] / 1e6:,.2f} M"
        st.metric("거래대금 (거래량)", tval_str, help=f"거래량: {target_row['거래량']:,} 주")
    with col_m4:
        ret_1m = target_row["1M(%)"]
        ret_1m_str = f"{ret_1m:+.2f}%" if pd.notna(ret_1m) else "N/A"
        st.metric("1개월 수익률", ret_1m_str)
    with col_m5:
        ret_1y = target_row["1Y(%)"]
        ret_1y_str = f"{ret_1y:+.2f}%" if pd.notna(ret_1y) else "N/A"
        st.metric("1년 수익률", ret_1y_str)
    with col_m6:
        st.metric("배율 / 최대낙폭(MDD)", f"{chosen_lev} ({mdd_val:.1f}%)", help=f"선택 기간 내 최고점 대비 최대 낙폭: {mdd_val:.2f}%")


# ==========================================
# 12. 4대 인터랙티브 시각화 차트 (2x2 Grid)
# ==========================================
col_ch1, col_ch2 = st.columns(2)

# ----------------------------------------------------
# [차트 1]: 주가 및 3대 이동평균선(20/60/120) & 거래량 복합 차트
# ----------------------------------------------------
with col_ch1:
    fig1 = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.08,
        subplot_titles=[f"📊 주가 및 이동평균선 ({period_label})", "거래량"],
        row_heights=[0.75, 0.25]
    )

    if not df_history.empty:
        date_strs = [d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else str(d)[:10] for d in df_history.index]

        # 종가 라인
        fig1.add_trace(go.Scatter(
            x=date_strs,
            y=df_history['종가'],
            mode='lines',
            name="주가(종가)",
            line=dict(color='#38bdf8', width=2),
            hovertemplate=f"%{{y:,.0f}}{currency_unit}" if is_korean else f"$%{{y:,.2f}}"
        ), row=1, col=1)

        # 20일 이평선
        if 'MA20' in df_history.columns:
            fig1.add_trace(go.Scatter(
                x=date_strs,
                y=df_history['MA20'],
                mode='lines',
                name="20일 이평선",
                line=dict(color='#f59e0b', width=1.4, dash='dot'),
                hovertemplate=f"%{{y:,.0f}}{currency_unit}" if is_korean else f"$%{{y:,.2f}}"
            ), row=1, col=1)

        # 60일 이평선
        if 'MA60' in df_history.columns:
            fig1.add_trace(go.Scatter(
                x=date_strs,
                y=df_history['MA60'],
                mode='lines',
                name="60일 이평선",
                line=dict(color='#a855f7', width=1.4, dash='dash'),
                hovertemplate=f"%{{y:,.0f}}{currency_unit}" if is_korean else f"$%{{y:,.2f}}"
            ), row=1, col=1)

        # 120일 이평선
        if 'MA120' in df_history.columns:
            fig1.add_trace(go.Scatter(
                x=date_strs,
                y=df_history['MA120'],
                mode='lines',
                name="120일 이평선",
                line=dict(color='#10b981', width=1.4, dash='dashdot'),
                hovertemplate=f"%{{y:,.0f}}{currency_unit}" if is_korean else f"$%{{y:,.2f}}"
            ), row=1, col=1)

        # 거래량 바
        vol_colors = ['#ef4444' if df_history['종가'].iloc[i] >= df_history['종가'].iloc[i-1] else '#3b82f6' for i in range(len(df_history))]
        fig1.add_trace(go.Bar(
            x=date_strs,
            y=df_history['거래량'],
            name="거래량",
            marker=dict(color=vol_colors, opacity=0.7),
            hovertemplate="%{y:,.0f}주" if is_korean else "%{y:,.0f}"
        ), row=2, col=1)

        # 주말 및 공휴일 공백 제거 (5일 주기 끊김 및 0값 방지)
        dt_all = pd.date_range(start=df_history.index[0], end=df_history.index[-1], freq='B')
        existing_dates = set(pd.to_datetime(df_history.index).normalize())
        holidays = [d.strftime('%Y-%m-%d') for d in dt_all if d.normalize() not in existing_dates]
        rbreaks = [dict(bounds=["sat", "mon"])]
        if holidays:
            rbreaks.append(dict(values=holidays))
        fig1.update_xaxes(rangebreaks=rbreaks)

    fig1.update_layout(
        hovermode="x unified",
        hoverlabel=dict(
            bgcolor="#1e293b",
            bordercolor="#475569",
            font=dict(color="#f8fafc", size=12, family="Malgun Gothic, -apple-system, sans-serif")
        ),
        template="plotly_dark",
        paper_bgcolor="#1e293b",
        plot_bgcolor="#0f172a",
        margin=dict(l=40, r=20, t=40, b=30),
        height=380,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        xaxis=dict(showgrid=True, gridcolor="#334155"),
        yaxis=dict(showgrid=True, gridcolor="#334155", hoverformat=",.0f" if is_korean else "$,.2f"),
        xaxis2=dict(showgrid=True, gridcolor="#334155"),
        yaxis2=dict(showgrid=True, gridcolor="#334155", hoverformat=",.0f"),
    )
    st.plotly_chart(fig1, use_container_width=True)

# ----------------------------------------------------
# [차트 2]: 7대 구간 수익률(1W, 2W, 1M, 3M, 6M, 1Y, 3Y) 바 차트
# ----------------------------------------------------
with col_ch2:
    periods = ['1W', '2W', '1M', '3M', '6M', '1Y', '3Y']
    ret_values = [
        target_row['1W(%)'],
        target_row['2W(%)'],
        target_row['1M(%)'],
        target_row['3M(%)'],
        target_row['6M(%)'],
        target_row['1Y(%)'],
        target_row['3Y(%)'],
    ]

    clean_vals = [round(float(v), 2) if pd.notna(v) else 0.0 for v in ret_values]
    bar_colors = ['#ef4444' if v >= 0 else '#3b82f6' for v in clean_vals]

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=periods,
        y=clean_vals,
        text=[f"{v:+.2f}%" if pd.notna(r) else "N/A" for v, r in zip(clean_vals, ret_values)],
        textposition="outside",
        marker=dict(color=bar_colors, line=dict(width=1, color='#475569')),
        hovertemplate="<b>기간: %{x}</b><br>수익률: %{y:+.2f}%<extra></extra>"
    ))

    fig2.update_layout(
        title=f"📈 7대 기간별 수익률 프로필 ({chosen_etf_name})",
        title_font=dict(size=13, color='#e2e8f0'),
        hoverlabel=dict(
            bgcolor="#1e293b",
            bordercolor="#475569",
            font=dict(color="#f8fafc", size=12, family="Malgun Gothic, -apple-system, sans-serif")
        ),
        template="plotly_dark",
        paper_bgcolor="#1e293b",
        plot_bgcolor="#0f172a",
        margin=dict(l=40, r=20, t=50, b=30),
        height=380,
        xaxis=dict(title="투자 기간", showgrid=False),
        yaxis=dict(title="수익률 (%)", showgrid=True, gridcolor="#334155", zeroline=True, zerolinecolor="#64748b", zerolinewidth=1.5, hoverformat="+.2f"),
    )
    st.plotly_chart(fig2, use_container_width=True)

col_ch3, col_ch4 = st.columns(2)

# ----------------------------------------------------
# [차트 3]: 벤치마크 지수 대비 누적 수익률 비교 차트
# ----------------------------------------------------
with col_ch3:
    fig3 = go.Figure()
    bm_name = stats.get('bm_name', '벤치마크 지수')

    if not df_history.empty and '누적수익률' in df_history.columns:
        date_strs = [d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else str(d)[:10] for d in df_history.index]
        cum_ret_series = df_history['누적수익률'].round(2)
        fig3.add_trace(go.Scatter(
            x=date_strs,
            y=cum_ret_series,
            mode='lines',
            name=f"{chosen_etf_name}",
            line=dict(color='#38bdf8', width=2.2),
            hovertemplate="%{y:+.2f}%"
        ))

    if not df_benchmark.empty and '누적수익률' in df_benchmark.columns:
        bm_dates = [d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else str(d)[:10] for d in df_benchmark.index]
        bm_ret_series = df_benchmark['누적수익률'].round(2)
        fig3.add_trace(go.Scatter(
            x=bm_dates,
            y=bm_ret_series,
            mode='lines',
            name=f"{bm_name}",
            line=dict(color='#94a3b8', width=1.5, dash='dash'),
            hovertemplate="%{y:+.2f}%"
        ))

    fig3.update_layout(
        title=f"⚖️ 벤치마크 대비 누적 수익률 추이 ({period_label})",
        title_font=dict(size=13, color='#e2e8f0'),
        hovermode="x unified",
        hoverlabel=dict(
            bgcolor="#1e293b",
            bordercolor="#475569",
            font=dict(color="#f8fafc", size=12, family="Malgun Gothic, -apple-system, sans-serif")
        ),
        template="plotly_dark",
        paper_bgcolor="#1e293b",
        plot_bgcolor="#0f172a",
        margin=dict(l=40, r=20, t=50, b=30),
        height=380,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        xaxis=dict(showgrid=True, gridcolor="#334155"),
        yaxis=dict(
            title="누적 수익률 (%)",
            showgrid=True,
            gridcolor="#334155",
            zeroline=True,
            zerolinecolor="#64748b",
            hoverformat="+.2f",
            tickformat="+.1f"
        ),
    )
    st.plotly_chart(fig3, use_container_width=True)

# ----------------------------------------------------
# [차트 4]: 최대 낙폭(MDD, Drawdown) 및 리스크 분석 차트
# ----------------------------------------------------
with col_ch4:
    fig4 = go.Figure()

    if not df_history.empty and '낙폭(Drawdown)' in df_history.columns:
        date_strs = [d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else str(d)[:10] for d in df_history.index]
        dd_series = df_history['낙폭(Drawdown)'].round(2)
        fig4.add_trace(go.Scatter(
            x=date_strs,
            y=dd_series,
            mode='lines',
            name="고점 대비 낙폭",
            line=dict(color='#ef4444', width=1.5),
            fill='tozeroy',
            fillcolor='rgba(239, 68, 68, 0.25)',
            hovertemplate="%{y:.2f}%"
        ))

    fig4.update_layout(
        title=f"⚠️ 고점 대비 낙폭(Drawdown) 및 MDD ({period_label}) [MDD: {mdd_val:.2f}%]",
        title_font=dict(size=13, color='#e2e8f0'),
        hovermode="x unified",
        hoverlabel=dict(
            bgcolor="#1e293b",
            bordercolor="#475569",
            font=dict(color="#f8fafc", size=12, family="Malgun Gothic, -apple-system, sans-serif")
        ),
        template="plotly_dark",
        paper_bgcolor="#1e293b",
        plot_bgcolor="#0f172a",
        margin=dict(l=40, r=20, t=50, b=30),
        height=380,
        showlegend=False,
        xaxis=dict(showgrid=True, gridcolor="#334155"),
        yaxis=dict(
            title="고점 대비 하락률 (%)",
            showgrid=True,
            gridcolor="#334155",
            zeroline=True,
            zerolinecolor="#64748b",
            hoverformat=".2f",
            tickformat=".1f"
        ),
    )
    st.plotly_chart(fig4, use_container_width=True)


# ==========================================
# 13. 전문가 종합 ETF 투자 진단 카드
# ==========================================
st.markdown("<hr style='border: 0; height: 1px; background-color: #334155; margin: 20px 0 15px 0;'>", unsafe_allow_html=True)
st.markdown("<div class='section-header'>💡 전문가 종합 ETF 투자 진단 보고서</div>", unsafe_allow_html=True)

col_d1, col_d2, col_d3 = st.columns(3)

# 1) 배율 및 포지션 적합성 진단
with col_d1:
    with st.container(border=True):
        st.markdown("<div style='font-size: 0.95rem; font-weight: 700; color: #8AB4F8;'>🎯 배율 및 보유 전략 진단</div>", unsafe_allow_html=True)
        if chosen_lev in ["2X", "3X"]:
            st.markdown(
                f"<span class='badge-bull'>{chosen_lev} 레버리지 상품</span><br><br>"
                "강한 상승 모멘텀에서 수익을 극대화하는 상품입니다. 단, 횡보장에서는 **음의 복리(Volatility Decay)** 현상으로 지수가 제자리여도 원금 손실이 누적될 수 있으므로 중장기 적립식보다는 **단기 추세추종 매매**에 적합합니다.",
                unsafe_allow_html=True
            )
        elif chosen_lev in ["-1X", "-2X", "-3X"]:
            st.markdown(
                f"<span class='badge-bear'>{chosen_lev} 인버스/숏 상품</span><br><br>"
                "시장 하락장에서 수익을 창출하거나 기존 주식 포트폴리오의 **위험 헷지(Hedging)** 목적에 최적화되어 있습니다. 장기 보유 시 증시 우상향에 따른 손실 위험이 있으므로 목표 수익률 도달 시 빠른 차익 실현이 권장됩니다.",
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                f"<span class='badge-neutral'>1X 정방향 기본형</span><br><br>"
                "기초지수를 1:1로 추종하여 변동성 잠식(Decay) 위험이 없으며, **중장기 가치투자 및 연금/적립식 투자**에 가장 이상적이고 안전한 구조입니다.",
                unsafe_allow_html=True
            )

# 2) 유동성 및 거래대금 진단
with col_d2:
    with st.container(border=True):
        st.markdown("<div style='font-size: 0.95rem; font-weight: 700; color: #34d399;'>💧 유동성 및 매매 환경 진단</div>", unsafe_allow_html=True)
        t_val = target_row["거래대금"]
        if is_korean:
            if t_val >= 100_000_000_000: # 1000억 이상
                st.write("🟢 **최상급 유동성 (초대형 ETF)**: 일일 거래대금이 1,000억 원을 초과하여 LP(유동성공급자) 호가 스프레드가 매우 촘촘하며 대량 매매 시에도 슬리피지가 거의 발생하지 않습니다.")
            elif t_val >= 10_000_000_000: # 100억 이상
                st.write("🟢 **우수한 유동성**: 일일 거래대금이 100억 원 이상으로 일반 개인 및 기관 투자자가 원활하게 매매하기에 충분한 거래량을 확보하고 있습니다.")
            else:
                st.write("🟡 **보통 유동성**: 거래량이 비교적 적을 수 있으므로 시장가 주문보다는 호가창을 확인한 후 지정가 매매를 권장합니다.")
        else:
            if t_val >= 1_000_000_000: # $1B 이상
                st.write("🟢 **글로벌 메가 유동성**: 일일 거래대금이 10억 달러($1B)를 초과하는 글로벌 메이저 ETF로 스프레드가 1센트 수준에 불과합니다.")
            elif t_val >= 100_000_000: # $100M 이상
                st.write("🟢 **원활한 유동성**: 기관 및 글로벌 투자자의 거래가 활발하여 슬리피지 없이 자유로운 진입/청산이 가능합니다.")
            else:
                st.write("🟡 **소형 ETF**: 거래 스프레드 및 장전/장후 시간외 매매 시 유동성을 체크하시기 바랍니다.")

# 3) 모멘텀 및 추세 진단
with col_d3:
    with st.container(border=True):
        st.markdown("<div style='font-size: 0.95rem; font-weight: 700; color: #fbbf24;'>⚡ 모멘텀 및 가격 위치 진단</div>", unsafe_allow_html=True)
        h_diff = stats.get('high_diff', 0.0)
        ret_1m = target_row['1M(%)']
        ret_6m = target_row['6M(%)']

        if pd.notna(ret_1m) and ret_1m > 5 and pd.notna(ret_6m) and ret_6m > 10:
            st.write(f"🚀 **강력한 상승 모멘텀**: 1개월({ret_1m:+.1f}%)과 6개월({ret_6m:+.1f}%) 수익률이 모두 우상향 중인 정배열 상승 국면입니다. 52주 고점 대비 괴리율은 {h_diff:.1f}%입니다.")
        elif pd.notna(ret_1m) and ret_1m < -5 and pd.notna(ret_6m) and ret_6m < -10:
            st.write(f"⚠️ **하락 추세 지속**: 단기 및 중장기 수익률이 모두 약세를 보이고 있습니다. 역배열 구간이므로 바닥 확인 후 분할 매수를 검토하세요. (52주 고점 대비 {h_diff:.1f}%)")
        else:
            st.write(f"⚖️ **박스권 횡보 국면**: 단기({ret_1m:+.1f}%)와 중장기 수익률이 혼조세를 보이며 매물대를 소화하고 있습니다. 지지선과 저항선을 기준으로 박스권 매매 전략이 유효합니다.")

# 하단 투자 유의사항 공통 푸터
st.markdown("<hr style='border: 0; height: 1px; background-color: #334155; margin: 30px 0 10px 0;'>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align: center; color: #64748b; font-size: 0.8rem; margin-top: 8px; margin-bottom: 24px; line-height: 1.6;'>"
    "⚠️ 본 서비스에서 제공하는 모든 정보는 투자 참고용이며, 투자의 최종 결정과 책임은 투자자 본인에게 있습니다."
    "</div>",
    unsafe_allow_html=True
)
