"""
build_master_data.py
한국(K Market) 및 미국(US Market) 증시 ETF 유니버스의 전 배율(3X, 2X, 1X, -1X, -2X, -3X)
전일 종가 기준 현재가, 거래량, 거래대금 및 7대 기간(1W, 2W, 1M, 3M, 6M, 1Y, 3Y) 수익률을 수집/계산하여
etf_master_data.csv 및 cache/ 디렉토리에 마스터 데이터를 구축하는 스크립트.
"""

import os
import sys
import time
import datetime
from concurrent.futures import ThreadPoolExecutor
import requests
import json
import pandas as pd
import numpy as np

# UTF-8 출력 인코딩 설정
if sys.stdout:
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_FILE = os.path.join(CURRENT_DIR, "etf_master_data.csv")
CACHE_DIR = os.path.join(CURRENT_DIR, "cache")
os.makedirs(CACHE_DIR, exist_ok=True)

# 라이브러리 로드
import yfinance as yf
try:
    import FinanceDataReader as fdr
except Exception:
    fdr = None


# ==========================================
# 1. 영업일 및 기준일 유틸리티
# ==========================================
def get_latest_business_date():
    """
    가장 최근 거래 완료된 영업일 YYYY-MM-DD 반환.
    - 한국 장 마감 시간(15:30) 고려하여 평일 16:00 이전에는 전일(또는 직전 평일)
    - 평일 16:00 이후에는 당일
    - 주말은 직전 금요일
    """
    now_utc = datetime.datetime.now(datetime.timezone.utc)
    now_kst = now_utc + datetime.timedelta(hours=9)
    start_offset = 0 if (now_kst.weekday() < 5 and now_kst.hour >= 16) else 1

    for i in range(start_offset, start_offset + 10):
        d = now_kst - datetime.timedelta(days=i)
        if d.weekday() < 5:
            return d.strftime('%Y-%m-%d')
    return (now_kst - datetime.timedelta(days=1)).strftime('%Y-%m-%d')


# ==========================================
# 2. 한국 시장 (K Market) ETF 배율 분류 및 수집
# ==========================================
def classify_kr_leverage(name: str) -> str:
    """한국 ETF 종목명을 기반으로 배율을 분류"""
    # 3X / -3X 여부 검사 (규제상 거의 없으나 확인)
    if '3X' in name or '3x' in name or '3배' in name:
        if '인버스' in name or 'Inverse' in name:
            return '-3X'
        return '3X'

    # -2X (인버스 2X, 곱버스)
    if ('인버스' in name or 'Inverse' in name) and ('2X' in name or '2x' in name or '2배' in name):
        return '-2X'

    # -1X (일반 인버스)
    if '인버스' in name or 'Inverse' in name:
        return '-1X'

    # 2X (레버리지)
    if '레버리지' in name or '2X' in name or '2x' in name or '2배' in name:
        return '2X'

    # 1X (정방향 1배수 기본형)
    return '1X'


def compute_period_returns(series_close: pd.Series):
    """
    시계열 종가 시리즈로부터 1W, 2W, 1M, 3M, 6M, 1Y, 3Y 수익률(%) 산출
    영업일 기준 매핑:
    - 1W: 5거래일 전 (-6 위치)
    - 2W: 10거래일 전 (-11 위치)
    - 1M: 21거래일 전 (-22 위치)
    - 3M: 63거래일 전 (-64 위치)
    - 6M: 126거래일 전 (-127 위치)
    - 1Y: 252거래일 전 (-253 위치)
    - 3Y: 756거래일 전 (-757 위치)
    """
    if series_close is None or len(series_close) < 2:
        return {k: np.nan for k in ['1W(%)', '2W(%)', '1M(%)', '3M(%)', '6M(%)', '1Y(%)', '3Y(%)']}

    p_curr = float(series_close.iloc[-1])
    n = len(series_close)

    def calc_ret(idx):
        if n >= idx + 1:
            p_past = float(series_close.iloc[-(idx + 1)])
            if p_past > 0:
                return round(((p_curr / p_past) - 1.0) * 100.0, 2)
        return np.nan

    return {
        '1W(%)': calc_ret(5),
        '2W(%)': calc_ret(10),
        '1M(%)': calc_ret(21),
        '3M(%)': calc_ret(63),
        '6M(%)': calc_ret(126),
        '1Y(%)': calc_ret(252),
        '3Y(%)': calc_ret(756),
    }


def build_kr_market_data(target_date: str = None):
    """
    네이버 금융 ETF 공식 API 및 FinanceDataReader를 활용하여
    K Market 전체 ETF 중 거래대금 상위 및 레버리지/인버스 종목 전수 데이터 수집.
    target_date: 영업일 기준일 (예: '2026-09-21'). 미지정 시 get_latest_business_date() 사용.
    """
    if not target_date:
        target_date = get_latest_business_date()

    print(f"[K Market] 한국 ETF 전체 목록 조회 중 (네이버 금융 API, 기준일: {target_date})...")
    url = "https://finance.naver.com/api/sise/etfItemList.nhn"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    try:
        res = requests.get(url, headers=headers, timeout=10).json()
        raw_items = res.get('result', {}).get('etfItemList', [])
    except Exception as e:
        print(f"[K Market] 네이버 API 호출 실패: {e}")
        return pd.DataFrame()

    print(f"[K Market] 총 {len(raw_items)}개 ETF 항목 파싱 완료.")

    # 한국 정규 거래 시간 여부 판별 (평일 09:00 ~ 15:30 KST만 '장중'으로 판정)
    # 그 외의 시간(평일 15:30~24:00, 평일 새벽 00:00~08:59, 주말 전체)은 장이 열리지 않은 마감 상태임!
    now_utc = datetime.datetime.now(datetime.timezone.utc)
    now_kst = now_utc + datetime.timedelta(hours=9)
    is_kr_trading_hours = (now_kst.weekday() < 5) and (
        (now_kst.hour == 9 and now_kst.minute >= 0) or
        (9 < now_kst.hour < 15) or
        (now_kst.hour == 15 and now_kst.minute <= 30)
    )

    # 1. 배율 분류 및 기본 데이터 매핑 (확정 종가 기준 산출)
    parsed_items = []
    for it in raw_items:
        code = str(it.get('itemcode', '')).zfill(6)
        name = it.get('itemname', '')
        now_val = float(it.get('nowVal', 0) or 0)
        change_val = float(it.get('changeVal', 0) or 0)

        if is_kr_trading_hours:
            # 평일 09:00~15:30 장중: 실시간 변동가이므로 전일 확정 종가 산출
            confirmed_price = (now_val - change_val) if (now_val > 0 and change_val is not None) else now_val
        else:
            # 장마감 후, 야간, 새벽, 주말: nowVal이 최근 마감 영업일의 확정 종가!
            confirmed_price = now_val

        quant = int(it.get('quant', 0) or 0)
        trade_val_won = float(it.get('amonut', 0) or 0) * 1_000_000
        leverage = classify_kr_leverage(name)
        mcap_won = float(it.get('marketSum', 0) or 0) * 100_000_000 # 억원 -> 원

        parsed_items.append({
            '코드/티커': code,
            '종목명': name,
            '시장': 'K Market',
            '배율': leverage,
            '현재가': confirmed_price,
            '거래량': quant,
            '거래대금': trade_val_won,
            '시가총액': mcap_won,
            '3M_api': float(it.get('threeMonthEarnRate', np.nan) or np.nan)
        })

    df_kr_base = pd.DataFrame(parsed_items)

    # 2. 시계열 수익률을 정밀 계산할 대상 선정
    # - 모든 레버리지 및 인버스 ETF (2X, -1X, -2X, 3X, -3X) 전수 포함
    # - 1X ETF 중 거래대금 상위 150개 포함 (거래대금 TOP 100을 충분히 커버)
    df_lev = df_kr_base[df_kr_base['배율'] != '1X']
    df_1x = df_kr_base[df_kr_base['배율'] == '1X'].sort_values(by='거래대금', ascending=False).head(160)
    df_targets = pd.concat([df_lev, df_1x]).drop_duplicates(subset=['코드/티커']).copy()

    target_codes = df_targets['코드/티커'].tolist()
    print(f"[K Market] 수익률 산출 대상 {len(target_codes)}개 ETF 시계열 병렬 수집 시작...")

    start_date = (datetime.datetime.now() - datetime.timedelta(days=365 * 3 + 60)).strftime('%Y-%m-%d')
    hist_returns = {}

    def fetch_kr_history(code):
        try:
            df_hist = fdr.DataReader(code, start_date)
            if df_hist is not None and not df_hist.empty:
                # 서버 타임존(UTC 등)과 무관하게, 명확한 기준일(target_date) 이하의 데이터만 정확히 슬라이싱
                df_target = df_hist.loc[:target_date]

                if len(df_target) >= 2:
                    p_close = float(df_target['Close'].iloc[-1])
                    p_vol = int(df_target['Volume'].iloc[-1])
                    rets = compute_period_returns(df_target['Close'])
                    return code, {
                        'prev_close': p_close,
                        'prev_vol': p_vol,
                        'returns': rets
                    }
        except Exception:
            pass
        return code, None

    with ThreadPoolExecutor(max_workers=10) as executor:
        results = list(executor.map(fetch_kr_history, target_codes))

    for code, info in results:
        if info:
            hist_returns[code] = info

    # 3. 수익률 병합
    records = []
    for _, row in df_kr_base.iterrows():
        c = row['코드/티커']
        r = hist_returns.get(c)
        if r:
            curr_p = r['prev_close']
            vol_p = r['prev_vol'] if r['prev_vol'] > 0 else row['거래량']
            val_p = row['거래대금'] if row['거래대금'] > 0 else float(curr_p * vol_p)
            rets = r['returns']
            ret_1w = rets['1W(%)']
            ret_2w = rets['2W(%)']
            ret_1m = rets['1M(%)']
            ret_3m = rets['3M(%)']
            ret_6m = rets['6M(%)']
            ret_1y = rets['1Y(%)']
            ret_3y = rets['3Y(%)']
        else:
            curr_p = row['현재가']
            vol_p = row['거래량']
            val_p = row['거래대금']
            ret_1w = np.nan
            ret_2w = np.nan
            ret_1m = np.nan
            ret_3m = row['3M_api'] if pd.notna(row['3M_api']) else np.nan
            ret_6m = np.nan
            ret_1y = np.nan
            ret_3y = np.nan

        records.append({
            '코드/티커': c,
            '종목명': row['종목명'],
            '시장': 'K Market',
            '배율': row['배율'],
            '현재가': curr_p,
            '거래량': vol_p,
            '거래대금': val_p,
            '1W(%)': ret_1w,
            '2W(%)': ret_2w,
            '1M(%)': ret_1m,
            '3M(%)': ret_3m,
            '6M(%)': ret_6m,
            '1Y(%)': ret_1y,
            '3Y(%)': ret_3y,
        })

    df_kr_final = pd.DataFrame(records)
    # 거래대금 기준 내림차순 정렬 후 순위 부여
    df_kr_final = df_kr_final.sort_values(by='거래대금', ascending=False).reset_index(drop=True)
    df_kr_final['순위'] = range(1, len(df_kr_final) + 1)

    print(f"[K Market] 최종 {len(df_kr_final)}개 ETF 정제 완료.")
    return df_kr_final


# ==========================================
# 3. 미국 시장 (US Market) ETF 배율 분류 및 수집
# ==========================================
US_ETF_UNIVERSE = [
    # --- 3X (Bull) ---
    ('TQQQ', 'ProShares UltraPro QQQ (나스닥100 3X)', '3X'),
    ('SOXL', 'Direxion Daily Semiconductor Bull 3X (반도체 3X)', '3X'),
    ('UPRO', 'ProShares UltraPro S&P500 (S&P500 3X)', '3X'),
    ('TECL', 'Direxion Daily Technology Bull 3X (테크 3X)', '3X'),
    ('FNGU', 'MicroSectors FANG+ Index 3X (빅테크 3X)', '3X'),
    ('BULZ', 'MicroSectors Solactive FANG+ 3X (FANG 3X)', '3X'),
    ('LABU', 'Direxion Daily S&P Biotech Bull 3X (바이오 3X)', '3X'),
    ('FAS', 'Direxion Daily Financial Bull 3X (금융 3X)', '3X'),
    ('TNA', 'Direxion Daily Small Cap Bull 3X (러셀2000 3X)', '3X'),
    ('DFEN', 'Direxion Daily Aerospace & Defense Bull 3X (방산 3X)', '3X'),
    ('CURE', 'Direxion Daily Healthcare Bull 3X (헬스케어 3X)', '3X'),
    ('NAIL', 'Direxion Daily Homebuilders & Supplies Bull 3X (건설 3X)', '3X'),
    ('RETL', 'Direxion Daily Retail Bull 3X (유통 3X)', '3X'),
    ('DPST', 'Direxion Daily Regional Banks Bull 3X (지역은행 3X)', '3X'),
    ('UTSL', 'Direxion Daily Utilities Bull 3X (유틸리티 3X)', '3X'),
    ('TMF', 'Direxion Daily 20+ Year Treasury Bull 3X (미국장기채 3X)', '3X'),
    ('WEBL', 'Direxion Daily Dow Jones Internet Bull 3X (인터넷 3X)', '3X'),
    ('UDOW', 'ProShares UltraPro Dow30 (다우30 3X)', '3X'),
    ('URTY', 'ProShares UltraPro Russell2000 (소형주 3X)', '3X'),
    ('MIDU', 'Direxion Daily Mid Cap Bull 3X (중형주 3X)', '3X'),
    ('YINN', 'Direxion Daily FTSE China Bull 3X (중국 3X)', '3X'),
    ('KORU', 'Direxion Daily MSCI South Korea Bull 3X (한국 3X)', '3X'),
    ('PILL', 'Direxion Daily Pharmaceutical Bull 3X (제약 3X)', '3X'),
    ('WANT', 'Direxion Daily Consumer Discretionary Bull 3X (소비재 3X)', '3X'),

    # --- -3X (Bear / Inverse) ---
    ('SQQQ', 'ProShares UltraPro Short QQQ (나스닥100 -3X)', '-3X'),
    ('SOXS', 'Direxion Daily Semiconductor Bear 3X (반도체 -3X)', '-3X'),
    ('SPXU', 'ProShares UltraPro Short S&P500 (S&P500 -3X)', '-3X'),
    ('TECS', 'Direxion Daily Technology Bear 3X (테크 -3X)', '-3X'),
    ('LABD', 'Direxion Daily S&P Biotech Bear 3X (바이오 -3X)', '-3X'),
    ('FAZ', 'Direxion Daily Financial Bear 3X (금융 -3X)', '-3X'),
    ('TZA', 'Direxion Daily Small Cap Bear 3X (러셀2000 -3X)', '-3X'),
    ('YANG', 'Direxion Daily FTSE China Bear 3X (중국 -3X)', '-3X'),
    ('TMV', 'Direxion Daily 20+ Year Treasury Bear 3X (장기채 -3X)', '-3X'),
    ('TTT', 'ProShares UltraPro Short 20+ Year Treasury (장기채 -3X)', '-3X'),
    ('EDZ', 'Direxion Daily MSCI Emerging Markets Bear 3X (신흥국 -3X)', '-3X'),
    ('SRTY', 'ProShares UltraPro Short Russell2000 (소형주 -3X)', '-3X'),
    ('SDOW', 'ProShares UltraPro Short Dow30 (다우30 -3X)', '-3X'),
    ('FNGD', 'MicroSectors FANG+ -3X Inverse (빅테크 -3X)', '-3X'),
    ('DRV', 'Direxion Daily Real Estate Bear 3X (부동산 -3X)', '-3X'),

    # --- 2X (Bull) ---
    ('QLD', 'ProShares Ultra QQQ (나스닥100 2X)', '2X'),
    ('SSO', 'ProShares Ultra S&P500 (S&P500 2X)', '2X'),
    ('USD', 'ProShares Ultra Semiconductors (반도체 2X)', '2X'),
    ('NVDL', 'GraniteShares 2x Long NVDA (엔비디아 2X)', '2X'),
    ('TSLL', 'Direxion Daily TSLA Bull 2X (테슬라 2X)', '2X'),
    ('AAPU', 'Direxion Daily AAPL Bull 2X (애플 2X)', '2X'),
    ('AMZU', 'Direxion Daily AMZN Bull 2X (아마존 2X)', '2X'),
    ('MSFU', 'Direxion Daily MSFT Bull 2X (마이크로소프트 2X)', '2X'),
    ('CONL', 'GraniteShares 2x Long COIN (코인베이스 2X)', '2X'),
    ('BITX', '2x Bitcoin Strategy ETF (비트코인 2X)', '2X'),
    ('AGQ', 'ProShares Ultra Silver (은 2X)', '2X'),
    ('UGL', 'ProShares Ultra Gold (금 2X)', '2X'),
    ('UCO', 'ProShares Ultra Bloomberg Crude Oil (원유 2X)', '2X'),
    ('BOIL', 'ProShares Ultra Bloomberg Natural Gas (천연가스 2X)', '2X'),
    ('BIB', 'ProShares Ultra Nasdaq Biotechnology (바이오 2X)', '2X'),
    ('DIG', 'ProShares Ultra Oil & Gas (에너지 2X)', '2X'),
    ('DDM', 'ProShares Ultra Dow30 (다우30 2X)', '2X'),
    ('SAA', 'ProShares Ultra SmallCap600 (소형주 2X)', '2X'),
    ('UWM', 'ProShares Ultra Russell2000 (러셀2000 2X)', '2X'),
    ('UST', 'ProShares Ultra 7-10 Year Treasury (중기채 2X)', '2X'),
    ('UBT', 'ProShares Ultra 20+ Year Treasury (장기채 2X)', '2X'),
    ('ROM', 'ProShares Ultra Technology (테크 2X)', '2X'),
    ('CWEB', 'Direxion Daily CSI China Internet 2X (중국인터넷 2X)', '2X'),
    ('INDL', 'Direxion Daily MSCI India Bull 2X (인도 2X)', '2X'),
    ('BRZU', 'Direxion Daily MSCI Brazil Bull 2X (브라질 2X)', '2X'),
    ('GUSH', 'Direxion Daily S&P Oil & Gas Bull 2X (오일가스 2X)', '2X'),

    # --- -2X (Bear / Inverse) ---
    ('QID', 'ProShares UltraShort QQQ (나스닥100 -2X)', '-2X'),
    ('SDS', 'ProShares UltraShort S&P500 (S&P500 -2X)', '-2X'),
    ('DXD', 'ProShares UltraShort Dow30 (다우30 -2X)', '-2X'),
    ('TWM', 'ProShares UltraShort Russell2000 (러셀2000 -2X)', '-2X'),
    ('DUG', 'ProShares UltraShort Oil & Gas (에너지 -2X)', '-2X'),
    ('SMN', 'ProShares UltraShort Basic Materials (원자재 -2X)', '-2X'),
    ('SCO', 'ProShares UltraShort Bloomberg Crude Oil (원유 -2X)', '-2X'),
    ('KOLD', 'ProShares UltraShort Bloomberg Natural Gas (가스 -2X)', '-2X'),
    ('ZSL', 'ProShares UltraShort Silver (은 -2X)', '-2X'),
    ('GLL', 'ProShares UltraShort Gold (금 -2X)', '-2X'),
    ('TBT', 'ProShares UltraShort 20+ Year Treasury (장기채 -2X)', '-2X'),
    ('PST', 'ProShares UltraShort 7-10 Year Treasury (중기채 -2X)', '-2X'),
    ('MZZ', 'ProShares UltraShort MidCap400 (중형주 -2X)', '-2X'),
    ('BIS', 'ProShares UltraShort Nasdaq Biotechnology (바이오 -2X)', '-2X'),
    ('REW', 'ProShares UltraShort Technology (테크 -2X)', '-2X'),
    ('FXP', 'ProShares UltraShort FTSE China 50 (중국 -2X)', '-2X'),
    ('DUST', 'Direxion Daily Gold Miners Bear 2X (금광 -2X)', '-2X'),
    ('JDST', 'Direxion Daily Jr Gold Miners Bear 2X (주니어금광 -2X)', '-2X'),

    # --- -1X (Inverse) ---
    ('PSQ', 'ProShares Short QQQ (나스닥100 -1X)', '-1X'),
    ('SH', 'ProShares Short S&P500 (S&P500 -1X)', '-1X'),
    ('DOG', 'ProShares Short Dow30 (다우30 -1X)', '-1X'),
    ('RWM', 'ProShares Short Russell2000 (러셀2000 -1X)', '-1X'),
    ('SEF', 'ProShares Short Financials (금융 -1X)', '-1X'),
    ('EFZ', 'ProShares Short MSCI EAFE (선진국 -1X)', '-1X'),
    ('MYY', 'ProShares Short MidCap400 (중형주 -1X)', '-1X'),
    ('SJB', 'ProShares Short High Yield (하이일드 -1X)', '-1X'),
    ('TBX', 'ProShares Short 7-10 Year Treasury (중기채 -1X)', '-1X'),
    ('HDGE', 'AdvisorShares Ranger Equity Bear ETF (액티브 숏)', '-1X'),

    # --- 1X (Core / Benchmark / Sector / Dividend / Commodity) ---
    ('SPY', 'SPDR S&P 500 ETF Trust', '1X'),
    ('VOO', 'Vanguard S&P 500 ETF', '1X'),
    ('IVV', 'iShares Core S&P 500 ETF', '1X'),
    ('QQQ', 'Invesco QQQ Trust (나스닥100)', '1X'),
    ('DIA', 'SPDR Dow Jones Industrial Average ETF (다우존스)', '1X'),
    ('IWM', 'iShares Russell 2000 ETF (소형주)', '1X'),
    ('VTI', 'Vanguard Total Stock Market ETF (미국 전체 주식)', '1X'),
    ('SCHD', 'Schwab U.S. Dividend Equity ETF (미국 배당성장)', '1X'),
    ('JEPI', 'JPMorgan Equity Premium Income ETF (월배당 커버드콜)', '1X'),
    ('JEPQ', 'JPMorgan Nasdaq Equity Premium Income ETF (나스닥 커버드콜)', '1X'),
    ('TLT', 'iShares 20+ Year Treasury Bond ETF (미국 20년+ 국채)', '1X'),
    ('IEF', 'iShares 7-10 Year Treasury Bond ETF (미국 7-10년 국채)', '1X'),
    ('SHY', 'iShares 1-3 Year Treasury Bond ETF (미국 1-3년 단기국채)', '1X'),
    ('BIL', 'SPDR Bloomberg 1-3 Month T-Bill ETF (초단기 국채)', '1X'),
    ('SGOV', 'iShares 0-3 Month Treasury Bond ETF (0-3개월 단기국채)', '1X'),
    ('GLD', 'SPDR Gold Shares (금 현물)', '1X'),
    ('SLV', 'iShares Silver Trust (은 현물)', '1X'),
    ('USO', 'United States Oil Fund (원유)', '1X'),
    ('UNG', 'United States Natural Gas Fund (천연가스)', '1X'),
    ('BND', 'Vanguard Total Bond Market ETF (미국 종합 채권)', '1X'),
    ('AGG', 'iShares Core U.S. Aggregate Bond ETF (미국 종합 채권)', '1X'),
    ('SMH', 'VanEck Semiconductor ETF (반도체 대표)', '1X'),
    ('SOXX', 'iShares Semiconductor ETF (반도체 지수)', '1X'),
    ('XLK', 'Technology Select Sector SPDR (테크 섹터)', '1X'),
    ('XLE', 'Energy Select Sector SPDR (에너지 섹터)', '1X'),
    ('XLF', 'Financial Select Sector SPDR (금융 섹터)', '1X'),
    ('XLV', 'Health Care Select Sector SPDR (헬스케어 섹터)', '1X'),
    ('XLI', 'Industrial Select Sector SPDR (산업재 섹터)', '1X'),
    ('XLY', 'Consumer Discretionary Select Sector SPDR (임의소비재)', '1X'),
    ('XLP', 'Consumer Staples Select Sector SPDR (필수소비재)', '1X'),
    ('XLU', 'Utilities Select Sector SPDR (유틸리티 섹터)', '1X'),
    ('XLB', 'Materials Select Sector SPDR (소재 섹터)', '1X'),
    ('XLRE', 'Real Estate Select Sector SPDR (부동산 섹터)', '1X'),
    ('VNQ', 'Vanguard Real Estate ETF (리츠/부동산)', '1X'),
    ('VT', 'Vanguard Total World Stock ETF (전세계 주식)', '1X'),
    ('VXUS', 'Vanguard Total International Stock ETF (미국 제외 전세계)', '1X'),
    ('EEM', 'iShares MSCI Emerging Markets ETF (신흥국 주식)', '1X'),
    ('IEMG', 'iShares Core MSCI Emerging Markets ETF (신흥국 코어)', '1X'),
    ('EFA', 'iShares MSCI EAFE ETF (선진국 주식)', '1X'),
    ('ARKK', 'ARK Innovation ETF (혁신성장 테크)', '1X'),
    ('VIG', 'Vanguard Dividend Appreciation ETF (배당성장)', '1X'),
    ('VYM', 'Vanguard High Dividend Yield ETF (고배당)', '1X'),
    ('IWF', 'iShares Russell 1000 Growth ETF (대형성장주)', '1X'),
    ('IWD', 'iShares Russell 1000 Value ETF (대형가치주)', '1X'),
    ('QUAL', 'iShares MSCI USA Quality Factor ETF (우량 퀄리티)', '1X'),
    ('COWZ', 'Pacer US Cash Cows 100 ETF (잉여현금흐름)', '1X'),
    ('RSP', 'Invesco S&P 500 Equal Weight ETF (동일가중 S&P500)', '1X'),
    ('HYG', 'iShares iBoxx $ High Yield Corporate Bond (하이일드 채권)', '1X'),
    ('LQD', 'iShares iBoxx $ Investment Grade Corporate Bond (투자등급 회사채)', '1X'),
    ('EMB', 'iShares J.P. Morgan USD Emerging Markets Bond (신흥국 국채)', '1X'),
]


def build_us_market_data(target_date: str = None):
    """
    yfinance 일괄 배치 다운로드를 통해 US Market 주요 ETF의 가격, 거래대금, 7대 기간 수익률 수집
    """
    if not target_date:
        target_date = get_latest_business_date()

    tickers = [item[0] for item in US_ETF_UNIVERSE]
    meta_map = {item[0]: {'name': item[1], 'leverage': item[2]} for item in US_ETF_UNIVERSE}

    print(f"[US Market] 총 {len(tickers)}개 US ETF 시계열 일괄 배치 다운로드 시작 (기준일: {target_date})...")
    t0 = time.time()
    try:
        data = yf.download(tickers, period="3y", interval="1d", progress=False, group_by='column')
    except Exception as e:
        print(f"[US Market] yfinance 다운로드 에러: {e}")
        return pd.DataFrame()
    t1 = time.time()
    print(f"[US Market] {len(tickers)}개 ETF 다운로드 완료 ({t1 - t0:.2f}초 소요).")

    records = []
    close_df = data.get('Close') if isinstance(data, pd.DataFrame) else None
    volume_df = data.get('Volume') if isinstance(data, pd.DataFrame) else None

    if close_df is None or close_df.empty:
        print("[US Market] Close 데이터를 파싱할 수 없습니다.")
        return pd.DataFrame()

    # 미국 시장 거래시간 여부 판별 (미국 동부 EDT 기준 평일 09:30 ~ 16:00 정규장)
    now_utc = datetime.datetime.now(datetime.timezone.utc)
    now_edt = now_utc - datetime.timedelta(hours=4) # EDT: UTC-4
    is_us_trading_hours = (now_edt.weekday() < 5) and (
        (now_edt.hour == 9 and now_edt.minute >= 30) or
        (10 <= now_edt.hour < 16)
    )
    today_edt_str = now_edt.strftime('%Y-%m-%d')

    for ticker in tickers:
        if ticker not in close_df.columns:
            continue

        c_series = close_df[ticker].dropna()
        if c_series.empty or len(c_series) < 2:
            continue

        v_series = volume_df[ticker].dropna() if volume_df is not None and ticker in volume_df.columns else None

        # 만약 미국 정규장 진행 중(is_us_trading_hours)이고 데이터에 오늘자 실시간 미마감 봉이 있다면 전일까지로 슬라이싱
        if is_us_trading_hours and c_series.index[-1].strftime('%Y-%m-%d') >= today_edt_str:
            c_series = c_series.iloc[:-1]
            if v_series is not None and not v_series.empty and len(v_series) > len(c_series):
                v_series = v_series.iloc[:-1]

        if c_series.empty or len(c_series) < 2:
            continue

        curr_price = round(float(c_series.iloc[-1]), 2)
        curr_volume = int(v_series.iloc[-1]) if v_series is not None and not v_series.empty else 0
        trade_val_usd = round(curr_price * curr_volume, 2)

        rets = compute_period_returns(c_series)
        name_info = meta_map.get(ticker, {})

        records.append({
            '코드/티커': ticker,
            '종목명': name_info.get('name', ticker),
            '시장': 'US Market',
            '배율': name_info.get('leverage', '1X'),
            '현재가': curr_price,
            '거래량': curr_volume,
            '거래대금': trade_val_usd,
            '1W(%)': rets['1W(%)'],
            '2W(%)': rets['2W(%)'],
            '1M(%)': rets['1M(%)'],
            '3M(%)': rets['3M(%)'],
            '6M(%)': rets['6M(%)'],
            '1Y(%)': rets['1Y(%)'],
            '3Y(%)': rets['3Y(%)'],
        })

    df_us = pd.DataFrame(records)
    # 거래대금 기준 내림차순 정렬 후 순위 부여
    df_us = df_us.sort_values(by='거래대금', ascending=False).reset_index(drop=True)
    df_us['순위'] = range(1, len(df_us) + 1)

    print(f"[US Market] 최종 {len(df_us)}개 ETF 정제 완료.")
    return df_us


# ==========================================
# 4. 메인 실행 및 파일 저장
# ==========================================
def main():
    print("=" * 60)
    print("  한국 및 미국 증시 ETF 마스터 데이터 수집 엔진 시작")
    print("=" * 60)

    target_date = get_latest_business_date()

    # 1. K Market 데이터 구축
    df_kr = build_kr_market_data(target_date=target_date)

    # 2. US Market 데이터 구축
    df_us = build_us_market_data(target_date=target_date)

    # 3. 데이터 통합
    if df_kr.empty and df_us.empty:
        print("❌ 수집된 데이터가 없습니다.")
        return False

    df_master = pd.concat([df_kr, df_us], ignore_index=True)

    # 컬럼 순서 고정 (요구사항 15개 항목 일치)
    final_cols = [
        '순위', '코드/티커', '종목명', '시장', '배율', '현재가', '거래량', '거래대금',
        '1W(%)', '2W(%)', '1M(%)', '3M(%)', '6M(%)', '1Y(%)', '3Y(%)'
    ]
    df_master = df_master[final_cols].copy()

    # 4. 마스터 파일, 캐시 파일 및 메타데이터 저장
    today_clean = target_date.replace('-', '')
    cache_path = os.path.join(CACHE_DIR, f"etf_summary_{today_clean}.csv")
    meta_path = os.path.join(CURRENT_DIR, "etf_master_meta.json")

    df_master.to_csv(MASTER_FILE, index=False, encoding='utf-8-sig')
    df_master.to_csv(cache_path, index=False, encoding='utf-8-sig')

    meta_info = {
        "target_date": target_date,
        "updated_at": datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        "kr_count": len(df_kr),
        "us_count": len(df_us),
        "total_count": len(df_master)
    }
    with open(meta_path, 'w', encoding='utf-8') as f:
        json.dump(meta_info, f, ensure_ascii=False, indent=2)

    print(f"\n✅ 마스터 데이터셋 생성 완료:")
    print(f"   - 총 종목 수: {len(df_master)}개 (한국: {len(df_kr)}개, 미국: {len(df_us)}개)")
    print(f"   - 마스터 파일: {MASTER_FILE}")
    print(f"   - 오늘자 캐시: {cache_path}")
    print(f"   - 메타 파일: {meta_path}")
    print(f"   - 기준일: {target_date}")
    return True


if __name__ == "__main__":
    main()
