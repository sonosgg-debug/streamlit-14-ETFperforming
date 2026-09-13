"""
data_loader.py
한국(K Market) 및 미국(US Market) 증시 ETF 데이터의 하이브리드 캐시 로딩,
시계열 주가 및 벤치마크/MDD 분석 지표 산출, 고품질 서식 적용 엑셀 다운로드를 제공하는 모듈.
"""

import os
import io
import datetime
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import yfinance as yf
try:
    import FinanceDataReader as fdr
except Exception:
    fdr = None

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(CURRENT_DIR, "cache")
os.makedirs(CACHE_DIR, exist_ok=True)
MASTER_FILE = os.path.join(CURRENT_DIR, "etf_master_data.csv")


def get_latest_business_date():
    """가장 최근 영업일 YYYY-MM-DD 반환"""
    now_utc = datetime.datetime.now(datetime.timezone.utc)
    now_kst = now_utc + datetime.timedelta(hours=9)
    start_offset = 0 if (now_kst.weekday() < 5 and now_kst.hour >= 16) else 1

    for i in range(start_offset, start_offset + 10):
        d = now_kst - datetime.timedelta(days=i)
        if d.weekday() < 5:
            return d.strftime('%Y-%m-%d')
    return (now_kst - datetime.timedelta(days=1)).strftime('%Y-%m-%d')


def is_cache_available(target_date: str = None) -> bool:
    """지정된 영업일(기본: 최신 영업일)의 캐시 파일이 존재하는지 확인"""
    if not target_date:
        target_date = get_latest_business_date()
    today_clean = target_date.replace('-', '')
    cache_path = os.path.join(CACHE_DIR, f"etf_summary_{today_clean}.csv")
    return os.path.exists(cache_path) and os.path.getsize(cache_path) > 10000


def cleanup_old_caches(keep_count: int = 5):
    """디스크 공간 절약을 위해 최근 keep_count개 이외의 오래된 캐시 파일 자동 삭제"""
    try:
        if not os.path.exists(CACHE_DIR):
            return
        files = [f for f in os.listdir(CACHE_DIR) if f.startswith("etf_summary_") and f.endswith(".csv")]
        files.sort(reverse=True) # 최신순 정렬
        for old_f in files[keep_count:]:
            old_path = os.path.join(CACHE_DIR, old_f)
            try:
                os.remove(old_path)
            except Exception:
                pass
    except Exception as e:
        print(f"[data_loader] 구형 캐시 정리 실패: {e}")


def get_fallback_data_date() -> str:
    """캐시 또는 마스터 파일로부터 실제 데이터의 기준 날짜를 추정"""
    try:
        if os.path.exists(CACHE_DIR):
            files = [f for f in os.listdir(CACHE_DIR) if f.startswith("etf_summary_") and f.endswith(".csv")]
            files.sort(reverse=True)
            if files:
                # etf_summary_YYYYMMDD.csv 에서 날짜 파싱
                dt_str = files[0].replace("etf_summary_", "").replace(".csv", "")
                if len(dt_str) == 8:
                    return f"{dt_str[:4]}-{dt_str[4:6]}-{dt_str[6:8]}"
        if os.path.exists(MASTER_FILE):
            mtime = os.path.getmtime(MASTER_FILE)
            return datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d')
    except Exception:
        pass
    return "이전"


def load_etf_data(force_refresh=False):
    """
    K Market 및 US Market 전체 ETF 데이터를 로드합니다.
    - 최신 영업일의 캐시가 존재하면 0.1초 즉시 반환.
    - 최신 영업일 캐시가 없거나 force_refresh=True일 경우:
      자동으로 build_master_data 모듈을 호출하여 최신 종가 및 수익률을 수집/구축 후 반환.
    - 외부 통신 장애로 최신 수집 실패 시에만 기존 마스터 파일을 Fallback으로 로드.
    반환값: (df_all, target_date_str, is_fallback)
    """
    target_date = get_latest_business_date()
    today_clean = target_date.replace('-', '')
    cache_path = os.path.join(CACHE_DIR, f"etf_summary_{today_clean}.csv")

    # 1. 강제 갱신이 아니고 당일 최신 캐시 파일이 이미 존재하는 경우 -> 초고속 반환
    if not force_refresh and os.path.exists(cache_path):
        try:
            df = pd.read_csv(cache_path, dtype={'코드/티커': str}, encoding='utf-8-sig')
            if not df.empty and len(df) >= 100:
                return df, target_date, False
        except Exception as e:
            print(f"[data_loader] 당일 캐시 로드 오류: {e}")

    # 2. 당일 캐시가 없거나 강제 갱신 요청 시 -> 최신 데이터 자동 수집 및 캐싱
    print(f"[data_loader] 최신 영업일({target_date}) 데이터 구축 엔진 가동 (force_refresh={force_refresh})...")
    build_success = False
    try:
        import build_master_data
        build_success = build_master_data.main()
    except Exception as e:
        print(f"[data_loader] 최신 데이터 자동 수집 중 예외 발생: {e}")

    # 수집 완료 후 생성된 최신 캐시 로드
    if os.path.exists(cache_path):
        try:
            df = pd.read_csv(cache_path, dtype={'코드/티커': str}, encoding='utf-8-sig')
            if not df.empty and len(df) >= 100:
                cleanup_old_caches(keep_count=5)
                return df, target_date, False
        except Exception as e:
            print(f"[data_loader] 새로 생성된 캐시 로드 실패: {e}")

    # 3. 비상 대비 Fallback: 외부 API 차단/네트워크 단절 등으로 당일 수집 실패 시
    #    과거 마스터 파일을 읽되, is_fallback=True 플래그와 실제 파일 날짜를 반환하여
    #    화면상에 경고 배너를 명시적으로 노출할 수 있도록 함.
    print("[data_loader] ⚠️ 최신 데이터 수집 실패로 기존 마스터 파일(Fallback) 로드를 시도합니다.")
    fallback_date = get_fallback_data_date()
    if os.path.exists(MASTER_FILE):
        try:
            df = pd.read_csv(MASTER_FILE, dtype={'코드/티커': str}, encoding='utf-8-sig')
            if not df.empty and len(df) >= 100:
                return df, fallback_date, True
        except Exception as e:
            print(f"[data_loader] Fallback 마스터 파일 로드 실패: {e}")

    return pd.DataFrame(), target_date, True


def load_etf_history(ticker: str, market: str, months: int = 12):
    """
    선택된 ETF의 시계열 주가(종가, MA20, MA60, MA120, 거래량)와
    벤치마크 지수(K Market: KODEX 200, US Market: SPY)의 비교 수익률 및 MDD 계산.
    반환값: (df_history, df_benchmark, mdd_percent, summary_stats)
    """
    days = int(months * 30.5 + 40)
    start_date = (datetime.datetime.now() - datetime.timedelta(days=days)).strftime('%Y-%m-%d')
    bm_ticker = "069500" if market == "K Market" else "SPY"
    bm_name = "코스피 200 (KODEX 200)" if market == "K Market" else "S&P 500 (SPY)"

    df_hist = pd.DataFrame()
    df_bm = pd.DataFrame()

    # 1. 대상 ETF 주가 수집
    try:
        if market == "K Market":
            if fdr is not None:
                df_hist = fdr.DataReader(ticker, start_date)
            if df_hist is None or df_hist.empty:
                # yfinance fallback
                t_ks = f"{ticker}.KS"
                df_hist = yf.download(t_ks, start=start_date, progress=False)
        else:
            df_hist = yf.download(ticker, start=start_date, progress=False)
            if isinstance(df_hist.columns, pd.MultiIndex):
                df_hist = df_hist.xs(ticker, axis=1, level=1) if ticker in df_hist.columns.levels[1] else df_hist
    except Exception as e:
        print(f"[data_loader] {ticker} 주가 로드 실패: {e}")

    # 2. 벤치마크 주가 수집
    try:
        if market == "K Market":
            if fdr is not None:
                df_bm = fdr.DataReader(bm_ticker, start_date)
            if df_bm is None or df_bm.empty:
                df_bm = yf.download(f"{bm_ticker}.KS", start=start_date, progress=False)
        else:
            df_bm = yf.download(bm_ticker, start=start_date, progress=False)
            if isinstance(df_bm.columns, pd.MultiIndex):
                df_bm = df_bm.xs(bm_ticker, axis=1, level=1) if bm_ticker in df_bm.columns.levels[1] else df_bm
    except Exception as e:
        print(f"[data_loader] 벤치마크 {bm_ticker} 주가 로드 실패: {e}")

    # 3. 데이터 가공 및 기술적 지표 계산
    summary_stats = {
        'high_52w': np.nan,
        'low_52w': np.nan,
        'high_diff': np.nan,
        'low_diff': np.nan,
        'mdd': 0.0,
        'current_price': 0.0,
        'bm_name': bm_name,
        'cum_return': 0.0,
        'bm_cum_return': 0.0,
    }

    if df_hist is not None and not df_hist.empty and len(df_hist) >= 2:
        # 컬럼 표준화
        col_map = {c: c.capitalize() for c in df_hist.columns}
        df_hist = df_hist.rename(columns=col_map)
        if 'Close' in df_hist.columns:
            df_hist['종가'] = df_hist['Close'].astype(float)
        elif 'close' in df_hist.columns:
            df_hist['종가'] = df_hist['close'].astype(float)

        if 'Volume' in df_hist.columns:
            df_hist['거래량'] = df_hist['Volume'].astype(float)
        elif 'volume' in df_hist.columns:
            df_hist['거래량'] = df_hist['volume'].astype(float)

        # 이동평균선
        df_hist['MA20'] = df_hist['종가'].rolling(window=20).mean()
        df_hist['MA60'] = df_hist['종가'].rolling(window=60).mean()
        df_hist['MA120'] = df_hist['종가'].rolling(window=120).mean()

        # 누적 수익률 (%) - 소수점 2자리 반올림
        p0 = float(df_hist['종가'].iloc[0])
        df_hist['누적수익률'] = (((df_hist['종가'] / p0) - 1.0) * 100.0).round(2)

        # 고점 대비 낙폭(Drawdown) 및 MDD - 소수점 2자리 반올림
        df_hist['고점'] = df_hist['종가'].cummax()
        df_hist['낙폭(Drawdown)'] = (((df_hist['종가'] - df_hist['고점']) / df_hist['고점']) * 100.0).round(2)
        mdd_val = round(float(df_hist['낙폭(Drawdown)'].min()), 2)
        summary_stats['mdd'] = mdd_val

        # 52주(최근 252거래일) 고점/저점
        hist_1y = df_hist.tail(252)
        h52 = float(hist_1y['종가'].max())
        l52 = float(hist_1y['종가'].min())
        curr_p = float(df_hist['종가'].iloc[-1])
        summary_stats['high_52w'] = h52
        summary_stats['low_52w'] = l52
        summary_stats['high_diff'] = round(((curr_p / h52) - 1.0) * 100.0, 2) if h52 > 0 else 0.0
        summary_stats['low_diff'] = round(((curr_p / l52) - 1.0) * 100.0, 2) if l52 > 0 else 0.0
        summary_stats['current_price'] = curr_p
        summary_stats['cum_return'] = round(float(df_hist['누적수익률'].iloc[-1]), 2)

    if df_bm is not None and not df_bm.empty and len(df_bm) >= 2:
        col_map = {c: c.capitalize() for c in df_bm.columns}
        df_bm = df_bm.rename(columns=col_map)
        bm_close_col = 'Close' if 'Close' in df_bm.columns else df_bm.columns[0]
        bm_p0 = float(df_bm[bm_close_col].iloc[0])
        df_bm['누적수익률'] = (((df_bm[bm_close_col].astype(float) / bm_p0) - 1.0) * 100.0).round(2)
        summary_stats['bm_cum_return'] = round(float(df_bm['누적수익률'].iloc[-1]), 2)

    return df_hist, df_bm, summary_stats['mdd'], summary_stats


def create_excel_download(df_input: pd.DataFrame, market: str, leverage: str = "1X"):
    """
    openpyxl 서식이 적용된 전문 엑셀 파일(.xlsx) 바이너리 스트림 생성
    - 다크 네이비 헤더 (#1E293B)
    - 통화 서식 자동 적용 (한국 원화 ₩ / 미국 달러 $)
    - 7대 기간 수익률 양수(초록/적색) / 음수(파란색) 셀 색상 강조
    - 자동 열 너비 계산
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{market}_ETF_TOP100"

    is_korean = (market == "K Market")

    # 1. 상단 타이틀 행 추가
    title_text = f"한국 및 미국 증시 ETF 수익률 비교 리포트 - [{market} | 배율: {leverage}]"
    ws.merge_cells("A1:O1")
    title_cell = ws["A1"]
    title_cell.value = title_text
    title_cell.font = Font(name="Malgun Gothic", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # 2. 메타 정보 행 추가
    meta_text = f"생성일시: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} | 기준: 전일 종가 기준 | 정렬 디폴트: 거래대금 상위순"
    ws.merge_cells("A2:O2")
    meta_cell = ws["A2"]
    meta_cell.value = meta_text
    meta_cell.font = Font(name="Malgun Gothic", size=9, italic=True, color="64748B")
    meta_cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    meta_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # 3. 빈 행
    ws.row_dimensions[3].height = 10

    # 4. 헤더 행 작성
    headers = list(df_input.columns)
    ws.row_dimensions[4].height = 28

    header_font = Font(name="Malgun Gothic", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border_side = Side(border_style="thin", color="CBD5E1")
    cell_border = Border(
        left=thin_border_side, right=thin_border_side,
        top=thin_border_side, bottom=thin_border_side
    )

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = cell_border

    # 5. 데이터 행 작성
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    regular_font = Font(name="Malgun Gothic", size=9, color="0F172A")
    bold_font = Font(name="Malgun Gothic", size=9, bold=True, color="0F172A")

    # 수익률 전용 폰트
    pos_font = Font(name="Malgun Gothic", size=9, color="DC2626") # 양수: 빨간색
    neg_font = Font(name="Malgun Gothic", size=9, color="2563EB") # 음수: 파란색

    return_cols = {'1W(%)', '2W(%)', '1M(%)', '3M(%)', '6M(%)', '1Y(%)', '3Y(%)'}

    for row_idx, row_data in enumerate(df_input.itertuples(index=False), 5):
        ws.row_dimensions[row_idx].height = 22
        current_fill = alt_fill if row_idx % 2 == 0 else white_fill

        for col_idx, (col_name, val) in enumerate(zip(headers, row_data), 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = current_fill
            cell.border = cell_border

            # 값 처리 및 서식 지정
            if pd.isna(val) or val is None:
                cell.value = "-"
                cell.font = regular_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                continue

            if col_name == "순위":
                cell.value = int(val)
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = "#,##0"

            elif col_name == "코드/티커":
                cell.value = str(val)
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            elif col_name == "종목명":
                cell.value = str(val)
                cell.font = regular_font
                cell.alignment = Alignment(horizontal="left", vertical="center")

            elif col_name in ["시장", "배율"]:
                cell.value = str(val)
                cell.font = regular_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            elif col_name == "현재가":
                cell.value = float(val)
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0원" if is_korean else "$#,##0.00"

            elif col_name == "거래량":
                cell.value = int(val)
                cell.font = regular_font
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

            elif col_name == "거래대금":
                cell.value = float(val)
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0원" if is_korean else "$#,##0"

            elif col_name in return_cols:
                num_val = float(val)
                cell.value = num_val / 100.0  # 엑셀 백분율을 위해 100으로 나눔
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "+0.00%;-0.00%;0.00%"
                if num_val > 0:
                    cell.font = pos_font
                elif num_val < 0:
                    cell.font = neg_font
                else:
                    cell.font = regular_font

            else:
                cell.value = val
                cell.font = regular_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # 6. 헤더 토글 필터 (AutoFilter) 적용 (오름차순/내림차순 및 값 필터 토글 단추)
    last_col_letter = get_column_letter(len(headers))
    last_row_num = len(df_input) + 4
    ws.auto_filter.ref = f"A4:{last_col_letter}{last_row_num}"

    # 7. 열 너비 자동 맞춤
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            # 1~3행은 타이틀이므로 제외
            if cell.row in [1, 2, 3]:
                continue
            if cell.value:
                val_str = str(cell.value)
                # 한글 문자 길이 1.8배 가중치
                korean_count = sum(1 for ch in val_str if ord(ch) > 127)
                effective_len = len(val_str) + (korean_count * 0.8)
                max_len = max(max_len, effective_len)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 11)

    # 7. 바이너리 스트림 반환
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
