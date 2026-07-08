# -*- coding: utf-8 -*-

from io import StringIO
import time

import pandas as pd
import requests


JAPAN_SHIPMENT_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR3-K2ZY0WEoktfTZFCLZEmlPJr_Pub9CwLvZDQSaE2ySGCdXNORnu5Wn6-Y-fzHNRlnNIvw01e4tyx/pub?gid=0&single=true&output=csv"
STOCK_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=0&single=true&output=csv"
EXCHANGE_RATE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=295228098&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxEu5k-Q6SRLYXHsZbhj02hNotycHwbz2vS0JwE3fKU_DteLCoA5H094vyC32m3-yKu/exec"
SPREADSHEET_ID = "1Fs1J2R-IY5KZVDJKRWn_jr7q6xhzVAg9xNlIOrbR8ow"
TARGET_SHEET_NAME = "JAPAN"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 300)
RETRIES = 3
RETRY_SLEEP_SEC = 2
EXCHANGE_RATE_RETRIES = 12
EXCHANGE_RATE_SLEEP_SEC = 5

JAPAN_USECOLS = ["A", "D", "F", "G", "I", "K"]
JAPAN_DATE_COLUMN = "판매일"
JAPAN_SOURCE_COLUMNS = [JAPAN_DATE_COLUMN, "품번", "사이즈", "수량", "총 판매가", "플랫폼"]
STOCK_LOOKUP_COLUMNS = ["품번", "브랜드", "할인가", "공급가"]
OUTPUT_COLUMNS = [
    "날짜",
    "플랫폼",
    "주문번호",
    "뒤 4자리",
    "브랜드",
    "코드",
    "품번",
    "사이즈",
    "수량",
    "매장명",
    "할인가",
    "총 판매가",
    "총 수수료",
    "총 수입",
    "환율",
    "총 KRW",
    "총 공급가",
    "마진",
]


class JapanUploadError(ValueError):
    pass


def log(message: str):
    print(f"[일본] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(url: str) -> str:
    last_error = None
    session = make_session()
    for attempt in range(1, RETRIES + 1):
        try:
            response = session.get(url, timeout=DOWNLOAD_TIMEOUT)
            response.raise_for_status()
            response.encoding = "utf-8"
            text = response.text
            if not text.strip():
                raise JapanUploadError("CSV 응답이 비어 있습니다.")
            return text
        except Exception as exc:
            last_error = exc
            log(f"다운로드 실패 ({attempt}/{RETRIES})")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_number_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.fillna("").astype(str).str.replace(",", "", regex=False).str.replace("₩", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0)


def parse_number(value) -> float:
    text = clean_text(value).replace(",", "")
    try:
        return float(text)
    except Exception:
        raise JapanUploadError(f"숫자로 읽을 수 없는 값입니다: {value}")


def excel_col_to_index(column: str) -> int:
    result = 0
    for char in column.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1


def read_japan_shipment_df(csv_text: str) -> pd.DataFrame:
    usecols = [excel_col_to_index(column) for column in JAPAN_USECOLS]
    df = pd.read_csv(StringIO(csv_text), header=2, usecols=usecols, dtype=str).fillna("")
    if df.shape[1] != len(JAPAN_SOURCE_COLUMNS):
        raise JapanUploadError(f"일본 출고 시트 열 수가 맞지 않습니다. 기대 {len(JAPAN_SOURCE_COLUMNS)}개, 실제 {df.shape[1]}개")
    df.columns = JAPAN_SOURCE_COLUMNS
    for column in JAPAN_SOURCE_COLUMNS:
        df[column] = df[column].apply(clean_text)
    today = time.strftime("%Y-%m-%d")
    sale_dates = pd.to_datetime(df[JAPAN_DATE_COLUMN], errors="coerce").dt.strftime("%Y-%m-%d")
    df = df[sale_dates.eq(today) | df[JAPAN_DATE_COLUMN].eq(today)].copy()
    df = df[df["품번"].ne("")].copy()
    return df.drop(columns=[JAPAN_DATE_COLUMN]).reset_index(drop=True)


def read_stock_lookup_df(csv_text: str) -> pd.DataFrame:
    df = pd.read_csv(StringIO(csv_text), dtype=str).fillna("")
    missing = [column for column in STOCK_LOOKUP_COLUMNS if column not in df.columns]
    if missing:
        raise JapanUploadError(f"현재고 시트 필수 열이 없습니다: {missing}")
    lookup_df = df[STOCK_LOOKUP_COLUMNS].copy()
    for column in STOCK_LOOKUP_COLUMNS:
        lookup_df[column] = lookup_df[column].apply(clean_text)
    lookup_df = lookup_df[lookup_df["품번"].ne("")].drop_duplicates(subset=["품번"], keep="first")
    return lookup_df.reset_index(drop=True)


def read_exchange_rate(csv_text: str) -> str:
    df = pd.read_csv(StringIO(csv_text), header=None, dtype=str).fillna("")
    if df.shape[0] < 4 or df.shape[1] < 2:
        raise JapanUploadError("환율 시트 B4 셀을 읽을 수 없습니다.")
    rate = clean_text(df.iloc[3, 1])
    parse_number(rate)
    return rate


def fetch_exchange_rate() -> str:
    last_error = None
    for attempt in range(1, EXCHANGE_RATE_RETRIES + 1):
        try:
            return read_exchange_rate(fetch_csv_text(EXCHANGE_RATE_URL))
        except Exception as exc:
            last_error = exc
            log(f"환율 읽기 실패 ({attempt}/{EXCHANGE_RATE_RETRIES}): {exc}")
            if attempt < EXCHANGE_RATE_RETRIES:
                time.sleep(EXCHANGE_RATE_SLEEP_SEC)
    raise last_error


def build_output_df() -> pd.DataFrame:
    log("일본 출고 시트 다운로드")
    japan_df = read_japan_shipment_df(fetch_csv_text(JAPAN_SHIPMENT_URL))

    log("현재고 시트 다운로드")
    stock_lookup_df = read_stock_lookup_df(fetch_csv_text(STOCK_URL))

    log("환율 시트 다운로드")
    exchange_rate = fetch_exchange_rate()

    output_df = japan_df.merge(stock_lookup_df, on="품번", how="left")
    for column in ["브랜드", "할인가", "공급가"]:
        output_df[column] = output_df[column].fillna("")

    qty = to_number_series(output_df["수량"])
    sale_price = to_number_series(output_df["총 판매가"])
    exchange_rate_num = parse_number(exchange_rate)
    supply_price = to_number_series(output_df["공급가"])

    output_df["날짜"] = time.strftime("%Y-%m-%d")
    output_df["플랫폼"] = output_df["플랫폼"].replace({"BM": "BUYMA"})
    output_df["주문번호"] = ""
    output_df["뒤 4자리"] = ""
    output_df["코드"] = output_df["품번"].astype(str) + output_df["사이즈"].astype(str)
    output_df["매장명"] = ""
    output_df["수량"] = qty.astype(int)
    output_df["총 판매가"] = sale_price.round(0).astype(int)
    output_df["환율"] = exchange_rate
    output_df["총 수수료"] = (sale_price * 0.077 + sale_price * 0.05 + (6000 / exchange_rate_num)).round(0).astype(int)
    output_df["총 수입"] = (output_df["총 판매가"] - output_df["총 수수료"]).round(0).astype(int)
    output_df["총 KRW"] = (output_df["총 수입"] * exchange_rate_num).round(0).astype(int)
    output_df["총 공급가"] = (supply_price / 1.1 * qty).round(0).astype(int)
    output_df["마진"] = (output_df["총 KRW"] - output_df["총 공급가"]).round(0).astype(int)
    return output_df[OUTPUT_COLUMNS].copy()


def upload_to_google_sheet(df: pd.DataFrame):
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": [df.columns.tolist()] + df.fillna("").astype(str).values.tolist(),
    }

    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            log(f"업로드 시작: {TARGET_SHEET_NAME} / {len(df)}행")
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            log(f"업로드 완료: {response.text}")
            return
        except requests.exceptions.RequestException as exc:
            last_error = exc
            log(f"업로드 실패 ({attempt}/{RETRIES}): {exc}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def main():
    start_time = time.perf_counter()
    output_df = build_output_df()
    upload_to_google_sheet(output_df)
    log(f"완료 ({time.perf_counter() - start_time:.2f}초)")


if __name__ == "__main__":
    main()
