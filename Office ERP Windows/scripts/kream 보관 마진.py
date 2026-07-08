# -*- coding: utf-8 -*-

from io import StringIO
import time

import pandas as pd
import requests


SOURCE_INPUT_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTK9BjPuP782_4akOc_w3CayC7LRrb_lp4DtrQYrelclSyJjIgS0gVvORws5uSGu0su3KDgYMGN92ny/pub?gid=1243415372&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbytRA2OXU2jNt0llG8w2y0voM2q6HV8Xm1nCstfBQ4b3qMMSjBWLnDobr8aalYErcXdOw/exec"
SPREADSHEET_ID = "1cmY9wL9zP86mdP_BiwIqbjigHFdPemPq0sV8ulL_XMs"
TARGET_SHEET_NAME = "플랫폼"
SUMMARY_SHEET_NAME = "플랫폼 마감"
BRAND_SUMMARY_SHEET_NAME = "브랜드 마감"
APPEND_MODE = True

DOWNLOAD_RETRIES = 3
DOWNLOAD_TIMEOUT = (10, 60)
DOWNLOAD_SLEEP_SEC = 2

UPLOAD_RETRIES = 3
UPLOAD_TIMEOUT = (10, 300)
UPLOAD_SLEEP_SEC = 3

SOURCE_COLUMNS = ["날짜", "플랫폼", "주문번호", "운송장번호", "브랜드", "코드", "품번", "사이즈", "수량", "매장명", "할인가", "총 판매가", "총 수수료", "총 수입", "환율", "총 KRW", "총 공급가", "마진"]
SUMMARY_COLUMNS = ["날짜", "플랫폼", "주문건수", "주문 수량", "판매가", "수수료", "총 KRW", "총 공급가", "마진", "마진율"]
BRAND_SUMMARY_COLUMNS = ["날짜", "플랫폼", "브랜드", "판매 수량", "총 판매가", "총 KRW", "총 공급가", "마진", "마진율"]


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[KREAM 보관 마진] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(session: requests.Session, url: str) -> str:
    last_error = None
    for attempt in range(1, DOWNLOAD_RETRIES + 1):
        try:
            response = session.get(url, timeout=DOWNLOAD_TIMEOUT)
            response.raise_for_status()
            response.encoding = "utf-8"
            return response.text
        except Exception as exc:
            last_error = exc
            if attempt < DOWNLOAD_RETRIES:
                time.sleep(DOWNLOAD_SLEEP_SEC)
    raise last_error


def read_source_sheet(csv_text: str) -> pd.DataFrame:
    df = pd.read_csv(StringIO(csv_text), dtype=str, keep_default_na=False).fillna("")
    if df.shape[1] < 18:
        raise DataValidationError(f"정리 시트 열 수가 부족합니다. A:R 기준 18열이 필요한데 실제 {df.shape[1]}열입니다.")

    df = df.iloc[:, :18].copy()
    df.columns = SOURCE_COLUMNS
    df = df.loc[~(df.astype(str).apply(lambda row: "".join(row).strip() == "", axis=1))].reset_index(drop=True)

    if df.empty:
        raise DataValidationError("정리 시트에 업로드할 데이터가 없습니다.")

    return df


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_number_series(series: pd.Series) -> pd.Series:
    cleaned = (
        series.fillna("")
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace('"', "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def format_percent(value) -> str:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return "0%"
    return f"{float(number) * 100:.2f}%"


def add_status_column(df: pd.DataFrame) -> pd.DataFrame:
    output_df = df.copy()
    output_df["내역"] = "출고"
    return output_df


def build_platform_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    summary_base_df = df.copy()
    if summary_base_df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    summary_base_df["주문번호"] = summary_base_df["주문번호"].apply(clean_text)
    summary_base_df["수량_num"] = to_number_series(summary_base_df["수량"])
    summary_base_df["총_판매가"] = to_number_series(summary_base_df["총 판매가"])
    summary_base_df["총_수수료"] = to_number_series(summary_base_df["총 수수료"])
    summary_base_df["총_KRW"] = to_number_series(summary_base_df["총 KRW"])
    summary_base_df["총_공급가"] = to_number_series(summary_base_df["총 공급가"])
    summary_base_df["마진_num"] = to_number_series(summary_base_df["마진"])

    summary_df = (
        summary_base_df.groupby(["날짜", "플랫폼"], dropna=False, sort=False)
        .agg(
            주문건수=("주문번호", "count"),
            주문_수량=("수량_num", "sum"),
            판매가=("총_판매가", "sum"),
            수수료=("총_수수료", "sum"),
            총_KRW=("총_KRW", "sum"),
            총_공급가=("총_공급가", "sum"),
            마진=("마진_num", "sum"),
        )
        .reset_index()
        .rename(columns={"주문_수량": "주문 수량", "총_KRW": "총 KRW", "총_공급가": "총 공급가"})
    )

    summary_df["마진율"] = summary_df["마진"].div(summary_df["총 KRW"].replace(0, pd.NA)).fillna(0)
    for col in ["주문건수", "주문 수량", "판매가", "수수료", "총 KRW", "총 공급가", "마진"]:
        summary_df[col] = summary_df[col].round(0).astype(int)
    summary_df["마진율"] = summary_df["마진율"].apply(format_percent)
    return summary_df[SUMMARY_COLUMNS]


def build_brand_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    summary_base_df = df.copy()
    if summary_base_df.empty:
        return pd.DataFrame(columns=BRAND_SUMMARY_COLUMNS)

    summary_base_df["브랜드"] = summary_base_df["브랜드"].apply(clean_text)
    summary_base_df["수량_num"] = to_number_series(summary_base_df["수량"])
    summary_base_df["총_판매가"] = to_number_series(summary_base_df["총 판매가"])
    summary_base_df["총_KRW"] = to_number_series(summary_base_df["총 KRW"])
    summary_base_df["총_공급가"] = to_number_series(summary_base_df["총 공급가"])
    summary_base_df["마진_num"] = to_number_series(summary_base_df["마진"])

    summary_df = (
        summary_base_df.groupby(["날짜", "플랫폼", "브랜드"], dropna=False, sort=False)
        .agg(
            판매_수량=("수량_num", "sum"),
            총_판매가=("총_판매가", "sum"),
            총_KRW=("총_KRW", "sum"),
            총_공급가=("총_공급가", "sum"),
            마진=("마진_num", "sum"),
        )
        .reset_index()
        .rename(columns={"판매_수량": "판매 수량", "총_판매가": "총 판매가", "총_KRW": "총 KRW", "총_공급가": "총 공급가"})
    )

    summary_df["마진율"] = summary_df["마진"].div(summary_df["총 KRW"].replace(0, pd.NA)).fillna(0)
    for col in ["판매 수량", "총 판매가", "총 KRW", "총 공급가", "마진"]:
        summary_df[col] = summary_df[col].round(0).astype(int)
    summary_df["마진율"] = summary_df["마진율"].apply(format_percent)
    return summary_df[BRAND_SUMMARY_COLUMNS]


def upload_to_google_sheet(df: pd.DataFrame, sheet_name: str):
    if df.shape[1] == 0:
        raise DataValidationError("업로드할 데이터프레임의 열 개수가 0입니다.")

    values = [df.columns.tolist()] + df.astype(str).values.tolist()
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": sheet_name,
        "values": values,
        "append": APPEND_MODE,
        "clear": not APPEND_MODE,
    }

    last_error = None
    mode_label = "누적" if APPEND_MODE else "덮어쓰기"
    log(f"업로드 시작 -> {sheet_name} ({len(df)}행, {df.shape[1]}열, {mode_label})")

    for attempt in range(1, UPLOAD_RETRIES + 1):
        try:
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            response_text = (response.text or "").strip()
            try:
                response_json = response.json()
            except Exception:
                response_json = None

            if isinstance(response_json, dict) and response_json.get("ok") is False:
                raise RuntimeError(response_json.get("error") or response_json.get("message") or response_text)

            log(f"업로드 완료 -> {sheet_name}")
            return
        except (requests.exceptions.RequestException, RuntimeError) as exc:
            last_error = exc
            log(f"업로드 실패 ({attempt}/{UPLOAD_RETRIES}) -> {exc}")
            if attempt < UPLOAD_RETRIES:
                time.sleep(UPLOAD_SLEEP_SEC)

    raise last_error


def main():
    start_time = time.perf_counter()
    log("KREAM 보관 정리 시트 다운로드 시작")

    csv_text = fetch_csv_text(make_session(), SOURCE_INPUT_URL)
    source_df = read_source_sheet(csv_text)
    source_df = add_status_column(source_df)
    summary_df = build_platform_summary_df(source_df)
    brand_summary_df = build_brand_summary_df(source_df)

    log(f"정리 시트 읽기 완료 -> {source_df.shape[0]}행, {source_df.shape[1]}열")
    upload_to_google_sheet(source_df, TARGET_SHEET_NAME)
    if summary_df.empty:
        log("플랫폼 마감 업로드 대상 없음")
    else:
        log(f"플랫폼 마감 집계 완료 -> {summary_df.shape[0]}행")
        upload_to_google_sheet(summary_df, SUMMARY_SHEET_NAME)

    if brand_summary_df.empty:
        log("브랜드 마감 업로드 대상 없음")
    else:
        log(f"브랜드 마감 집계 완료 -> {brand_summary_df.shape[0]}행")
        upload_to_google_sheet(brand_summary_df, BRAND_SUMMARY_SHEET_NAME)

    elapsed = time.perf_counter() - start_time
    log(f"완료 ({elapsed:.2f}초)")


if __name__ == "__main__":
    main()
