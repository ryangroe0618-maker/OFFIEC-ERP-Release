# -*- coding: utf-8 -*-

from io import StringIO
import time

import pandas as pd
import requests


SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT2wyxHsCwYuJfNAsHv7V9RN_kQkhcYxWC4VQPT2ERGs0MtBx__y5dD3G9kw_rMgRweQEC0GaS4m4v2/pub?gid=0&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbytRA2OXU2jNt0llG8w2y0voM2q6HV8Xm1nCstfBQ4b3qMMSjBWLnDobr8aalYErcXdOw/exec"
SPREADSHEET_ID = "1cmY9wL9zP86mdP_BiwIqbjigHFdPemPq0sV8ulL_XMs"
TARGET_SHEET_NAME = "플랫폼"
SUMMARY_SHEET_NAME = "플랫폼 마감"
BRAND_SUMMARY_SHEET_NAME = "브랜드 마감"
APPEND_MODE = True

SOURCE_COLUMNS = ["날짜", "플랫폼", "브랜드", "코드", "품번", "사이즈", "수량", "총 판매가", "공급가", "마진", "매장명"]
OUTPUT_COLUMNS = [
    "날짜",
    "플랫폼",
    "주문번호",
    "운송장번호",
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
    "내역",
]
SUMMARY_COLUMNS = ["날짜", "플랫폼", "주문건수", "주문 수량", "판매가", "수수료", "총 KRW", "총 공급가", "마진", "마진율"]
BRAND_SUMMARY_COLUMNS = ["날짜", "플랫폼", "브랜드", "판매 수량", "총 판매가", "총 KRW", "총 공급가", "마진", "마진율"]

DOWNLOAD_RETRIES = 3
DOWNLOAD_TIMEOUT = (10, 60)
DOWNLOAD_SLEEP_SEC = 2

UPLOAD_RETRIES = 3
UPLOAD_TIMEOUT = (10, 300)
UPLOAD_SLEEP_SEC = 3


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[BRAND 마진] {message}", flush=True)


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
            log(f"CSV 다운로드 실패 ({attempt}/{DOWNLOAD_RETRIES}) -> {exc}")
            if attempt < DOWNLOAD_RETRIES:
                time.sleep(DOWNLOAD_SLEEP_SEC)
    raise last_error


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace('"', "").replace("\t", "").replace("\r", "").replace("\n", "")
    return " ".join(text.split())


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


def make_status_by_sales(value) -> str:
    sales = pd.to_numeric(str(value).replace(",", "").replace('"', "").strip(), errors="coerce")
    if pd.notna(sales) and sales < 0:
        return "반품"
    return "출고"


def read_source_sheet(csv_text: str) -> pd.DataFrame:
    df = pd.read_csv(
        StringIO(csv_text),
        dtype=str,
        header=0,
        keep_default_na=False,
        skip_blank_lines=False,
    ).fillna("")

    missing_columns = [col for col in SOURCE_COLUMNS if col not in df.columns]
    if missing_columns:
        raise DataValidationError(f"brand 취합 시트에 필수 열이 없습니다: {missing_columns}")

    df = df[SOURCE_COLUMNS].copy()
    for col in SOURCE_COLUMNS:
        df[col] = df[col].map(clean_text)

    df = df.loc[~df.apply(lambda row: "".join(row.astype(str)).strip() == "", axis=1)].reset_index(drop=True)
    if df.empty:
        raise DataValidationError("brand 취합 시트에 업로드할 데이터가 없습니다.")

    return df


def build_output_df(source_df: pd.DataFrame) -> pd.DataFrame:
    output_df = pd.DataFrame("", index=source_df.index, columns=OUTPUT_COLUMNS)
    output_df["날짜"] = source_df["날짜"]
    output_df["플랫폼"] = source_df["플랫폼"]
    output_df["브랜드"] = source_df["브랜드"]
    output_df["코드"] = source_df["코드"]
    output_df["품번"] = source_df["품번"]
    output_df["사이즈"] = source_df["사이즈"]
    output_df["수량"] = source_df["수량"]
    output_df["매장명"] = source_df["매장명"]
    output_df["총 판매가"] = source_df["총 판매가"]
    output_df["총 수입"] = source_df["총 판매가"]
    output_df["총 KRW"] = source_df["총 판매가"]
    output_df["총 공급가"] = source_df["공급가"]
    output_df["마진"] = source_df["마진"]
    output_df["내역"] = source_df["총 판매가"].map(make_status_by_sales)
    return output_df


def build_platform_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = ["날짜", "플랫폼", "수량", "총 판매가", "총 수수료", "총 KRW", "총 공급가", "마진", "내역"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise DataValidationError(f"플랫폼 마감 집계용 필수 열이 없습니다: {missing_columns}")

    summary_base_df = df[df["내역"].map(clean_text).eq("출고")].copy()
    if summary_base_df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    summary_base_df["수량_num"] = to_number_series(summary_base_df["수량"])
    summary_base_df["총_판매가"] = to_number_series(summary_base_df["총 판매가"])
    summary_base_df["총_수수료"] = to_number_series(summary_base_df["총 수수료"])
    summary_base_df["총_KRW"] = to_number_series(summary_base_df["총 KRW"])
    summary_base_df["총_공급가"] = to_number_series(summary_base_df["총 공급가"])
    summary_base_df["마진_num"] = to_number_series(summary_base_df["마진"])

    summary_df = (
        summary_base_df.groupby(["날짜", "플랫폼"], dropna=False, sort=False)
        .agg(
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
    summary_df["주문건수"] = summary_df["주문 수량"]

    summary_df["마진율"] = summary_df["마진"].div(summary_df["총 KRW"].replace(0, pd.NA)).fillna(0)

    for col in ["주문건수", "주문 수량", "판매가", "수수료", "총 KRW", "총 공급가", "마진"]:
        summary_df[col] = summary_df[col].round(0).astype(int)
    summary_df["마진율"] = summary_df["마진율"].map(format_percent)

    return summary_df[SUMMARY_COLUMNS]


def build_brand_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = ["날짜", "플랫폼", "브랜드", "수량", "총 판매가", "총 KRW", "총 공급가", "마진", "내역"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise DataValidationError(f"브랜드 마감 집계용 필수 열이 없습니다: {missing_columns}")

    summary_base_df = df[df["내역"].map(clean_text).eq("출고")].copy()
    if summary_base_df.empty:
        return pd.DataFrame(columns=BRAND_SUMMARY_COLUMNS)

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
        .rename(
            columns={
                "판매_수량": "판매 수량",
                "총_판매가": "총 판매가",
                "총_KRW": "총 KRW",
                "총_공급가": "총 공급가",
            }
        )
    )

    summary_df["마진율"] = summary_df["마진"].div(summary_df["총 KRW"].replace(0, pd.NA)).fillna(0)

    for col in ["판매 수량", "총 판매가", "총 KRW", "총 공급가", "마진"]:
        summary_df[col] = summary_df[col].round(0).astype(int)
    summary_df["마진율"] = summary_df["마진율"].map(format_percent)

    return summary_df[BRAND_SUMMARY_COLUMNS]


def upload_to_google_sheet(df: pd.DataFrame, sheet_name: str):
    if df.empty:
        raise DataValidationError("업로드할 데이터가 없습니다.")

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
    log("brand 취합 시트 다운로드 시작")

    csv_text = fetch_csv_text(make_session(), SOURCE_CSV_URL)
    source_df = read_source_sheet(csv_text)
    output_df = build_output_df(source_df)
    summary_df = build_platform_summary_df(output_df)
    brand_summary_df = build_brand_summary_df(output_df)

    log(f"정리 완료 -> {output_df.shape[0]}행, {output_df.shape[1]}열")
    upload_to_google_sheet(output_df, TARGET_SHEET_NAME)

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
