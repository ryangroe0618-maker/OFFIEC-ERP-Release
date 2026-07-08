# -*- coding: utf-8 -*-

from io import StringIO
import time

import pandas as pd
import requests


SOURCE_INPUT_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSzpo-DtuISMc_boM0XnjqnY-1hIlD2s_LMhzbvaRWdeNFBxdtO1Z0Fl94s4Dxo52wOwrBfisgDyQYt/pub?gid=2077691561&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbytRA2OXU2jNt0llG8w2y0voM2q6HV8Xm1nCstfBQ4b3qMMSjBWLnDobr8aalYErcXdOw/exec"
SPREADSHEET_ID = "1cmY9wL9zP86mdP_BiwIqbjigHFdPemPq0sV8ulL_XMs"
TARGET_SHEET_NAME = "플랫폼"
SUMMARY_SHEET_NAME = "플랫폼 마감"
BRAND_SUMMARY_SHEET_NAME = "브랜드 마감"
APPEND_MODE = True

SOURCE_RANGE = "A:G,I:S"
SUMMARY_COLUMNS = ["날짜", "플랫폼", "주문건수", "주문 수량", "판매가", "수수료", "총 KRW", "총 공급가", "마진", "마진율"]
BRAND_SUMMARY_COLUMNS = ["날짜", "플랫폼", "브랜드", "판매 수량", "총 판매가", "총 KRW", "총 공급가", "마진", "마진율"]

DOWNLOAD_RETRIES = 3
DOWNLOAD_TIMEOUT = (10, 60)
DOWNLOAD_SLEEP_SEC = 2

UPLOAD_RETRIES = 3
UPLOAD_TIMEOUT = (10, 300)
UPLOAD_SLEEP_SEC = 3
SOURCE_SELECTED_INDICES = list(range(0, 7)) + list(range(8, 19))


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[POIZON] {message}")


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
        except Exception as e:
            last_error = e
            if attempt < DOWNLOAD_RETRIES:
                time.sleep(DOWNLOAD_SLEEP_SEC)

    raise last_error


def read_source_sheet(csv_text: str) -> pd.DataFrame:
    df = pd.read_csv(StringIO(csv_text), dtype=str, keep_default_na=False).fillna("")
    df = df.iloc[:, SOURCE_SELECTED_INDICES].copy()

    if df.shape[1] == 0:
        raise DataValidationError("원본 시트에 업로드할 열이 없습니다.")

    if df.shape[1] != 18:
        raise DataValidationError(
            f"원본 시트 열 수가 부족합니다. {SOURCE_RANGE} 기준 18열이어야 하는데 실제 {df.shape[1]}열입니다."
        )

    df.columns = [str(col).strip() for col in df.columns]
    non_empty_mask = df.ne("").any(axis=1)
    df = df.loc[non_empty_mask].reset_index(drop=True)

    if df.empty:
        raise DataValidationError("원본 시트에 업로드할 데이터가 없습니다.")

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

    if output_df.shape[1] < 10:
        raise DataValidationError("J열 매장명 기준 상태값을 만들 수 없습니다. 원본 열 수를 확인해주세요.")

    store_series = output_df.iloc[:, 9].astype(str).str.strip()
    output_df["내역"] = "출고"
    output_df.loc[store_series.isin({"취소", "취소건"}), "내역"] = "취소건"
    output_df.loc[store_series.isin({"재고없음", "재고 없음"}), "내역"] = "재고 없음"
    return output_df


def build_shipped_metrics_df(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = ["날짜", "플랫폼", "브랜드", "수량", "총 판매가", "총 수수료", "환율", "총 KRW", "총 공급가", "마진", "내역"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise DataValidationError(f"마감 집계용 필수 열이 없습니다: {missing_columns}")

    shipped_df = df[df["내역"].astype(str).str.strip().eq("출고")].copy()
    if shipped_df.empty:
        return shipped_df

    shipped_df["총 판매가_num"] = to_number_series(shipped_df["총 판매가"])
    shipped_df["총 수수료_num"] = to_number_series(shipped_df["총 수수료"])
    shipped_df["환율_num"] = to_number_series(shipped_df["환율"])
    shipped_df["총 KRW_num"] = to_number_series(shipped_df["총 KRW"])
    shipped_df["총 공급가_num"] = to_number_series(shipped_df["총 공급가"])
    shipped_df["마진_num"] = to_number_series(shipped_df["마진"])
    shipped_df["수량_num"] = to_number_series(shipped_df["수량"])
    shipped_df["브랜드"] = shipped_df["브랜드"].astype(str).str.strip()
    return shipped_df


def build_platform_summary_df(shipped_df: pd.DataFrame) -> pd.DataFrame:
    if shipped_df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    shipped_df["판매가"] = shipped_df["총 판매가_num"] * shipped_df["환율_num"]
    shipped_df["수수료"] = shipped_df["총 수수료_num"] * shipped_df["환율_num"]

    summary_df = (
        shipped_df.groupby(["날짜", "플랫폼"], dropna=False, sort=False)
        .agg(
            {
                "수량_num": "sum",
                "판매가": "sum",
                "수수료": "sum",
                "총 KRW_num": "sum",
                "총 공급가_num": "sum",
                "마진_num": "sum",
            }
        )
        .reset_index()
    )

    summary_df = summary_df.rename(
        columns={
            "수량_num": "주문건수",
            "총 KRW_num": "총 KRW",
            "총 공급가_num": "총 공급가",
            "마진_num": "마진",
        }
    )

    summary_df["주문 수량"] = summary_df["주문건수"]
    summary_df["마진율"] = summary_df["마진"].div(summary_df["총 KRW"].replace(0, pd.NA)).fillna(0)

    for col in ["주문건수", "주문 수량", "판매가", "수수료", "총 KRW", "총 공급가", "마진"]:
        summary_df[col] = summary_df[col].round(0).astype(int)
    summary_df["마진율"] = summary_df["마진율"].apply(format_percent)

    return summary_df[SUMMARY_COLUMNS]


def build_brand_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    shipped_df = df
    if shipped_df.empty:
        return pd.DataFrame(columns=BRAND_SUMMARY_COLUMNS)
    shipped_df["판매가_원화"] = shipped_df["총 판매가_num"] * shipped_df["환율_num"]

    summary_df = (
        shipped_df.groupby(["날짜", "플랫폼", "브랜드"], dropna=False, sort=False)
        .agg(
            {
                "수량_num": "sum",
                "판매가_원화": "sum",
                "총 KRW_num": "sum",
                "총 공급가_num": "sum",
                "마진_num": "sum",
            }
        )
        .reset_index()
        .rename(
            columns={
                "수량_num": "판매 수량",
                "판매가_원화": "총 판매가",
                "총 KRW_num": "총 KRW",
                "총 공급가_num": "총 공급가",
                "마진_num": "마진",
            }
        )
    )

    for col in ["판매 수량", "총 판매가", "총 KRW", "총 공급가", "마진"]:
        summary_df[col] = summary_df[col].round(0).astype(int)
    summary_df["마진율"] = summary_df["마진"].div(summary_df["총 KRW"].replace(0, pd.NA)).fillna(0)
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

    mode_label = "누적" if APPEND_MODE else "덮어쓰기"
    log(f"업로드 시작 -> {sheet_name} ({len(df)}행, {df.shape[1]}열, {mode_label})")
    last_error = None

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
        except (requests.exceptions.RequestException, RuntimeError) as e:
            last_error = e
            log(f"업로드 실패 ({attempt}/{UPLOAD_RETRIES}) -> {e}")
            if attempt < UPLOAD_RETRIES:
                time.sleep(UPLOAD_SLEEP_SEC)

    raise last_error


def main():
    start_time = time.perf_counter()
    log("원본 구글 시트 다운로드 시작")

    csv_text = fetch_csv_text(make_session(), SOURCE_INPUT_URL)
    source_df = read_source_sheet(csv_text)
    source_df = add_status_column(source_df)
    shipped_metrics_df = build_shipped_metrics_df(source_df)
    summary_df = build_platform_summary_df(shipped_metrics_df)
    brand_summary_df = build_brand_summary_df(shipped_metrics_df)
    log(f"원본 시트 읽기 완료 -> {source_df.shape[0]}행, {source_df.shape[1]}열")

    upload_to_google_sheet(source_df, TARGET_SHEET_NAME)

    if summary_df.empty:
        log("플랫폼 마감 업로드 대상 없음 -> 내역이 '출고'인 행이 없습니다.")
    else:
        log(f"플랫폼 마감 집계 완료 -> {summary_df.shape[0]}행")
        upload_to_google_sheet(summary_df, SUMMARY_SHEET_NAME)

    if brand_summary_df.empty:
        log("브랜드 마감 업로드 대상 없음 -> 내역이 '출고'인 행이 없습니다.")
    else:
        log(f"브랜드 마감 집계 완료 -> {brand_summary_df.shape[0]}행")
        upload_to_google_sheet(brand_summary_df, BRAND_SUMMARY_SHEET_NAME)

    elapsed = time.perf_counter() - start_time
    log(f"완료 ({elapsed:.2f}초)")


if __name__ == "__main__":
    main()
