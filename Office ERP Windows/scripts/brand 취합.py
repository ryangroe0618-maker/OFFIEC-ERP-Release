# -*- coding: utf-8 -*-

from datetime import datetime
from io import StringIO
import time
from zoneinfo import ZoneInfo

import pandas as pd
import requests


SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTZ-xM3cvMwsZfOXG8CAJC0oGCXJ4bXGcN8_edDCotqk_7cwrRN362ZuzdP7q_JRacgaSXfPzQodHCa/pub?gid=1395584309&single=true&output=csv"
REFERENCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=0&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzHik7Q5bQ8rxsyTTsUuX7FH4ShNbvmHDUwhKH7TcYTFI921JUpE7HxXWF_bGcbv7SoOw/exec"
SPREADSHEET_ID = "1Johoz-uvmt6L2nrZazSCSNexHNNMzgbln2Sem729jRk"
TARGET_SHEET_NAME = "BRAND"
APPEND_MODE = False

SOURCE_COLUMN_INDEXES = [1, 2, 4, 5, 11, 12]  # B, C, E, F, L, M
SOURCE_COLUMNS = ["날짜", "품번", "사이즈", "수량", "총 판매가", "매장명"]
OUTPUT_COLUMNS = ["날짜", "플랫폼", "브랜드", "코드", "품번", "사이즈", "수량", "총 판매가", "공급가", "마진", "매장명"]
PLATFORM_NAME = "브랜더"
TODAY_TZ = ZoneInfo("Asia/Seoul")

DOWNLOAD_RETRIES = 3
DOWNLOAD_TIMEOUT = (10, 60)
DOWNLOAD_SLEEP_SEC = 2

UPLOAD_RETRIES = 3
UPLOAD_TIMEOUT = (10, 300)
UPLOAD_SLEEP_SEC = 3


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[BRAND 정리] {message}", flush=True)


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


def normalize_date_text(value) -> str:
    return clean_text(value).replace(" ", "")


def make_part_no_key(value) -> str:
    return clean_text(value).replace("-", "")


def to_number_series(series: pd.Series) -> pd.Series:
    cleaned = (
        series.fillna("")
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace('"', "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def format_number(value) -> str:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return ""
    if float(number).is_integer():
        return str(int(number))
    return str(float(number))


def today_date_texts() -> set[str]:
    now = datetime.now(TODAY_TZ)
    return {
        f"{now.month}월{now.day}일",
        f"{now.month}/{now.day}",
        f"{now.month}-{now.day}",
        now.strftime("%Y-%m-%d"),
        now.strftime("%Y/%m/%d"),
        now.strftime("%Y.%m.%d"),
        now.strftime("%m/%d"),
        now.strftime("%m-%d"),
    }


def read_source_sheet(csv_text: str) -> pd.DataFrame:
    df_raw = pd.read_csv(
        StringIO(csv_text),
        dtype=str,
        header=0,
        keep_default_na=False,
        skip_blank_lines=False,
    ).fillna("")

    required_cols = max(SOURCE_COLUMN_INDEXES) + 1
    if df_raw.shape[1] < required_cols:
        raise DataValidationError(
            f"브랜더 출고 내역 열 수가 부족합니다. B,C,E,F,L,M 열이 필요한데 실제 {df_raw.shape[1]}열입니다."
        )

    df = df_raw.iloc[:, SOURCE_COLUMN_INDEXES].copy()
    df.columns = SOURCE_COLUMNS

    for col in SOURCE_COLUMNS:
        df[col] = df[col].map(clean_text)

    df = df.loc[~df.apply(lambda row: "".join(row.astype(str)).strip() == "", axis=1)].reset_index(drop=True)
    df = df.loc[df["날짜"].map(normalize_date_text).isin(today_date_texts())].reset_index(drop=True)
    if df.empty:
        raise DataValidationError("오늘 날짜 기준 업로드할 브랜더 출고 내역 데이터가 없습니다.")

    return df


def read_reference_sheet(csv_text: str) -> pd.DataFrame:
    df = pd.read_csv(
        StringIO(csv_text),
        dtype=str,
        header=0,
        keep_default_na=False,
        skip_blank_lines=False,
    ).fillna("")

    required_columns = ["브랜드", "품번", "공급가"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise DataValidationError(f"기준 시트에 필수 열이 없습니다: {missing_columns}")

    ref_df = df[required_columns].copy()
    for col in required_columns:
        ref_df[col] = ref_df[col].map(clean_text)

    ref_df = ref_df.loc[ref_df["품번"].ne("")].drop_duplicates(subset=["품번"], keep="first")
    if ref_df.empty:
        raise DataValidationError("기준 시트에 매칭할 품번 데이터가 없습니다.")

    return ref_df


def enrich_source_df(source_df: pd.DataFrame, reference_df: pd.DataFrame) -> pd.DataFrame:
    output_df = source_df.merge(reference_df, on="품번", how="left")
    missing_mask = output_df["브랜드"].isna() | output_df["브랜드"].map(clean_text).eq("")
    if missing_mask.any():
        normalized_reference_df = reference_df.copy()
        normalized_reference_df["품번_key"] = normalized_reference_df["품번"].map(make_part_no_key)
        normalized_reference_df = normalized_reference_df.loc[normalized_reference_df["품번_key"].ne("")]
        normalized_reference_df = normalized_reference_df.drop_duplicates(subset=["품번_key"], keep="first")

        normalized_source_df = output_df.loc[missing_mask, ["품번"]].copy()
        normalized_source_df["품번_key"] = normalized_source_df["품번"].map(make_part_no_key)
        fallback_df = normalized_source_df.merge(
            normalized_reference_df[["품번_key", "브랜드", "공급가"]],
            on="품번_key",
            how="left",
            suffixes=("", "_fallback"),
        )

        fallback_indexes = output_df.index[missing_mask]
        output_df.loc[fallback_indexes, "브랜드"] = fallback_df["브랜드"].values
        output_df.loc[fallback_indexes, "공급가"] = fallback_df["공급가"].values

    output_df["플랫폼"] = PLATFORM_NAME
    output_df["브랜드"] = output_df["브랜드"].fillna("").map(clean_text)
    output_df["공급가"] = output_df["공급가"].fillna("").map(clean_text)
    fila_mask = output_df["브랜드"].str.upper().eq("FILA")
    output_df.loc[fila_mask, "품번"] = output_df.loc[fila_mask, "품번"].str.replace("-", "", regex=False)
    output_df["코드"] = output_df["품번"] + output_df["사이즈"]

    qty_series = to_number_series(output_df["수량"])
    sales_series = to_number_series(output_df["총 판매가"])
    supply_total_series = to_number_series(output_df["공급가"]) * qty_series
    margin_series = sales_series - supply_total_series

    output_df["공급가"] = supply_total_series.map(format_number)
    output_df["마진"] = margin_series.map(format_number)

    missing_count = output_df["브랜드"].eq("").sum()
    if missing_count:
        log(f"기준 시트 품번 미매칭 -> {missing_count}행")

    return output_df[OUTPUT_COLUMNS]


def upload_to_google_sheet(df: pd.DataFrame):
    if df.empty:
        raise DataValidationError("업로드할 데이터가 없습니다.")

    values = [df.columns.tolist()] + df.astype(str).values.tolist()
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": values,
        "append": APPEND_MODE,
        "clear": not APPEND_MODE,
    }

    last_error = None
    mode_label = "누적" if APPEND_MODE else "덮어쓰기"
    log(f"업로드 시작 -> {TARGET_SHEET_NAME} ({len(df)}행, {df.shape[1]}열, {mode_label})")

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

            log(f"업로드 완료 -> {TARGET_SHEET_NAME}")
            return
        except (requests.exceptions.RequestException, RuntimeError) as exc:
            last_error = exc
            log(f"업로드 실패 ({attempt}/{UPLOAD_RETRIES}) -> {exc}")
            if attempt < UPLOAD_RETRIES:
                time.sleep(UPLOAD_SLEEP_SEC)

    raise last_error


def main():
    start_time = time.perf_counter()
    log("브랜더 출고 내역 다운로드 시작")

    session = make_session()
    source_csv_text = fetch_csv_text(session, SOURCE_CSV_URL)
    reference_csv_text = fetch_csv_text(session, REFERENCE_CSV_URL)
    source_df = read_source_sheet(source_csv_text)
    reference_df = read_reference_sheet(reference_csv_text)
    output_df = enrich_source_df(source_df, reference_df)

    log(f"정리 완료 -> {output_df.shape[0]}행, {output_df.shape[1]}열")
    upload_to_google_sheet(output_df)

    elapsed = time.perf_counter() - start_time
    log(f"완료 ({elapsed:.2f}초)")


if __name__ == "__main__":
    main()
