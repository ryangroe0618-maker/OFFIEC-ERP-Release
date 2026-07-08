# -*- coding: utf-8 -*-

from io import StringIO
import json
import time

import pandas as pd
import requests


SOURCE_INPUT_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTsZm7LRlCqZRYE6BqWfqeTH_PqhxT-hTDK4ypOPPnwcTr7hpr98L6_nL19I2k_P-uk5WlOGREye13p/pub?gid=2059999598&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbytRA2OXU2jNt0llG8w2y0voM2q6HV8Xm1nCstfBQ4b3qMMSjBWLnDobr8aalYErcXdOw/exec"
SPREADSHEET_ID = "1cmY9wL9zP86mdP_BiwIqbjigHFdPemPq0sV8ulL_XMs"
TARGET_SHEET_NAME = "반품"
APPEND_MODE = True

DOWNLOAD_RETRIES = 3
DOWNLOAD_TIMEOUT = (10, 60)
DOWNLOAD_SLEEP_SEC = 2
UPLOAD_RETRIES = 3
UPLOAD_TIMEOUT = (10, 300)
UPLOAD_SLEEP_SEC = 3

SOURCE_COL_COUNT = 20  # A:T
START_TIME = time.perf_counter()


class DataValidationError(ValueError):
    pass


def log(message: str):
    elapsed = time.perf_counter() - START_TIME
    print(f"[KREAM 반품 마진 {elapsed:6.1f}s] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(session: requests.Session, url: str) -> str:
    last_error = None
    for attempt in range(1, DOWNLOAD_RETRIES + 1):
        try:
            log(f"원본 CSV 다운로드 시작 ({attempt}/{DOWNLOAD_RETRIES})")
            response = session.get(url, timeout=DOWNLOAD_TIMEOUT)
            response.raise_for_status()
            response.encoding = "utf-8"
            text = response.text
            if not text.strip():
                raise DataValidationError("원본 CSV 응답이 비어 있습니다.")
            log("원본 CSV 다운로드 완료")
            return text
        except Exception as exc:
            last_error = exc
            log(f"원본 CSV 다운로드 실패 ({attempt}/{DOWNLOAD_RETRIES}) -> {exc}")
            if attempt < DOWNLOAD_RETRIES:
                time.sleep(DOWNLOAD_SLEEP_SEC)
    raise last_error


def read_source_sheet(csv_text: str) -> pd.DataFrame:
    df_raw = pd.read_csv(StringIO(csv_text), dtype=str, keep_default_na=False).fillna("")
    if df_raw.shape[1] < SOURCE_COL_COUNT:
        raise DataValidationError(f"원본 시트 열 수가 부족합니다. A:T 20열 필요, 실제 {df_raw.shape[1]}열")

    df = df_raw.iloc[:, :SOURCE_COL_COUNT].copy()
    df.columns = [str(col).strip() for col in df.columns]
    df = df.apply(lambda col: col.astype(str).str.strip())
    df = df.loc[df.ne("").any(axis=1)].reset_index(drop=True)

    if df.empty:
        raise DataValidationError("업로드할 KREAM 반품 데이터가 없습니다.")

    return df


def upload_to_google_sheet(session: requests.Session, df: pd.DataFrame):
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": values,
        "append": APPEND_MODE,
        "clear": not APPEND_MODE,
    }

    mode_label = "누적" if APPEND_MODE else "덮어쓰기"
    log(f"업로드 시작 -> {TARGET_SHEET_NAME} ({len(df)}행, {df.shape[1]}열, {mode_label})")
    last_error = None

    for attempt in range(1, UPLOAD_RETRIES + 1):
        try:
            response = session.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            response_text = (response.text or "").strip()

            try:
                response_json = json.loads(response_text)
            except Exception:
                response_json = None

            if response_json is None and "<html" in response_text.lower():
                raise RuntimeError(f"Apps Script가 JSON 대신 HTML을 반환했습니다: {response_text[:300]}")
            if isinstance(response_json, dict) and response_json.get("ok") is False:
                raise RuntimeError(response_json.get("error") or response_json.get("message") or response_text)
            if not isinstance(response_json, dict):
                raise RuntimeError(f"Apps Script 응답을 JSON으로 확인할 수 없습니다: {response_text[:300]}")

            log(f"업로드 완료 -> {TARGET_SHEET_NAME}")
            log(f"응답: {response_text}")
            return
        except Exception as exc:
            last_error = exc
            log(f"업로드 실패 ({attempt}/{UPLOAD_RETRIES}) -> {exc}")
            if attempt < UPLOAD_RETRIES:
                time.sleep(UPLOAD_SLEEP_SEC)
    raise last_error


def main():
    session = make_session()
    csv_text = fetch_csv_text(session, SOURCE_INPUT_URL)
    source_df = read_source_sheet(csv_text)
    log(f"원본 시트 읽기 완료 -> {source_df.shape[0]}행, {source_df.shape[1]}열")
    upload_to_google_sheet(session, source_df)


if __name__ == "__main__":
    main()
