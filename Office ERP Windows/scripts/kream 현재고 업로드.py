# -*- coding: utf-8 -*-
import json
import time
from io import StringIO

import pandas as pd
import requests


INPUT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxv-EAKz_2W-u3kaG9YZ3r6qNZQz52P-bgAHEZ30DRvcfaVjbPWZ5dQqaXkWXtCwwlf/exec"
SPREADSHEET_ID = "1GzCiXK6XUzB44Ot-BwvNRKEJNz4Gz94pGben0Fb8aV4"
TARGET_SHEET_NAME = "현재고"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 180)
RETRIES = 3
SLEEP_SEC = 2

# B, H, M:O, V:AG
INPUT_COLUMN_INDEXES = [
    1, 7,
    12, 13, 14,
    21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32,
]

START_TIME = time.time()


class DataValidationError(ValueError):
    pass


def log(message: str):
    elapsed = time.time() - START_TIME
    print(f"[{elapsed:6.1f}s] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(session: requests.Session, url: str, label: str) -> str:
    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            log(f"{label} 다운로드 시작 ({attempt}/{RETRIES})")
            response = session.get(url, timeout=DOWNLOAD_TIMEOUT)
            response.raise_for_status()
            response.encoding = "utf-8"
            text = response.text
            if not text.strip():
                raise DataValidationError(f"{label} CSV 응답이 비어 있습니다.")
            log(f"{label} 다운로드 완료")
            return text
        except Exception as exc:
            last_error = exc
            log(f"{label} 다운로드 실패 ({attempt}/{RETRIES}): {exc}")
            if attempt < RETRIES:
                time.sleep(SLEEP_SEC)
    raise last_error


def read_selected_csv_text(csv_text: str) -> pd.DataFrame:
    try:
        return pd.read_csv(
            StringIO(csv_text),
            dtype=str,
            usecols=INPUT_COLUMN_INDEXES,
        ).fillna("")
    except ValueError as exc:
        raise DataValidationError(
            "현재고 변환 시트 열 개수가 부족하거나 선택 열을 읽을 수 없습니다. "
            f"필요한 최대 열 수: {max(INPUT_COLUMN_INDEXES) + 1}"
        ) from exc


def strip_dataframe_text(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(lambda col: col.fillna("").astype(str).str.strip())


def parse_input_sheet(csv_text: str) -> pd.DataFrame:
    output_df = read_selected_csv_text(csv_text)
    if output_df.shape[1] != len(INPUT_COLUMN_INDEXES):
        raise DataValidationError(
            f"선택된 열 개수가 맞지 않습니다. 필요: {len(INPUT_COLUMN_INDEXES)}, 실제: {output_df.shape[1]}"
        )

    output_df = strip_dataframe_text(output_df)
    output_df = output_df.loc[output_df.ne("").any(axis=1)].copy()

    if output_df.empty:
        raise DataValidationError("업로드할 현재고 데이터가 없습니다.")

    return output_df


def upload_to_google_sheet(session: requests.Session, df: pd.DataFrame):
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": values,
        "clear": True,
    }

    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            log(f"구글 시트 업로드 시작 ({attempt}/{RETRIES}) - {TARGET_SHEET_NAME} / {len(df)}행")
            response = session.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            response_text = (response.text or "").strip()

            try:
                response_json = json.loads(response_text)
            except Exception:
                response_json = None

            if isinstance(response_json, dict) and response_json.get("ok") is False:
                raise RuntimeError(response_json.get("message", "Apps Script 처리 실패"))

            log("구글 시트 업로드 완료")
            log(f"응답: {response_text}")
            return
        except Exception as exc:
            last_error = exc
            log(f"구글 시트 업로드 실패 ({attempt}/{RETRIES}): {exc}")
            if attempt < RETRIES:
                time.sleep(SLEEP_SEC)
    raise last_error


def main():
    session = make_session()
    csv_text = fetch_csv_text(session, INPUT_CSV_URL, "현재고 변환 시트")
    output_df = parse_input_sheet(csv_text)
    log(f"업로드 데이터 생성 완료: {len(output_df)}행, {output_df.shape[1]}열")
    upload_to_google_sheet(session, output_df)


if __name__ == "__main__":
    main()
