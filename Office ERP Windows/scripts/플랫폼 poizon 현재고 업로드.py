# -*- coding: utf-8 -*-
import json
import time
from io import StringIO

import pandas as pd
import requests


INPUT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyXoaf-SaykThpMWV3MkqBFfjriw2lkM9I8-rkxPGnxk_mf0TduYPj9C7SQKPnT6mxN/exec"
SPREADSHEET_ID = "1xczCw81ddbt1xh-PIKF3pVdztZu3Vl_5hkQZnySW9P0"
TARGET_SHEET_NAME = "현재고"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 180)
RETRIES = 3
SLEEP_SEC = 2

INPUT_COLUMN_INDEXES = [0, 1, 3, 7, 12, 13, 14, 15, 19, 20]
OUTPUT_HEADERS = [
    "브랜드",
    "코드",
    "품번",
    "사이즈",
    "변환코드1",
    "변환코드2",
    "변환코드3",
    "EU",
    "공급가",
    "현재고",
]

START_TIME = time.time()


def log(message: str):
    elapsed = time.time() - START_TIME
    print(f"[{elapsed:6.1f}s] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def validate_apps_script_response(response_text: str) -> dict:
    text = (response_text or "").strip()
    try:
        response_json = json.loads(text)
    except Exception:
        response_json = None

    if isinstance(response_json, dict):
        if response_json.get("ok") is False:
            raise RuntimeError(
                response_json.get("error")
                or response_json.get("message")
                or text
            )
        return response_json

    lowered = text.lower()
    if "<html" in lowered or "<!doctype html" in lowered:
        if "doget" in lowered:
            raise RuntimeError(
                "Apps Script가 doGet 오류 HTML을 반환했습니다. "
                "웹앱 배포 코드에 doGet/doPost가 있는지 확인하고 새 배포로 업데이트해 주세요."
            )
        if "dopost" in lowered:
            raise RuntimeError(
                "Apps Script가 doPost 오류 HTML을 반환했습니다. "
                "웹앱 배포 코드에 doPost가 있는지 확인하고 새 배포로 업데이트해 주세요."
            )
        raise RuntimeError(f"Apps Script가 JSON 대신 HTML을 반환했습니다: {text[:300]}")

    raise RuntimeError(f"Apps Script 응답을 JSON으로 확인할 수 없습니다: {text[:300]}")


def fetch_csv_text(session: requests.Session, url: str, label: str) -> str:
    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            log(f"{label} 다운로드 시작 ({attempt}/{RETRIES})")
            response = session.get(url, timeout=DOWNLOAD_TIMEOUT)
            response.raise_for_status()
            response.encoding = "utf-8"
            log(f"{label} 다운로드 완료")
            return response.text
        except Exception as exc:
            last_error = exc
            log(f"{label} 다운로드 실패 ({attempt}/{RETRIES}): {exc}")
            if attempt < RETRIES:
                time.sleep(SLEEP_SEC)
    raise last_error


def read_csv_text(csv_text: str) -> pd.DataFrame:
    return pd.read_csv(StringIO(csv_text), dtype=str).fillna("")


def strip_dataframe_text(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(lambda col: col.fillna("").astype(str).str.strip())


def parse_input_sheet(csv_text: str) -> pd.DataFrame:
    raw_df = read_csv_text(csv_text)
    if raw_df.shape[1] <= max(INPUT_COLUMN_INDEXES):
        raise ValueError(
            f"현재고 변환 시트 열 개수가 부족합니다. 필요한 열 수: {max(INPUT_COLUMN_INDEXES) + 1}, 실제 열 수: {raw_df.shape[1]}"
        )

    df = raw_df.iloc[:, INPUT_COLUMN_INDEXES].copy()
    df.columns = OUTPUT_HEADERS
    return strip_dataframe_text(df)


def upload_to_google_sheet(df: pd.DataFrame):
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": values,
        "clear": True,
        "clearRangeOnly": True,
        "startRow": 1,
        "startCol": 1,
        "clearCols": len(OUTPUT_HEADERS),
    }

    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            log(f"구글 시트 업로드 시작 ({attempt}/{RETRIES}) - {TARGET_SHEET_NAME} / {len(df)}행")
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            response_text = (response.text or "").strip()
            validate_apps_script_response(response_text)

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

    log(f"업로드 데이터 생성 완료: {len(output_df)}행")
    upload_to_google_sheet(output_df)


if __name__ == "__main__":
    main()
