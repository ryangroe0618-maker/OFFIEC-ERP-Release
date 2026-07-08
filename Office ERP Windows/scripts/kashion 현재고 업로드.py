# -*- coding: utf-8 -*-
import json
import math
import re
import time
from io import StringIO

import pandas as pd
import requests


INPUT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=0&single=true&output=csv"
EXCLUDE_ITEM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRnjBxdJoQqxzPA4e_1rFONzV6VgQ7OxD96IodGHyBZfMm_65VdoGTosK0p3wwP5zEl1CgTINsuS3yL/pub?gid=2048040204&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwJNkRKbSbVGVZPTUzZCQiL37pHR_7vYG1q-RLOdJnHqBtx5JE-DJ9VLC3mZnR_ihtEJA/exec"
SPREADSHEET_ID = "1M65xPYDmXd8xmHkvBZp1bqxjG9pca1LWWBfEHKnHm-o"
TARGET_SHEET_NAME = "总库存"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 180)
RETRIES = 3
SLEEP_SEC = 2

INPUT_COLUMN_INDEXES = list(range(0, 9)) + [11, 12]
INPUT_HEADERS = [
    "브랜드",
    "코드",
    "바코드",
    "품번",
    "상품명",
    "컬러",
    "사이즈",
    "EU",
    "최초가",
    "공급가",
    "현재고",
]
OUTPUT_HEADERS = [
    "品牌",
    "款式",
    "条形码",
    "货号",
    "商品名",
    "颜色",
    "尺码",
    "EU",
    "发售价",
    "供货价",
    "库存",
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


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("\\", "/")
    text = re.sub(r"\s+", "", text)
    return text


def strip_dataframe_text(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(lambda col: col.fillna("").astype(str).str.strip())


def to_number_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.fillna("").astype(str).str.strip().str.replace(",", "", regex=False),
        errors="coerce",
    ).fillna(0.0)


def round_half_up(value: float) -> int:
    return int(math.floor(value + 0.5))


def format_int(value: float) -> str:
    if pd.isna(value):
        return ""
    return str(int(value))


def parse_input_sheet(csv_text: str) -> pd.DataFrame:
    raw_df = read_csv_text(csv_text)
    if raw_df.shape[1] <= max(INPUT_COLUMN_INDEXES):
        raise ValueError(
            f"현재고 시트 열 개수가 부족합니다. 필요한 열 수: {max(INPUT_COLUMN_INDEXES) + 1}, 실제 열 수: {raw_df.shape[1]}"
        )

    df = raw_df.iloc[:, INPUT_COLUMN_INDEXES].copy()
    df.columns = INPUT_HEADERS
    return strip_dataframe_text(df)


def parse_exclude_item_sheet(csv_text: str) -> set[str]:
    df = pd.read_csv(StringIO(csv_text), dtype=str, header=2).fillna("")
    if df.shape[1] < 1:
        raise ValueError(f"제외 품번 정보 시트 열 개수가 부족합니다. 실제 열 수: {df.shape[1]}")

    exclude_series = df.iloc[:, 0].fillna("").astype(str).apply(clean_text)
    exclude_series = exclude_series[exclude_series != ""]
    return set(exclude_series.tolist())


def build_output_df(input_df: pd.DataFrame, exclude_item_set: set[str]) -> pd.DataFrame:
    output_df = input_df.copy()
    output_df["_품번_key"] = output_df["품번"].apply(clean_text)

    output_df["코드"] = (
        output_df["품번"].astype(str)
        + output_df["컬러"].astype(str)
        + output_df["사이즈"].astype(str)
    )

    output_df["최초가_num"] = to_number_series(output_df["최초가"])
    output_df["공급가_num"] = to_number_series(output_df["공급가"])
    output_df["현재고_num"] = to_number_series(output_df["현재고"])

    is_nf_dc = output_df["브랜드"].eq("THE NORTH FACE (DC)")
    output_df["공급가_num"] = output_df["공급가_num"].where(
        ~is_nf_dc,
        output_df["최초가_num"] * 0.3,
    )
    output_df["공급가_num"] = output_df["공급가_num"].where(
        is_nf_dc,
        output_df["공급가_num"] * 1.05,
    )
    output_df["현재고_num"] = (output_df["현재고_num"] / 2.0 + 0.5).apply(math.floor)

    if exclude_item_set:
        output_df = output_df[~output_df["_품번_key"].isin(exclude_item_set)].copy()

    output_df["최초가"] = output_df["최초가_num"].astype(int).astype(str)
    output_df["공급가"] = output_df["공급가_num"].round().astype(int).astype(str)
    output_df["현재고"] = output_df["현재고_num"].astype(int).astype(str)
    output_df = output_df.drop(columns=["_품번_key"])

    output_df = output_df.rename(
        columns={
            "브랜드": "品牌",
            "코드": "款式",
            "바코드": "条形码",
            "품번": "货号",
            "상품명": "商品名",
            "컬러": "颜色",
            "사이즈": "尺码",
            "최초가": "发售价",
            "공급가": "供货价",
            "현재고": "库存",
        }
    )
    output_df = output_df[OUTPUT_HEADERS].copy()
    return output_df


def upload_to_google_sheet(df: pd.DataFrame):
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": values,
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
    input_csv_text = fetch_csv_text(session, INPUT_CSV_URL, "현재고 시트")
    exclude_csv_text = fetch_csv_text(session, EXCLUDE_ITEM_CSV_URL, "제외 품번 정보 시트")
    input_df = parse_input_sheet(input_csv_text)
    exclude_item_set = parse_exclude_item_sheet(exclude_csv_text)
    output_df = build_output_df(input_df, exclude_item_set)

    log(f"업로드 데이터 생성 완료: {len(output_df)}행")
    upload_to_google_sheet(output_df)


if __name__ == "__main__":
    main()
