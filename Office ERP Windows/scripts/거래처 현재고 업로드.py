# -*- coding: utf-8 -*-
from concurrent.futures import ThreadPoolExecutor
import json
import re
import time
from io import StringIO

import pandas as pd
import requests


INPUT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=0&single=true&output=csv"
NF_DC_INFO_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=599962083&single=true&output=csv"
EXCLUDE_ITEM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRnjBxdJoQqxzPA4e_1rFONzV6VgQ7OxD96IodGHyBZfMm_65VdoGTosK0p3wwP5zEl1CgTINsuS3yL/pub?gid=2048040204&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbz8R_CBBbNtdjSpWpb1M6DjW-8izufgGy2qikBfx48R96A7RI52WRB_qDTcrYgd7WwF/exec"
SPREADSHEET_ID = "1mFA9Cu5xsKBI7Hce-bah5QfIQv80AgR1uER7COB5D-U"
TARGET_SHEET_NAME = "현재고"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 180)
RETRIES = 3
SLEEP_SEC = 2

TARGET_HEADERS = [
    "브랜드",
    "코드",
    "바코드",
    "품번",
    "상품명",
    "컬러",
    "사이즈",
    "최초가",
    "할인율",
    "할인가",
    "공급율",
    "현재고",
]

INPUT_HEADERS = [
    "브랜드",
    "코드",
    "바코드",
    "품번",
    "상품명",
    "컬러",
    "사이즈",
    "최초가",
    "할인율",
    "할인가",
    "현재고",
]

INPUT_COLUMN_INDEXES = [0, 1, 2, 3, 4, 5, 6, 8, 9, 10, 12]

START_TIME = time.time()


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


def format_ratio_percent(value) -> str:
    text = str(value).strip().replace("%", "").replace(",", "")
    if text == "":
        return ""
    try:
        number = float(text)
    except Exception:
        return ""

    if number > 1:
        number = number / 100.0

    percent = number * 100
    if float(percent).is_integer():
        return f"{int(percent)}%"
    return f"{percent:.2f}".rstrip("0").rstrip(".") + "%"


def normalize_header(value) -> str:
    return clean_text(value).upper()


def find_column(df: pd.DataFrame, *candidates: str) -> str:
    normalized_map = {normalize_header(col): col for col in df.columns}
    for candidate in candidates:
        found = normalized_map.get(normalize_header(candidate))
        if found:
            return found
    raise ValueError(f"필수 열을 찾을 수 없습니다: {', '.join(candidates)}")


def strip_dataframe_text(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(lambda col: col.fillna("").astype(str).str.strip())


def parse_input_sheet(csv_text: str) -> pd.DataFrame:
    raw_df = read_csv_text(csv_text)
    if raw_df.shape[1] <= max(INPUT_COLUMN_INDEXES):
        raise ValueError(
            f"현재고 취합 시트 열 개수가 부족합니다. 필요한 열 수: {max(INPUT_COLUMN_INDEXES) + 1}, 실제 열 수: {raw_df.shape[1]}"
        )

    df = raw_df.iloc[:, INPUT_COLUMN_INDEXES].copy()
    df.columns = INPUT_HEADERS
    return strip_dataframe_text(df)


def parse_nf_dc_info_sheet(csv_text: str) -> pd.DataFrame:
    df = read_csv_text(csv_text)
    item_col = find_column(df, "품번")
    supply_rate_col = find_column(df, "업체 공급율", "업체공급율", "공급율", "공급비율")

    result_df = df[[item_col, supply_rate_col]].copy()
    result_df.columns = ["품번", "공급율"]
    result_df["품번"] = result_df["품번"].apply(clean_text)
    result_df["공급율"] = result_df["공급율"].apply(format_ratio_percent)
    result_df = result_df[result_df["품번"] != ""].copy()
    return result_df.drop_duplicates(subset=["품번"], keep="first").reset_index(drop=True)


def parse_exclude_item_sheet(csv_text: str) -> set[str]:
    df = pd.read_csv(StringIO(csv_text), dtype=str, header=2).fillna("")
    if df.shape[1] <= 3:
        raise ValueError(f"제외 품번 정보 시트 열 개수가 부족합니다. 실제 열 수: {df.shape[1]}")

    exclude_series = df.iloc[:, 3].fillna("").astype(str).apply(clean_text)
    exclude_series = exclude_series[exclude_series != ""]
    return set(exclude_series.tolist())


def build_output_df(input_df: pd.DataFrame, nf_dc_info_df: pd.DataFrame, exclude_item_set: set[str]) -> pd.DataFrame:
    output_df = input_df.copy()
    output_df["_품번_key"] = output_df["품번"].apply(clean_text)
    nf_dc_info_df = nf_dc_info_df.rename(columns={"품번": "_품번_key"})

    output_df = output_df.merge(nf_dc_info_df, on="_품번_key", how="left")
    output_df["공급율"] = output_df["공급율"].fillna("").astype(str).str.strip()
    if exclude_item_set:
        output_df = output_df[~output_df["_품번_key"].isin(exclude_item_set)].copy()
    output_df = output_df.drop(columns=["_품번_key"])

    output_df["현재고"] = output_df["현재고"].fillna("").astype(str).str.strip()
    output_df = output_df[
        [
            "브랜드",
            "코드",
            "바코드",
            "품번",
            "상품명",
            "컬러",
            "사이즈",
            "최초가",
            "할인율",
            "할인가",
            "공급율",
            "현재고",
        ]
    ].copy()
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
    def download_job(url: str, label: str) -> str:
        return fetch_csv_text(make_session(), url, label)

    with ThreadPoolExecutor(max_workers=3) as executor:
        input_future = executor.submit(download_job, INPUT_CSV_URL, "현재고 취합 시트")
        nf_dc_future = executor.submit(download_job, NF_DC_INFO_CSV_URL, "노스페이스 다년차 정보 시트")
        exclude_future = executor.submit(download_job, EXCLUDE_ITEM_CSV_URL, "제외 품번 정보 시트")

        input_csv_text = input_future.result()
        nf_dc_csv_text = nf_dc_future.result()
        exclude_csv_text = exclude_future.result()

    input_df = parse_input_sheet(input_csv_text)
    nf_dc_info_df = parse_nf_dc_info_sheet(nf_dc_csv_text)
    exclude_item_set = parse_exclude_item_sheet(exclude_csv_text)
    output_df = build_output_df(input_df, nf_dc_info_df, exclude_item_set)

    log(f"업로드 데이터 생성 완료: {len(output_df)}행")
    upload_to_google_sheet(output_df)


if __name__ == "__main__":
    main()
