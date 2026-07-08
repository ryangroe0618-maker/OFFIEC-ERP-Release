# -*- coding: utf-8 -*-
import json
import re
import time
from io import StringIO
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import pandas as pd
import requests


SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR0qZuV-fBLGX8vinl0MH3xEMz6g1WntsyX1WY9d9KnXOg5Uzuav8DfHLdisyp3VCv4hocMtD5q_EJD/pub?gid=0&single=true&output=csv"
STOCK_TRANSFORM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxVzGDXRAFM0BzBXT1mWA6rD3s7Co8y1VLVeBzz_fK58K4F9uta2MVgs2LgbYT8itcuSw/exec"
SPREADSHEET_ID = "1fe3j49JdTGfT6tnjt4ipU_hINcclhtfg9uEpLb7Szz4"
TARGET_SHEET_NAME = "OUT"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 180)
RETRIES = 3
SLEEP_SEC = 2

# C, E, H, I, N
INPUT_COLUMN_INDEXES = [2, 4, 7, 8, 13]
OUTPUT_COLUMNS = [
    "구분",
    "브랜드",
    "변환 코드",
    "품번",
    "변환 사이즈",
    "등록 수량",
    "현재고",
    "차이 수량",
    "등록 판매가",
    "최근 거래가",
    "공급가",
    "마진",
]
GROUP_COLUMNS = [
    "구분",
    "브랜드",
    "품번",
    "변환 사이즈",
    "등록 판매가",
    "최근 거래가",
    "공급가",
    "변환 코드",
]
STOCK_REQUIRED_COLUMNS = [
    "브랜드",
    "코드",
    "품번",
    "사이즈",
    "변환코드1",
    "변환코드2",
    "변환코드3",
    "공급가",
    "현재고",
]

START_TIME = time.time()


class DataValidationError(ValueError):
    pass


def log(message: str):
    elapsed = time.time() - START_TIME
    print(f"[{elapsed:6.1f}s] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0",
            "Cache-Control": "no-cache, no-store, max-age=0",
            "Pragma": "no-cache",
        }
    )
    return session


def add_cache_buster(url: str) -> str:
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["_ts"] = str(time.time_ns())
    return urlunsplit(
        (
            parts.scheme,
            parts.netloc,
            parts.path,
            urlencode(query),
            parts.fragment,
        )
    )


def fetch_csv_text(session: requests.Session, url: str, label: str) -> str:
    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            log(f"{label} 다운로드 시작 ({attempt}/{RETRIES})")
            response = session.get(
                add_cache_buster(url),
                timeout=DOWNLOAD_TIMEOUT,
            )
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


def normalize_size(value: str) -> str:
    text = str(value or "").strip()

    if "(" in text:
        start = text.find("(")
        end = text.find(")", start + 1)
        if end == -1:
            return ""
        text = text[start + 1:end]
    elif text.startswith("W"):
        text = text[1:]

    return "".join(re.findall(r"[A-Za-z0-9]", text))


def normalize_code_key(value: str) -> str:
    return str(value or "").strip().upper()


def compact_code_key(value: str) -> str:
    return "".join(re.findall(r"[A-Z0-9]", normalize_code_key(value)))


def normalize_product_key(value: str) -> str:
    return str(value or "").strip().upper()


def parse_stock_transform(csv_text: str) -> pd.DataFrame:
    stock_df = pd.read_csv(
        StringIO(csv_text),
        dtype=str,
        keep_default_na=False,
    )
    missing = [
        column for column in STOCK_REQUIRED_COLUMNS if column not in stock_df.columns
    ]
    if missing:
        raise DataValidationError(
            f"현재고 변환 시트에 필요한 열이 없습니다: {missing}"
        )

    stock_df = stock_df.loc[:, STOCK_REQUIRED_COLUMNS].copy()
    for column in stock_df.columns:
        stock_df[column] = stock_df[column].fillna("").astype(str).str.strip()
    return stock_df


def build_stock_lookups(
    stock_df: pd.DataFrame,
) -> tuple[dict[str, str], dict[str, str], dict[str, str], dict[str, int]]:
    code_to_size = {}
    brand_by_product = {}
    supply_by_product = {}
    stock_by_code = {}

    for row in stock_df.itertuples(index=False):
        product_key = normalize_product_key(row.품번)
        if product_key:
            if row.브랜드:
                brand_by_product.setdefault(product_key, row.브랜드)
            if row.공급가:
                supply_by_product.setdefault(product_key, row.공급가)

        stock_value = pd.to_numeric(
            str(row.현재고 or "").replace(",", ""),
            errors="coerce",
        )
        stock_value = 0 if pd.isna(stock_value) else int(stock_value)
        canonical_code = str(getattr(row, "코드", "") or "").strip()
        for stock_code_key in (
            normalize_code_key(canonical_code),
            compact_code_key(canonical_code),
        ):
            if stock_code_key:
                stock_by_code.setdefault(stock_code_key, stock_value)

        converted_size = str(row.사이즈 or "").strip()
        if not converted_size:
            continue

        for code in (row.변환코드1, row.변환코드2, row.변환코드3):
            code_key = normalize_code_key(code)
            compact_key = compact_code_key(code)
            if code_key:
                code_to_size.setdefault(code_key, converted_size)
            if compact_key:
                code_to_size.setdefault(compact_key, converted_size)

    return code_to_size, brand_by_product, supply_by_product, stock_by_code


def lookup_converted_size(code: str, code_to_size: dict[str, str]) -> str:
    code_key = normalize_code_key(code)
    compact_key = compact_code_key(code)
    return code_to_size.get(code_key, code_to_size.get(compact_key, ""))


def lookup_current_stock(code: str, stock_by_code: dict[str, int]) -> int:
    code_key = normalize_code_key(code)
    compact_key = compact_code_key(code)
    return stock_by_code.get(code_key, stock_by_code.get(compact_key, 0))


def subtract_price(sale_price: str, supply_price: str):
    sale_value = pd.to_numeric(
        str(sale_price or "").replace(",", ""),
        errors="coerce",
    )
    supply_value = pd.to_numeric(
        str(supply_price or "").replace(",", ""),
        errors="coerce",
    )
    if pd.isna(sale_value) or pd.isna(supply_value):
        return ""
    return int(sale_value - supply_value)


def parse_and_aggregate(
    source_csv_text: str,
    stock_csv_text: str,
) -> pd.DataFrame:
    try:
        raw_df = pd.read_csv(
            StringIO(source_csv_text),
            header=None,
            skiprows=3,
            dtype=str,
            keep_default_na=False,
            usecols=INPUT_COLUMN_INDEXES,
        )
    except ValueError as exc:
        raise DataValidationError(
            "KREAM 시트에서 C, E, H, I, N 열을 읽을 수 없습니다."
        ) from exc

    if raw_df.shape[1] != len(INPUT_COLUMN_INDEXES):
        raise DataValidationError(
            f"선택된 열 개수가 맞지 않습니다. 필요: {len(INPUT_COLUMN_INDEXES)}, 실제: {raw_df.shape[1]}"
        )

    raw_df.columns = [
        "구분",
        "품번",
        "원본사이즈",
        "등록 판매가",
        "최근 거래가",
    ]
    for column in raw_df.columns:
        raw_df[column] = raw_df[column].fillna("").astype(str).str.strip()

    raw_df = raw_df.loc[raw_df.ne("").any(axis=1)].copy()
    raw_df["원본사이즈정리"] = raw_df["원본사이즈"].map(normalize_size)
    raw_df["원본코드"] = raw_df["품번"] + raw_df["원본사이즈정리"]

    valid_mask = raw_df["품번"].ne("") & raw_df["원본사이즈정리"].ne("")
    skipped_count = int((~valid_mask).sum())
    if skipped_count:
        log(f"품번 또는 원본 사이즈가 없는 행 제외: {skipped_count}행")
    raw_df = raw_df.loc[valid_mask].copy()

    if raw_df.empty:
        raise DataValidationError("변환 후 업로드할 KREAM 데이터가 없습니다.")

    stock_df = parse_stock_transform(stock_csv_text)
    (
        code_to_size,
        brand_by_product,
        supply_by_product,
        stock_by_code,
    ) = build_stock_lookups(stock_df)

    raw_df["변환 사이즈"] = raw_df["원본코드"].map(
        lambda code: lookup_converted_size(code, code_to_size)
    )
    unmatched_count = int(raw_df["변환 사이즈"].eq("").sum())
    if unmatched_count:
        log(
            f"현재고 변환 시트에서 변환코드를 찾지 못해 "
            f"최초 변환 사이즈를 사용한 행: {unmatched_count}행"
        )
        unmatched_mask = raw_df["변환 사이즈"].eq("")
        raw_df.loc[unmatched_mask, "변환 사이즈"] = raw_df.loc[
            unmatched_mask,
            "원본사이즈정리",
        ]

    product_keys = raw_df["품번"].map(normalize_product_key)
    raw_df["브랜드"] = product_keys.map(brand_by_product).fillna("")
    raw_df["공급가"] = product_keys.map(supply_by_product).fillna("")
    raw_df["등록 수량"] = 1
    raw_df["변환 코드"] = raw_df["품번"] + raw_df["변환 사이즈"]

    result_df = (
        raw_df.groupby(
            GROUP_COLUMNS,
            as_index=False,
            sort=False,
            dropna=False,
        )["등록 수량"]
        .sum()
    )
    result_df["현재고"] = result_df["변환 코드"].map(
        lambda code: lookup_current_stock(code, stock_by_code)
    )
    result_df["차이 수량"] = result_df["현재고"] - result_df["등록 수량"]
    result_df["마진"] = [
        subtract_price(sale_price, supply_price)
        for sale_price, supply_price in zip(
            result_df["등록 판매가"],
            result_df["공급가"],
        )
    ]
    result_df = result_df.loc[:, OUTPUT_COLUMNS]
    return result_df


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
            log(
                f"구글 시트 업로드 시작 ({attempt}/{RETRIES}) "
                f"- {TARGET_SHEET_NAME} / {len(df)}행"
            )
            response = session.post(
                WEB_APP_URL,
                json=payload,
                timeout=UPLOAD_TIMEOUT,
            )
            response.raise_for_status()
            response_text = (response.text or "").strip()

            try:
                response_json = json.loads(response_text)
            except (TypeError, json.JSONDecodeError):
                response_json = None

            if isinstance(response_json, dict) and response_json.get("ok") is False:
                raise RuntimeError(
                    response_json.get("message", "Apps Script 처리 실패")
                )

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
    source_csv_text = fetch_csv_text(
        session,
        SOURCE_CSV_URL,
        "플랫폼 KREAM 시트",
    )
    stock_csv_text = fetch_csv_text(
        session,
        STOCK_TRANSFORM_CSV_URL,
        "현재고 변환 시트",
    )
    output_df = parse_and_aggregate(source_csv_text, stock_csv_text)
    log(
        f"취합 완료: {len(output_df)}행, "
        f"총 등록 수량 {output_df['등록 수량'].sum()}개"
    )
    upload_to_google_sheet(session, output_df)


if __name__ == "__main__":
    main()
