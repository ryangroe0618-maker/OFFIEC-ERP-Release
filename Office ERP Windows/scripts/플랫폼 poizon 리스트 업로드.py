# -*- coding: utf-8 -*-
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import StringIO

import pandas as pd
import requests


INPUT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQeEO4OxuiGM8p9aemnFHZ6Z1JlimLAYGEFSizPu125tPuRt3W6IXIMa6xFqRqXd1jNjX9snXNjA31S/pub?gid=0&single=true&output=csv"
SIZE_CHART_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=1121158649&single=true&output=csv"
STOCK_TRANSFORM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"
EXCHANGE_RATE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=295228098&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyXoaf-SaykThpMWV3MkqBFfjriw2lkM9I8-rkxPGnxk_mf0TduYPj9C7SQKPnT6mxN/exec"
SPREADSHEET_ID = "1xczCw81ddbt1xh-PIKF3pVdztZu3Vl_5hkQZnySW9P0"
TARGET_SHEET_NAME = "OUT"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 180)
RETRIES = 3
SLEEP_SEC = 2

# E, H, I, K, Q
INPUT_COLUMN_INDEXES = [4, 7, 8, 10, 16]
INPUT_HEADERS = ["브랜드", "품번", "사이즈", "등록 수량", "정산가"]
OUTPUT_HEADERS = ["브랜드", "코드", "코드 변환", "품번", "사이즈", "EU", "사이즈 변환", "사이즈 변환2", "등록 수량", "현재고", "부족 수량", "정산가", "환율", "KRW", "공급가", "마진"]

STOCK_TRANSFORM_COLUMN_INDEXES = [0, 1, 3, 7, 12, 13, 14, 15, 19, 20]
STOCK_TRANSFORM_HEADERS = [
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


def fetch_all_csv_texts(url_map: dict[str, tuple[str, str]]) -> dict[str, str]:
    results = {}
    with ThreadPoolExecutor(max_workers=min(4, len(url_map))) as executor:
        futures = {
            executor.submit(fetch_csv_text, make_session(), url, label): name
            for name, (url, label) in url_map.items()
        }
        for future in as_completed(futures):
            results[futures[future]] = future.result()
    return results


def read_csv_text(csv_text: str) -> pd.DataFrame:
    return pd.read_csv(StringIO(csv_text), dtype=str).fillna("")


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def to_number(value) -> float:
    text = clean_text(value).replace(",", "")
    if text == "":
        return 0.0
    try:
        return float(text)
    except Exception:
        return 0.0


def format_number(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.4f}".rstrip("0").rstrip(".")


def normalize_brand_from_e(value) -> str:
    text = clean_text(value)
    if re.search("adidas", text):
        return "adidas"
    if re.search("PUMA", text):
        return "PUMA"
    if re.search("THE NORTH FACE", text):
        return "THE NORTH FACE"
    if re.search("FILA", text):
        return "FILA"
    if re.search("CONBERSE|CONVERSE", text, flags=re.IGNORECASE):
        return "CONVERSE"
    return ""


def strip_leading_1100(value: str) -> str:
    if value.startswith("1100"):
        return value[4:]
    return value


def remove_chinese(value: str) -> str:
    return clean_text(re.sub(r"[\u4e00-\u9fff]+", "", value))


def normalize_product_lookup_key(value) -> str:
    return clean_text(remove_chinese(value))


def normalize_product_code(value, brand: str) -> str:
    text = normalize_product_lookup_key(value)
    text = re.sub(r"\(\s*\)", "", text)

    if brand == "PUMA":
        return text.replace("（黑色标）", "").replace("鞋", "")

    if brand == "FILA":
        text = strip_leading_1100(text)
        return text.replace("服", "").replace("鞋", "").replace("_", "").replace("-", "")

    if brand == "THE NORTH FACE":
        text = strip_leading_1100(text)
        text = text.split("-", 1)[0]
        return text.replace("包", "").replace("_", "")

    return text


APPAREL_SIZE_RE = r"(?:W(?:XS|S|M|L|XL|XXL|XXXL|[2-9]XL)|[2-9]XL|XXXXL|XXXL|XXL|XL|XS|S|M|L|FREE|ONE|OS)"
SIZE_TOKEN_RE = rf"(?<![A-Z0-9])(?:{APPAREL_SIZE_RE}|F|[A-Z]{{1,2}}\d{{2,3}}|\d{{1,3}}(?:\.\d+)?)(?![A-Z0-9])"


def normalize_fraction_marks(value: str) -> str:
    return clean_text(str(value).replace("⅔", ".5").replace("½", ".5").replace("⅓", ".5"))


def remove_size_descriptions(value: str) -> str:
    text = clean_text(value)
    text = re.sub(r"^(?:size\s+)?specification\s+", "", text, flags=re.IGNORECASE)
    return clean_text(text)


def is_apparel_size(value: str) -> bool:
    return bool(re.fullmatch(APPAREL_SIZE_RE, clean_text(value), flags=re.IGNORECASE))


def is_size_token(value: str) -> bool:
    return bool(re.fullmatch(SIZE_TOKEN_RE, clean_text(value), flags=re.IGNORECASE))


def pick_size_token(value: str) -> str:
    text = normalize_fraction_marks(remove_size_descriptions(remove_chinese(value)))
    if "/" in text:
        text = text.rsplit("/", 1)[-1]

    text = clean_text(text)
    if is_size_token(text):
        return text.upper() if re.fullmatch(r"[A-Za-z]+", text) else text

    for token in re.findall(SIZE_TOKEN_RE, text, flags=re.IGNORECASE):
        token = clean_text(token)
        if is_size_token(token):
            return token.upper() if re.fullmatch(r"[A-Za-z]+", token) else token

    return text


def extract_size_detail(value: str) -> str:
    text = clean_text(value)

    bracket_match = re.search(r"[（(]([^（）()]*)[）)]", text)
    if bracket_match:
        before_bracket = pick_size_token(text[:bracket_match.start()])
        bracket_text = pick_size_token(bracket_match.group(1) or "")

        if is_apparel_size(before_bracket):
            text = before_bracket
        elif bracket_text and is_size_token(bracket_text):
            text = bracket_text
        else:
            text = re.sub(r"[（(][^（）()]*[）)]", "", text)

    return pick_size_token(text)


def finalize_size_value(size_value: str, brand: str) -> str:
    size_text = clean_text(size_value)
    brand_text = clean_text(brand).upper()

    if brand_text == "ADIDAS" and size_text == "37":
        return "37.5"
    if brand_text == "THE NORTH FACE" and size_text.upper() == "F":
        return "ONE"
    if brand_text == "THE NORTH FACE" and size_text.upper() in {"2XL", "3XL", "4XL"}:
        return {"2XL": "XXL", "3XL": "XXXL", "4XL": "XXXXL"}[size_text.upper()]
    return size_text


def normalize_size_value(value, brand: str = "") -> str:
    text = clean_text(value)
    if re.search(r"\d+\s*寸|\d+\s*[lL]\b", text):
        return "ONE"
    match = re.search(r"\b(EU|EH|JP|KR|SIZE|KOREA|CHN)\b\s*[:：]?\s*(.+)$", text, flags=re.IGNORECASE)
    if match:
        label = match.group(1).upper()
        size_text = extract_size_detail(match.group(2))
        if size_text == "":
            return "ONE"
        if label == "JP" and re.fullmatch(r"\d{2}(?:\.\d+)?", size_text):
            return str(int(float(size_text) * 10))
        return finalize_size_value(size_text, brand)
    fallback_size = pick_size_token(text)
    if fallback_size and is_size_token(fallback_size):
        return finalize_size_value(fallback_size, brand)
    return "ONE"


def normalize_lookup_key(value) -> str:
    return clean_text(value).upper()


def normalize_code_lookup_key(value) -> str:
    return re.sub(r"\s+", "", clean_text(value)).upper()


def make_output_code(product_code: str, converted_size: str) -> str:
    size_text = re.sub(r"^0+(?=\d)", "", clean_text(converted_size))
    return f"{clean_text(product_code)}{size_text}"


def make_converted_output_code(product_code: str, converted_size2: str, fallback_code: str) -> str:
    if clean_text(converted_size2):
        return f"{clean_text(product_code)}{clean_text(converted_size2)}"
    return clean_text(fallback_code)


def build_size_chart_lookup(csv_text: str) -> dict[tuple[str, str, str], str]:
    raw_df = read_csv_text(csv_text)
    if raw_df.shape[1] < 4:
        raise ValueError(f"사이즈표 시트 열 개수가 부족합니다. 필요한 열 수: 4, 실제 열 수: {raw_df.shape[1]}")

    size_df = raw_df.iloc[:, :4].copy()
    size_df.columns = ["브랜드", "품번", "사이즈", "EU"]
    size_df = strip_dataframe_text(size_df)

    lookup = {}
    for _, row in size_df.iterrows():
        brand_text = normalize_brand_from_e(row["브랜드"]) or clean_text(row["브랜드"])
        brand = normalize_lookup_key(brand_text)
        product_code = normalize_lookup_key(normalize_product_code(row["품번"], brand_text))
        eu_size = normalize_lookup_key(row["EU"])
        size_value = clean_text(row["사이즈"])
        if brand and eu_size and size_value:
            lookup.setdefault((brand, product_code, eu_size), size_value)
    return lookup


def apply_size_chart_lookup(brand: str, product_code: str, converted_size: str, size_lookup: dict[tuple[str, str, str], str]) -> str:
    brand_key = normalize_lookup_key(brand)
    product_code_key = normalize_lookup_key(normalize_product_code(product_code, brand))
    if brand_key == "THE NORTH FACE":
        product_code_key = product_code_key[:7]

    key = (
        brand_key,
        product_code_key,
        normalize_lookup_key(converted_size),
    )
    if brand_key == "THE NORTH FACE":
        return size_lookup.get(key, converted_size)

    fallback_key = (
        brand_key,
        "",
        normalize_lookup_key(converted_size),
    )
    return size_lookup.get(key, size_lookup.get(fallback_key, converted_size))


def load_stock_transform_df(csv_text: str) -> pd.DataFrame:
    raw_df = read_csv_text(csv_text)
    if raw_df.shape[1] <= max(STOCK_TRANSFORM_COLUMN_INDEXES):
        raise ValueError(
            f"현재고 변환 시트 열 개수가 부족합니다. 필요한 열 수: {max(STOCK_TRANSFORM_COLUMN_INDEXES) + 1}, 실제 열 수: {raw_df.shape[1]}"
        )

    stock_df = raw_df.iloc[:, STOCK_TRANSFORM_COLUMN_INDEXES].copy()
    stock_df.columns = STOCK_TRANSFORM_HEADERS
    return strip_dataframe_text(stock_df)


def build_stock_lookups(stock_df: pd.DataFrame) -> tuple[dict[str, str], dict[str, str], dict[str, str], dict[str, str]]:
    stock_size_lookup = {}
    supply_price_lookup = {}
    current_stock_lookup = {}
    eu_lookup = {}
    code_series_list = [stock_df[col].tolist() for col in ["코드", "변환코드1", "변환코드2", "변환코드3"]]

    for code_values, product_code, size_value, supply_price, current_stock, eu_value in zip(
        zip(*code_series_list),
        stock_df["품번"].tolist(),
        stock_df["사이즈"].tolist(),
        stock_df["공급가"].tolist(),
        stock_df["현재고"].tolist(),
        stock_df["EU"].tolist(),
    ):
        size_value = clean_text(size_value)
        for code_value in code_values:
            code_key = normalize_code_lookup_key(code_value)
            if code_key and size_value:
                stock_size_lookup.setdefault(code_key, size_value)

        product_code_key = normalize_lookup_key(normalize_product_code(product_code, ""))
        supply_price = clean_text(supply_price)
        if product_code_key and supply_price:
            supply_price_lookup.setdefault(product_code_key, supply_price)

        code_key = normalize_code_lookup_key(code_values[0])
        current_stock = clean_text(current_stock)
        eu_value = clean_text(eu_value)
        if code_key:
            current_stock_lookup.setdefault(code_key, current_stock)
            eu_lookup.setdefault(code_key, eu_value)

    return stock_size_lookup, supply_price_lookup, current_stock_lookup, eu_lookup


def lookup_stock_size(output_code: str, stock_size_lookup: dict[str, str]) -> str:
    return stock_size_lookup.get(normalize_code_lookup_key(output_code), "")


def lookup_supply_price(product_code: str, supply_price_lookup: dict[str, str]) -> str:
    return supply_price_lookup.get(normalize_lookup_key(normalize_product_lookup_key(product_code)), "")


def lookup_current_stock(output_code: str, current_stock_lookup: dict[str, str]) -> str:
    return current_stock_lookup.get(normalize_code_lookup_key(output_code), "")


def lookup_eu(output_code: str, eu_lookup: dict[str, str]) -> str:
    return eu_lookup.get(normalize_code_lookup_key(output_code), "")


def get_exchange_rate(csv_text: str) -> str:
    exchange_df = read_csv_text(csv_text)
    if exchange_df.shape[0] < 1 or exchange_df.shape[1] < 2:
        raise ValueError(
            f"환율 시트에서 B2 셀을 찾을 수 없습니다. 실제 행/열: {exchange_df.shape[0] + 1}행, {exchange_df.shape[1]}열"
        )
    return clean_text(exchange_df.iloc[0, 1])


def strip_dataframe_text(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(lambda col: col.fillna("").astype(str).map(clean_text))


def parse_input_sheet(
    csv_text: str,
    size_lookup: dict[tuple[str, str, str], str] | None = None,
    stock_size_lookup: dict[str, str] | None = None,
    supply_price_lookup: dict[str, str] | None = None,
    current_stock_lookup: dict[str, str] | None = None,
    eu_lookup: dict[str, str] | None = None,
    exchange_rate: str = "",
) -> pd.DataFrame:
    size_lookup = size_lookup or {}
    stock_size_lookup = stock_size_lookup or {}
    supply_price_lookup = supply_price_lookup or {}
    current_stock_lookup = current_stock_lookup or {}
    eu_lookup = eu_lookup or {}
    raw_df = read_csv_text(csv_text)
    if raw_df.shape[1] <= max(INPUT_COLUMN_INDEXES):
        raise ValueError(
            f"플랫폼 poizon 시트 열 개수가 부족합니다. 필요한 열 수: {max(INPUT_COLUMN_INDEXES) + 1}, 실제 열 수: {raw_df.shape[1]}"
        )

    df = raw_df.iloc[:, INPUT_COLUMN_INDEXES].copy()
    df.columns = INPUT_HEADERS
    df = strip_dataframe_text(df)
    df["브랜드"] = df["브랜드"].map(normalize_brand_from_e)
    df["품번"] = [
        normalize_product_code(product_code, brand)
        for product_code, brand in zip(df["품번"], df["브랜드"])
    ]
    converted_sizes = [
        normalize_size_value(size, brand)
        for size, brand in zip(df["사이즈"], df["브랜드"])
    ]
    df["사이즈 변환"] = [
        apply_size_chart_lookup(brand, product_code, converted_size, size_lookup)
        for brand, product_code, converted_size in zip(df["브랜드"], df["품번"], converted_sizes)
    ]
    df["코드"] = [
        make_output_code(product_code, converted_size)
        for product_code, converted_size in zip(df["품번"], df["사이즈 변환"])
    ]
    stock_size_matches = [
        lookup_stock_size(code, stock_size_lookup)
        for code in df["코드"]
    ]
    df["사이즈 변환2"] = [
        stock_size if clean_text(stock_size) else converted_size
        for stock_size, converted_size in zip(stock_size_matches, df["사이즈 변환"])
    ]
    df["사이즈 변환2"] = df["사이즈 변환2"].where(
        df["사이즈 변환2"].map(clean_text) != "",
        df["사이즈 변환"],
    )
    df["코드 변환"] = [
        make_converted_output_code(product_code, stock_size, code)
        for product_code, stock_size, code in zip(df["품번"], stock_size_matches, df["코드"])
    ]
    df["EU"] = df["코드 변환"].map(lambda code: lookup_eu(code, eu_lookup))
    df["현재고"] = df["코드 변환"].map(lambda code: lookup_current_stock(code, current_stock_lookup))
    df["부족 수량"] = [
        format_number(to_number(current_stock) - to_number(register_qty))
        for current_stock, register_qty in zip(df["현재고"], df["등록 수량"])
    ]
    df["환율"] = clean_text(exchange_rate)
    df["KRW"] = [
        format_number(to_number(settlement_price) * to_number(exchange_rate))
        for settlement_price in df["정산가"]
    ]
    raw_supply_prices = [
        lookup_supply_price(product_code, supply_price_lookup)
        for product_code in df["품번"]
    ]
    df["공급가"] = [
        format_number(to_number(supply_price) / 1.1) if clean_text(supply_price) else ""
        for supply_price in raw_supply_prices
    ]
    df["마진"] = [
        format_number(to_number(krw) - to_number(supply_price)) if clean_text(supply_price) else ""
        for krw, supply_price in zip(df["KRW"], df["공급가"])
    ]
    return df[OUTPUT_HEADERS].sort_values(
        by=["브랜드", "품번", "사이즈"],
        kind="mergesort",
    ).reset_index(drop=True)


def upload_to_google_sheet(df: pd.DataFrame):
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
    csv_texts = fetch_all_csv_texts({
        "source": (INPUT_CSV_URL, "플랫폼 poizon 시트"),
        "size_chart": (SIZE_CHART_CSV_URL, "사이즈표 시트"),
        "stock_transform": (STOCK_TRANSFORM_CSV_URL, "현재고 변환 시트"),
        "exchange_rate": (EXCHANGE_RATE_CSV_URL, "환율 시트"),
    })

    size_lookup = build_size_chart_lookup(csv_texts["size_chart"])
    stock_transform_df = load_stock_transform_df(csv_texts["stock_transform"])
    stock_size_lookup, supply_price_lookup, current_stock_lookup, eu_lookup = build_stock_lookups(stock_transform_df)
    exchange_rate = get_exchange_rate(csv_texts["exchange_rate"])
    output_df = parse_input_sheet(
        csv_texts["source"],
        size_lookup,
        stock_size_lookup,
        supply_price_lookup,
        current_stock_lookup,
        eu_lookup,
        exchange_rate,
    )

    log(f"업로드 데이터 생성 완료: {len(output_df)}행")
    upload_to_google_sheet(output_df)


if __name__ == "__main__":
    main()
