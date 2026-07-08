# -*- coding: utf-8 -*-

from datetime import datetime
from io import StringIO
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests


# =========================
# 입력 구글 시트 CSV
# =========================
POIZON_INPUT_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSzpo-DtuISMc_boM0XnjqnY-1hIlD2s_LMhzbvaRWdeNFBxdtO1Z0Fl94s4Dxo52wOwrBfisgDyQYt/pub?gid=0&single=true&output=csv"
SIZE_CHART_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=1121158649&single=true&output=csv"
STOCK_PREP_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"
EXCHANGE_RATE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=295228098&single=true&output=csv"
SQUAREONE_LOCATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSZ4Mgu9j6y27nLBYU8gAhDTfy4eMpvBgvs3oorR3BUCpcgoyf6Z1SllaqsFyos8LcH5DfxoUsN4NYG/pub?gid=289091756&single=true&output=csv"
FILA_PAJU_LOCATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSA47SgFq9QQPg0D3AlBnpJX6q7Yx_Dh66E1ID9MlXTahJjL0FmFVtPgyTEtj4iVj7PvRkCUoCgbjkd/pub?gid=1813802704&single=true&output=csv"

# =========================
# 출력 구글 시트 설정
# =========================
SUMMARY_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxpM0FeOfGZawWdZ0J6C62ttFUZp5PZhB878sOBNQicFt4GIR6zF5osIGjLGGrVJ8T4kA/exec"
SUMMARY_SPREADSHEET_ID = "1IZ1d6exXI-dFO5pFyKrPvhbBMQRdNIg9sXSRlSeJymw"
SUMMARY_SHEET_NAME = "OUT"
SQUAREONE_LOCATION_SHEET_NAME = "스퀘어원 제품 위치"
FILA_PAJU_LOCATION_SHEET_NAME = "휠라 파주 제품 위치"

DOWNLOAD_RETRIES = 3
DOWNLOAD_TIMEOUT = (10, 60)
DOWNLOAD_SLEEP_SEC = 2

UPLOAD_RETRIES = 3
UPLOAD_TIMEOUT = (10, 300)
UPLOAD_SLEEP_SEC = 3

# =========================
# 자동분배 우선순위
# =========================
OFFICE_ALLOC_PRIORITY = [
    "사무실 - 사무실",
    "사무실 - 스퀘어원",
    "사무실 - 구월",
    "사무실 - 부천",
    "사무실 - 푸마 여주",
    "사무실 - 아디다스 키즈",
]

STORE_ALLOC_PRIORITY = [
    "스퀘어원",
    "부천",
    "구월",
    "푸마 여주",
]

SECONDARY_DYNAMIC_STORES = [
    "스퀘어원",
    "부천",
    "구월",
]

SECONDARY_FALLBACK_STORES = [
    "푸마 여주",
]

FILA_FIXED_STORE = "휠라 파주"
RETURN_FALLBACK_STORE = "사무실 - 반품"
LEGACY_STORE_ALIASES = {
    "사무실 - 아디다스 키즈": "사무실 - S마켓",
}

ALLOC_PRIORITY = OFFICE_ALLOC_PRIORITY + STORE_ALLOC_PRIORITY
STOCK_REFERENCE_COLUMNS = ALLOC_PRIORITY + [FILA_FIXED_STORE, RETURN_FALLBACK_STORE]
VALID_STORE_VALUES = ALLOC_PRIORITY + [FILA_FIXED_STORE, RETURN_FALLBACK_STORE, "재고없음"]

STORE_SORT_ORDER = VALID_STORE_VALUES
SUMMARY_CALC_COLUMNS = ["할인가", "총 판매가", "총 수수료", "총 수입", "환율", "총 KRW", "총 공급가", "마진"]
SUMMARY_OUTPUT_COLUMNS = ["날짜", "플랫폼", "주문번호", "뒤 4자리", "브랜드", "코드", "품번", "EU 사이즈", "사이즈", "수량", "매장명"] + SUMMARY_CALC_COLUMNS + STOCK_REFERENCE_COLUMNS
DETAIL_OUTPUT_COLUMNS = [
    "주문번호",
    "브랜드",
    "최종코드",
    "품번",
    "EU 사이즈",
    "최종사이즈",
    "주문수량",
    "할인가",
    "판매가",
    "환율",
    "KRW",
    "공급가",
    "마진",
    "출고매장",
]


# =========================
# 공통 함수
# =========================
class DataValidationError(ValueError):
    pass


def make_session():
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(session, url, retries=DOWNLOAD_RETRIES, timeout=DOWNLOAD_TIMEOUT):
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            r = session.get(url, timeout=timeout)
            r.raise_for_status()
            r.encoding = "utf-8"
            return r.text
        except Exception as e:
            last_error = e
            if attempt < retries:
                time.sleep(DOWNLOAD_SLEEP_SEC)
    raise last_error


def read_google_sheet_csv_from_text(csv_text: str) -> pd.DataFrame:
    return pd.read_csv(StringIO(csv_text), dtype=str).fillna("")


def upload_to_google_sheet(df: pd.DataFrame, web_app_url: str, spreadsheet_id: str, sheet_name: str):
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()

    payload = {
        "spreadsheetId": spreadsheet_id,
        "sheetName": sheet_name,
        "values": values,
    }

    print(f"구글 시트 업로드 시작: {sheet_name} / {len(df)}행")
    last_error = None

    for attempt in range(1, UPLOAD_RETRIES + 1):
        try:
            r = requests.post(web_app_url, json=payload, timeout=UPLOAD_TIMEOUT)
            r.raise_for_status()
            print("구글 시트 업로드 완료")
            print("응답:", r.text)
            return
        except requests.exceptions.RequestException as e:
            last_error = e
            print(f"구글 시트 업로드 실패 ({sheet_name}, {attempt}/{UPLOAD_RETRIES}): {e}")
            if attempt < UPLOAD_RETRIES:
                time.sleep(UPLOAD_SLEEP_SEC)

    raise last_error


def clean_text(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).strip()
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("\\", "/")
    s = re.sub(r"\s+", " ", s)
    return s


def to_number(v, default=0):
    s = clean_text(v).replace(",", "").replace("%", "")
    if s == "":
        return default
    try:
        return float(s)
    except Exception:
        return default


def normalize_brand_series(series: pd.Series) -> pd.Series:
    s = (
        series.fillna("").astype(str).str.strip().str.upper()
        .str.replace("（", "(", regex=False)
        .str.replace("）", ")", regex=False)
        .str.replace("\\", "/", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )

    result = pd.Series("", index=series.index, dtype=str)
    result = result.mask(s.str.contains("NORTH FACE", na=False), "THE NORTH FACE")
    result = result.mask(s.str.contains("NIKE|나이키", na=False), "NIKE")
    result = result.mask(s.str.contains("ASICS|아식스", na=False), "ASICS")
    result = result.mask(s.str.contains("CONVERSE|CONBERSE|컨버스", na=False), "CONVERSE")
    result = result.mask(s.str.contains("ADIDAS", na=False), "ADIDAS")
    result = result.mask(s.str.contains("FILA", na=False), "FILA")
    result = result.mask(s.str.contains("PUMA", na=False), "PUMA")
    return result


def normalize_brand(v) -> str:
    return normalize_brand_series(pd.Series([v], dtype=str)).iloc[0]


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
        return text.replace("(黑色标)", "").replace("鞋", "")

    if brand == "FILA":
        text = strip_leading_1100(text)
        return text.replace("服", "").replace("鞋", "").replace("_", "").replace("-", "")

    if brand == "THE NORTH FACE":
        text = strip_leading_1100(text)
        return text.replace("包", "").replace("_", "").replace("-", "")

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

    bracket_match = re.search(r"\(([^()]*)\)", text)
    if bracket_match:
        before_bracket = pick_size_token(text[:bracket_match.start()])
        bracket_text = pick_size_token(bracket_match.group(1) or "")

        if is_apparel_size(before_bracket):
            text = before_bracket
        elif bracket_text and is_size_token(bracket_text):
            text = bracket_text
        else:
            text = re.sub(r"\([^()]*\)", "", text)

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


def get_suffix(value, length: int = 4) -> str:
    text = clean_text(value)
    if text == "":
        return ""
    suffix = text[-length:]
    if suffix.startswith("0"):
        return f"'{suffix}"
    return suffix


def excel_col_to_index(col: str) -> int:
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def ensure_min_columns(df_raw: pd.DataFrame, required_indices: list[int], label: str):
    if df_raw.shape[1] <= max(required_indices):
        raise DataValidationError(
            f"{label} 열 개수가 부족합니다. 필요한 최대 열 인덱스: {max(required_indices) + 1}, 실제 열 수: {df_raw.shape[1]}"
        )


# =========================
# 병렬 다운로드
# =========================
def download_all_csvs():
    urls = {
        "poizon": POIZON_INPUT_URL,
        "size_chart": SIZE_CHART_URL,
        "stock_prep": STOCK_PREP_URL,
        "exchange": EXCHANGE_RATE_URL,
        "squareone_location": SQUAREONE_LOCATION_URL,
        "fila_paju_location": FILA_PAJU_LOCATION_URL,
    }

    results = {}

    with ThreadPoolExecutor(max_workers=min(len(urls), 6)) as executor:
        future_map = {
            executor.submit(fetch_csv_text, make_session(), url): key
            for key, url in urls.items()
        }

        for future in as_completed(future_map):
            key = future_map[future]
            results[key] = future.result()

    return results


# =========================
# 기준 데이터 로드
# =========================
def load_exchange_rate(df_raw: pd.DataFrame) -> float:
    for _, row in df_raw.iterrows():
        for value in row.tolist():
            num = to_number(value, default=None)
            if num is not None and num != 0:
                return num
    return 0


def build_size_chart_lookup(df_raw: pd.DataFrame) -> dict:
    ensure_min_columns(df_raw, [0, 1, 2, 3], "사이즈표")
    size_df = df_raw.iloc[:, :4].copy()
    size_df.columns = ["브랜드", "품번", "사이즈", "EU"]
    size_df = size_df.fillna("").astype(str).apply(lambda col: col.map(clean_text))

    lookup = {}
    for brand, product_code, size_value, eu_size in size_df.itertuples(index=False, name=None):
        brand_key = normalize_lookup_key(brand)
        product_code_key = normalize_lookup_key(product_code)
        eu_key = normalize_lookup_key(eu_size)
        if brand_key and eu_key and size_value:
            lookup.setdefault((brand_key, product_code_key, eu_key), clean_text(size_value))
    return lookup


def apply_size_chart_lookup(brand: str, product_code: str, converted_size: str, size_lookup: dict) -> str:
    brand_key = normalize_lookup_key(brand)
    product_code_key = normalize_lookup_key(product_code)
    if brand_key == "THE NORTH FACE":
        product_code_key = product_code_key[:7]

    key = (brand_key, product_code_key, normalize_lookup_key(converted_size))
    if brand_key == "THE NORTH FACE":
        return size_lookup.get(key, converted_size)

    fallback_key = (brand_key, "", normalize_lookup_key(converted_size))
    return size_lookup.get(key, size_lookup.get(fallback_key, converted_size))


def load_stock_prepare(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()

    required = ["코드", "품번", "사이즈", "공급가"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"분배준비 시트에 필수 열이 없습니다: {col}")

    for col in ["변환코드1", "변환코드2", "변환코드3", "EU"]:
        if col not in df.columns:
            df[col] = ""

    for col in STOCK_REFERENCE_COLUMNS:
        legacy_col = LEGACY_STORE_ALIASES.get(col)
        if col not in df.columns and legacy_col in df.columns:
            df[col] = df[legacy_col]
        if col not in df.columns:
            df[col] = 0

    keep_cols = required + ["변환코드1", "변환코드2", "변환코드3", "EU"] + STOCK_REFERENCE_COLUMNS
    if "브랜드" in df.columns:
        keep_cols.append("브랜드")
    if "품번_변환" in df.columns:
        keep_cols.append("품번_변환")
    if "할인가" in df.columns:
        keep_cols.append("할인가")

    df = df[keep_cols].copy()

    text_cols = ["코드", "품번", "사이즈", "변환코드1", "변환코드2", "변환코드3", "EU"]
    if "브랜드" in df.columns:
        text_cols.append("브랜드")
    if "품번_변환" in df.columns:
        text_cols.append("품번_변환")
    for col in text_cols:
        df[col] = (
            df[col].fillna("").astype(str).str.strip()
            .str.replace("（", "(", regex=False)
            .str.replace("）", ")", regex=False)
            .str.replace("\\", "/", regex=False)
            .str.replace(r"\s+", " ", regex=True)
        )

    df["공급가"] = df["공급가"].apply(to_number)

    if "할인가" in df.columns:
        df["할인가"] = pd.to_numeric(df["할인가"], errors="coerce")

    for col in STOCK_REFERENCE_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df


def load_selected_excel_cols(df_raw: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    use_cols = [excel_col_to_index(col) for col in cols]
    ensure_min_columns(df_raw, use_cols, f"선택 열 {cols}")
    df = df_raw.iloc[:, use_cols].copy().fillna("")
    df.columns = cols
    return df


def build_supply_price_map(stock_df: pd.DataFrame) -> dict:
    temp = stock_df[["품번", "공급가"]].copy()
    temp["품번"] = temp["품번"].apply(normalize_product_lookup_key)
    temp = temp.drop_duplicates(subset=["품번"], keep="first")
    return dict(zip(temp["품번"], temp["공급가"]))


def build_sale_price_map(stock_df: pd.DataFrame) -> dict:
    if "할인가" not in stock_df.columns:
        return {}

    temp = stock_df[["품번", "할인가"]].copy()
    temp["품번"] = temp["품번"].apply(normalize_product_lookup_key)
    temp["할인가"] = pd.to_numeric(temp["할인가"], errors="coerce")
    temp = temp.dropna(subset=["할인가"])
    temp = temp.drop_duplicates(subset=["품번"], keep="first")
    return dict(zip(temp["품번"], temp["할인가"]))


def build_store_stock_lookup(stock_df: pd.DataFrame) -> pd.DataFrame:
    keep_cols = ["코드"] + STOCK_REFERENCE_COLUMNS
    temp = stock_df[keep_cols].copy()
    temp["코드"] = temp["코드"].apply(clean_text)
    temp = temp.drop_duplicates(subset=["코드"], keep="first")
    return temp


def build_stock_size_lookup(stock_df: pd.DataFrame) -> dict:
    lookup = {}
    code_columns = ["코드", "변환코드1", "변환코드2", "변환코드3"]
    for row in stock_df[["사이즈"] + code_columns].itertuples(index=False, name=None):
        stock_size = clean_text(row[0])
        if not stock_size:
            continue
        for code in row[1:]:
            code_key = normalize_code_lookup_key(code)
            if code_key:
                lookup.setdefault(code_key, stock_size)
    return lookup


def build_stock_brand_lookup(stock_df: pd.DataFrame) -> dict:
    if "브랜드" not in stock_df.columns:
        return {}

    temp = stock_df.copy()
    temp["브랜드"] = normalize_brand_series(temp["브랜드"])
    temp = temp[temp["브랜드"].ne("")]

    lookup = {}
    for _, row in temp.iterrows():
        brand = clean_text(row.get("브랜드", ""))
        raw_product_code = clean_text(row.get("품번", ""))
        converted_product_code = clean_text(row.get("품번_변환", ""))
        for product_code in [
            raw_product_code,
            converted_product_code,
            normalize_product_lookup_key(raw_product_code),
            normalize_product_lookup_key(converted_product_code),
            normalize_product_code(raw_product_code, brand),
        ]:
            product_code_key = clean_text(product_code)
            if product_code_key:
                lookup.setdefault(product_code_key, brand)
    return lookup


def lookup_stock_brand(product_code: str, stock_brand_lookup: dict) -> str:
    text = normalize_product_lookup_key(product_code)
    if text in stock_brand_lookup:
        return stock_brand_lookup[text]
    stripped_text = strip_leading_1100(text)
    if stripped_text in stock_brand_lookup:
        return stock_brand_lookup[stripped_text]
    return ""


# =========================
# 자동분배
# =========================
def allocate_round_robin(
    stock_entry: dict,
    stores: list[str],
    remain: int,
    start_index: int = 0,
) -> tuple[list[tuple[str, int]], int]:
    allocations = []
    if not stores:
        return allocations, 0

    store_count = len(stores)
    current_index = start_index % store_count

    while remain > 0:
        progressed = False
        checked = 0

        while checked < store_count and remain > 0:
            store = stores[current_index]
            current_index = (current_index + 1) % store_count
            checked += 1
            available = int(stock_entry.get(store, 0))
            if available <= 0:
                continue

            stock_entry[store] = available - 1
            remain -= 1
            progressed = True
            allocations.append((store, 1))

        if not progressed:
            break

    return allocations, current_index


def allocate_by_priority(
    stock_entry: dict,
    stores: list[str],
    remain: int,
) -> list[tuple[str, int]]:
    allocations = []
    if remain <= 0:
        return allocations

    for store in stores:
        if remain <= 0:
            break

        available = int(stock_entry.get(store, 0))
        if available <= 0:
            continue

        use_qty = min(available, remain)
        stock_entry[store] = available - use_qty
        remain -= use_qty
        allocations.append((store, use_qty))

    return allocations


def allocate_secondary_stores(
    stock_entry: dict,
    remain: int,
    start_index: int = 0,
) -> tuple[list[tuple[str, int]], int]:
    allocations = []
    priority_order = {store: idx for idx, store in enumerate(SECONDARY_DYNAMIC_STORES)}
    dynamic_stores = sorted(
        SECONDARY_DYNAMIC_STORES,
        key=lambda store: (
            -int(stock_entry.get(store, 0)),
            priority_order[store],
        ),
    )

    dynamic_allocations = allocate_by_priority(
        stock_entry,
        dynamic_stores,
        remain,
    )
    allocations.extend(dynamic_allocations)
    remain -= sum(qty for _, qty in dynamic_allocations)

    if remain > 0:
        allocations.extend(
            allocate_by_priority(stock_entry, SECONDARY_FALLBACK_STORES, remain)
        )

    return allocations, start_index


def auto_allocate(order_df: pd.DataFrame, stock_df: pd.DataFrame) -> pd.DataFrame:
    alloc_rows = []
    office_rr_state = {}
    store_rr_state = {}
    stock_work = (
        stock_df[["코드"] + STOCK_REFERENCE_COLUMNS]
        .drop_duplicates(subset=["코드"], keep="first")
        .set_index("코드")[STOCK_REFERENCE_COLUMNS]
        .to_dict("index")
    )
    stock_index = set(stock_work.keys())

    for row in order_df.itertuples(index=False):
        source_row_id = int(row.원본순번)
        order_no = row.주문번호
        brand = row.브랜드
        item_no = row.품번
        final_size = row.출력사이즈
        shoe_size = row.사이즈2
        final_code = row.최종코드
        match_code = row.매칭코드
        requested_qty = int(row.수량)
        price = row.판매가

        if match_code == "" or match_code not in stock_index:
            target_store = "재고없음"
            alloc_rows.append({
                "원본순번": source_row_id,
                "주문번호": order_no,
                "브랜드": brand,
                "최종코드": final_code,
                "매칭코드": match_code,
                "품번": item_no,
                "최종사이즈": final_size,
                "신발 사이즈": shoe_size,
                "주문수량": requested_qty,
                "판매가": price,
                "출고매장": target_store,
            })
            continue

        remain = requested_qty
        stock_entry = stock_work[match_code]
        rr_start_index = office_rr_state.get(match_code, 0)

        rr_allocations, next_rr_index = allocate_round_robin(
            stock_entry,
            OFFICE_ALLOC_PRIORITY,
            remain,
            rr_start_index,
        )
        office_rr_state[match_code] = next_rr_index
        for store, use_qty in rr_allocations:
            alloc_rows.append({
                "원본순번": source_row_id,
                "주문번호": order_no,
                "브랜드": brand,
                "최종코드": final_code,
                "매칭코드": match_code,
                "품번": item_no,
                "최종사이즈": final_size,
                "신발 사이즈": shoe_size,
                "주문수량": use_qty,
                "판매가": price,
                "출고매장": store,
            })
        remain -= sum(qty for _, qty in rr_allocations)

        if remain > 0 and brand == "FILA":
            target_store = FILA_FIXED_STORE
            available = int(stock_entry.get(target_store, 0))
            use_qty = min(available, remain)
            if use_qty > 0:
                stock_entry[target_store] = available - use_qty
                alloc_rows.append({
                    "원본순번": source_row_id,
                    "주문번호": order_no,
                    "브랜드": brand,
                    "최종코드": final_code,
                    "매칭코드": match_code,
                    "품번": item_no,
                    "최종사이즈": final_size,
                    "신발 사이즈": shoe_size,
                    "주문수량": use_qty,
                    "판매가": price,
                    "출고매장": target_store,
                })
                remain -= use_qty

        if remain > 0:
            store_start_index = store_rr_state.get(match_code, 0)
            store_allocations, next_store_index = allocate_secondary_stores(
                stock_entry,
                remain,
                store_start_index,
            )
            store_rr_state[match_code] = next_store_index
            for store, use_qty in store_allocations:
                alloc_rows.append({
                    "원본순번": source_row_id,
                    "주문번호": order_no,
                    "브랜드": brand,
                    "최종코드": final_code,
                    "매칭코드": match_code,
                    "품번": item_no,
                    "최종사이즈": final_size,
                    "신발 사이즈": shoe_size,
                    "주문수량": use_qty,
                    "판매가": price,
                    "출고매장": store,
                })
            remain -= sum(qty for _, qty in store_allocations)

        if remain > 0:
            return_allocations = allocate_by_priority(
                stock_entry,
                [RETURN_FALLBACK_STORE],
                remain,
            )
            for store, use_qty in return_allocations:
                alloc_rows.append({
                    "원본순번": source_row_id,
                    "주문번호": order_no,
                    "브랜드": brand,
                    "최종코드": final_code,
                    "매칭코드": match_code,
                    "품번": item_no,
                    "최종사이즈": final_size,
                    "신발 사이즈": shoe_size,
                    "주문수량": use_qty,
                    "판매가": price,
                    "출고매장": store,
                })
            remain -= sum(qty for _, qty in return_allocations)

        if remain > 0:
            alloc_rows.append({
                "원본순번": source_row_id,
                "주문번호": order_no,
                "브랜드": brand,
                "최종코드": final_code,
                "매칭코드": match_code,
                "품번": item_no,
                "최종사이즈": final_size,
                "신발 사이즈": shoe_size,
                "주문수량": remain,
                "판매가": price,
                "출고매장": "재고없음",
            })

    return pd.DataFrame(alloc_rows)


def sort_alloc_df_by_store(df: pd.DataFrame) -> pd.DataFrame:
    temp = df.copy()
    temp["출고매장_순서"] = temp["출고매장"].apply(
        lambda x: STORE_SORT_ORDER.index(x) if x in STORE_SORT_ORDER else 999
    )

    temp = temp.sort_values(
        by=["출고매장_순서", "출고매장", "브랜드", "품번", "최종사이즈", "주문번호"],
        ascending=[True, True, True, True, True, True]
    ).drop(columns=["출고매장_순서"]).reset_index(drop=True)

    return temp


def build_summary_metrics(order_df: pd.DataFrame, stock_df: pd.DataFrame, exchange_raw: pd.DataFrame) -> pd.DataFrame:
    temp = order_df.copy()
    temp["원본순번"] = pd.to_numeric(temp["원본순번"], errors="coerce").fillna(0).astype(int)
    if "출력사이즈" in temp.columns and "최종사이즈" not in temp.columns:
        temp = temp.rename(columns={"출력사이즈": "최종사이즈"})
    if "최종코드" in temp.columns and "코드" not in temp.columns:
        temp = temp.rename(columns={"최종코드": "코드"})
    if "사이즈2" in temp.columns and "EU 사이즈" not in temp.columns:
        temp = temp.rename(columns={"사이즈2": "EU 사이즈"})
    temp["수량"] = pd.to_numeric(temp["수량"], errors="coerce").fillna(0).astype(int)
    temp["판매가"] = pd.to_numeric(temp["판매가"], errors="coerce").fillna(0)
    temp["최초 판매가"] = pd.to_numeric(temp["최초 판매가"], errors="coerce").fillna(0)
    temp["총 판매가"] = temp["수량"] * temp["최초 판매가"]
    temp["총 수수료"] = temp["수량"] * (temp["최초 판매가"] - temp["판매가"])
    temp["총 수입"] = temp["수량"] * temp["판매가"]
    temp["코드"] = temp["코드"].apply(clean_text)
    temp["사이즈"] = temp["최종사이즈"].apply(clean_text)
    if "EU 사이즈" not in temp.columns:
        temp["EU 사이즈"] = ""
    temp["EU 사이즈"] = temp["EU 사이즈"].apply(clean_text)
    temp["단위판매가"] = temp["최초 판매가"].round(0)
    temp["단위수입"] = temp["판매가"].round(0)
    temp["단위수수료"] = (temp["최초 판매가"] - temp["판매가"]).round(0)

    exchange_rate = load_exchange_rate(exchange_raw)
    sale_price_map = build_sale_price_map(stock_df)
    supply_price_map = build_supply_price_map(stock_df)
    temp["할인가"] = temp["품번"].map(sale_price_map).fillna(0)
    temp["할인가"] = pd.to_numeric(temp["할인가"], errors="coerce").fillna(0)
    temp["환율"] = exchange_rate
    temp["KRW"] = (temp["판매가"] * temp["환율"]).round(0)
    temp["공급가"] = temp["품번"].map(supply_price_map).fillna(0)
    temp["공급가"] = pd.to_numeric(temp["공급가"], errors="coerce").fillna(0).div(1.1)
    temp["총 KRW"] = (temp["KRW"] * temp["수량"]).round(0)
    temp["총 공급가"] = (temp["공급가"] * temp["수량"]).round(0)
    temp["마진"] = (temp["총 KRW"] - temp["총 공급가"]).round(0)

    temp["할인가"] = pd.to_numeric(temp["할인가"], errors="coerce").fillna(0).round(0).astype(int)
    temp["판매가"] = pd.to_numeric(temp["판매가"], errors="coerce").fillna(0).round(0).astype(int)
    temp["최초 판매가"] = pd.to_numeric(temp["최초 판매가"], errors="coerce").fillna(0).round(0).astype(int)
    temp["총 판매가"] = pd.to_numeric(temp["총 판매가"], errors="coerce").fillna(0).round(0).astype(int)
    temp["총 수입"] = pd.to_numeric(temp["총 수입"], errors="coerce").fillna(0).round(0).astype(int)
    temp["총 수수료"] = pd.to_numeric(temp["총 수수료"], errors="coerce").fillna(0).round(0).astype(int)
    temp["단위판매가"] = pd.to_numeric(temp["단위판매가"], errors="coerce").fillna(0).round(0).astype(int)
    temp["단위수입"] = pd.to_numeric(temp["단위수입"], errors="coerce").fillna(0).round(0).astype(int)
    temp["단위수수료"] = pd.to_numeric(temp["단위수수료"], errors="coerce").fillna(0).round(0).astype(int)
    temp["환율"] = pd.to_numeric(temp["환율"], errors="coerce").fillna(0)
    temp["KRW"] = pd.to_numeric(temp["KRW"], errors="coerce").fillna(0).round(0).astype(int)
    temp["총 KRW"] = pd.to_numeric(temp["총 KRW"], errors="coerce").fillna(0).round(0).astype(int)
    temp["공급가"] = pd.to_numeric(temp["공급가"], errors="coerce").fillna(0).round(0).astype(int)
    temp["총 공급가"] = pd.to_numeric(temp["총 공급가"], errors="coerce").fillna(0).round(0).astype(int)
    temp["마진"] = pd.to_numeric(temp["마진"], errors="coerce").fillna(0).round(0).astype(int)
    internal_columns = ["KRW", "공급가", "단위판매가", "단위수입", "단위수수료"]
    return temp[["원본순번", "주문번호", "브랜드", "코드", "품번", "EU 사이즈", "사이즈"] + SUMMARY_CALC_COLUMNS + internal_columns].copy()


def build_store_pick_lists(df: pd.DataFrame, stock_df: pd.DataFrame, order_df: pd.DataFrame, exchange_raw: pd.DataFrame) -> pd.DataFrame:
    temp = df[df["출고매장"].isin(VALID_STORE_VALUES)].copy()
    temp["원본순번"] = pd.to_numeric(temp["원본순번"], errors="coerce").fillna(0).astype(int)

    stock_lookup_df = build_store_stock_lookup(stock_df)
    summary_metrics_df = build_summary_metrics(order_df, stock_df, exchange_raw)

    if len(temp) == 0:
        return pd.DataFrame(columns=SUMMARY_OUTPUT_COLUMNS)

    pick_df = temp.copy()

    pick_df["출고매장_순서"] = pick_df["출고매장"].apply(
        lambda x: STORE_SORT_ORDER.index(x) if x in STORE_SORT_ORDER else 999
    )

    pick_df = (
        pick_df.sort_values(
            by=["출고매장_순서", "출고매장", "브랜드", "품번", "최종사이즈", "주문번호", "최종코드"],
            ascending=[True, True, True, True, True, True, True]
        )
        .drop(columns=["출고매장_순서"])
        .reset_index(drop=True)
    )

    pick_df = pick_df.rename(columns={
        "최종코드": "코드",
        "최종사이즈": "사이즈",
        "신발 사이즈": "EU 사이즈",
        "주문수량": "수량",
        "출고매장": "매장명",
    })
    if "EU 사이즈" not in pick_df.columns:
        pick_df["EU 사이즈"] = ""

    pick_df = pick_df.merge(
        stock_lookup_df,
        left_on="매칭코드",
        right_on="코드",
        how="left",
        suffixes=("", "_재고"),
    )
    if "코드_재고" in pick_df.columns:
        pick_df = pick_df.drop(columns=["코드_재고"])
    pick_df = pick_df.merge(
        summary_metrics_df,
        on=["원본순번", "주문번호", "브랜드", "코드", "품번", "EU 사이즈", "사이즈"],
        how="left",
    )

    for col in STOCK_REFERENCE_COLUMNS:
        if col not in pick_df.columns:
            pick_df[col] = 0
        pick_df[col] = pd.to_numeric(pick_df[col], errors="coerce").fillna(0).astype(int)
        pick_df[col] = pick_df[col].mask(pick_df[col] == 0, "").astype(object)

    for col in SUMMARY_CALC_COLUMNS:
        if col not in pick_df.columns:
            pick_df[col] = 0

    today_text = datetime.now().strftime("%Y-%m-%d")
    pick_df["날짜"] = today_text
    pick_df["플랫폼"] = "POIZON"
    pick_df["주문번호"] = pick_df["주문번호"].apply(clean_text)
    pick_df["뒤 4자리"] = ""
    pick_df["수량"] = pd.to_numeric(pick_df["수량"], errors="coerce").fillna(0).astype(int)
    pick_df["할인가"] = pd.to_numeric(pick_df["할인가"], errors="coerce").fillna(0).astype(int)
    pick_df["단위판매가"] = pd.to_numeric(pick_df["단위판매가"], errors="coerce").fillna(0)
    pick_df["단위수입"] = pd.to_numeric(pick_df["단위수입"], errors="coerce").fillna(0)
    pick_df["단위수수료"] = pd.to_numeric(pick_df["단위수수료"], errors="coerce").fillna(0)
    pick_df["환율"] = pd.to_numeric(pick_df["환율"], errors="coerce").fillna(0)
    pick_df["KRW"] = pd.to_numeric(pick_df["KRW"], errors="coerce").fillna(0)
    pick_df["공급가"] = pd.to_numeric(pick_df["공급가"], errors="coerce").fillna(0)
    pick_df["총 판매가"] = (pick_df["단위판매가"] * pick_df["수량"]).round(0).astype(int)
    pick_df["총 수입"] = (pick_df["단위수입"] * pick_df["수량"]).round(0).astype(int)
    pick_df["총 수수료"] = (pick_df["단위수수료"] * pick_df["수량"]).round(0).astype(int)
    pick_df["총 KRW"] = (pick_df["KRW"] * pick_df["수량"]).round(0).astype(int)
    pick_df["총 공급가"] = (pick_df["공급가"] * pick_df["수량"]).round(0).astype(int)
    pick_df["마진"] = (pick_df["총 KRW"] - pick_df["총 공급가"]).round(0).astype(int)
    return pick_df[SUMMARY_OUTPUT_COLUMNS].copy()


def prepare_order_source(src: pd.DataFrame, size_lookup: dict, stock_size_lookup: dict, stock_brand_lookup: dict) -> pd.DataFrame:
    use_cols = [
        excel_col_to_index("B"),
        excel_col_to_index("E"),
        excel_col_to_index("F"),
        excel_col_to_index("L"),
        excel_col_to_index("U"),
        excel_col_to_index("AI"),
        excel_col_to_index("AO"),
    ]

    ensure_min_columns(src, use_cols, "POIZON 주문 시트")
    src = src.iloc[:, use_cols].copy()
    src.columns = ["주문번호", "브랜드원본", "품번원본", "사이즈원본", "수량", "최초 판매가", "판매가"]
    src.insert(0, "원본순번", range(1, len(src) + 1))
    src = src.fillna("").astype(str)
    src = src.apply(lambda col: col.str.strip())

    src["최초 판매가"] = src["최초 판매가"].str.replace("¥", "", regex=False)
    src["최초 판매가"] = src["최초 판매가"].str.replace(",", "", regex=False)
    src["판매가"] = src["판매가"].str.replace("¥", "", regex=False)
    src["판매가"] = src["판매가"].str.replace(",", "", regex=False)
    src["수량"] = pd.to_numeric(src["수량"], errors="coerce").fillna(0).astype(int)
    src["최초 판매가"] = pd.to_numeric(src["최초 판매가"], errors="coerce").fillna(0)
    src["판매가"] = pd.to_numeric(src["판매가"], errors="coerce").fillna(0)
    if src["주문번호"].eq("").all():
        raise DataValidationError("POIZON 주문 시트의 주문번호가 모두 비어 있습니다.")

    source_brands = normalize_brand_series(src["브랜드원본"])
    stock_brands = [
        lookup_stock_brand(product_code, stock_brand_lookup)
        for product_code in src["품번원본"]
    ]
    src["브랜드"] = [
        stock_brand if clean_text(stock_brand) else source_brand
        for stock_brand, source_brand in zip(stock_brands, source_brands)
    ]
    src["품번"] = [
        normalize_product_code(product_code, brand)
        for product_code, brand in zip(src["품번원본"], src["브랜드"])
    ]
    converted_sizes = [
        normalize_size_value(size_value, brand)
        for size_value, brand in zip(src["사이즈원본"], src["브랜드"])
    ]
    src["사이즈2"] = converted_sizes
    src["사이즈변환"] = [
        apply_size_chart_lookup(brand, product_code, converted_size, size_lookup)
        for brand, product_code, converted_size in zip(src["브랜드"], src["품번"], converted_sizes)
    ]
    base_codes = [
        make_output_code(product_code, converted_size)
        for product_code, converted_size in zip(src["품번"], src["사이즈변환"])
    ]
    stock_size_matches = [
        stock_size_lookup.get(normalize_code_lookup_key(code), "")
        for code in base_codes
    ]
    src["출력사이즈"] = [
        stock_size if clean_text(stock_size) else converted_size
        for stock_size, converted_size in zip(stock_size_matches, src["사이즈변환"])
    ]
    src["최종코드"] = [
        make_converted_output_code(product_code, stock_size, base_code)
        for product_code, stock_size, base_code in zip(src["품번"], stock_size_matches, base_codes)
    ]
    src["매칭코드"] = src["최종코드"]

    return src[
        ["원본순번", "주문번호", "브랜드", "품번", "사이즈2", "출력사이즈", "최종코드", "매칭코드", "수량", "최초 판매가", "판매가"]
    ].copy()


def finalize_alloc_df(alloc_df: pd.DataFrame, stock_df: pd.DataFrame, exchange_raw: pd.DataFrame) -> pd.DataFrame:
    exchange_rate = load_exchange_rate(exchange_raw)
    supply_price_map = build_supply_price_map(stock_df)
    sale_price_map = build_sale_price_map(stock_df)

    alloc_df = alloc_df.copy()
    alloc_df["주문수량"] = pd.to_numeric(alloc_df["주문수량"], errors="coerce").fillna(0).astype(int)
    alloc_df["판매가"] = pd.to_numeric(alloc_df["판매가"], errors="coerce").fillna(0)
    alloc_df["환율"] = exchange_rate
    alloc_df["KRW"] = (alloc_df["판매가"].astype(float) * exchange_rate).round(0)
    alloc_df["할인가"] = alloc_df["품번"].map(sale_price_map)
    alloc_df["할인가"] = pd.to_numeric(alloc_df["할인가"], errors="coerce").fillna(0)
    alloc_df["공급가"] = alloc_df["품번"].map(supply_price_map).fillna(0)
    alloc_df["공급가"] = pd.to_numeric(alloc_df["공급가"], errors="coerce").fillna(0).div(1.1)
    alloc_df["마진"] = (alloc_df["KRW"] - alloc_df["공급가"]).round(0)
    alloc_df["총 할인가"] = alloc_df["할인가"] * alloc_df["주문수량"]
    alloc_df["총 판매가"] = alloc_df["판매가"] * alloc_df["주문수량"]
    alloc_df["총 KRW"] = alloc_df["KRW"] * alloc_df["주문수량"]
    alloc_df["총 공급가"] = alloc_df["공급가"] * alloc_df["주문수량"]
    alloc_df["총 마진"] = alloc_df["마진"] * alloc_df["주문수량"]

    for col in ["할인가", "총 할인가", "판매가", "총 판매가", "환율", "KRW", "총 KRW", "공급가", "총 공급가", "마진", "총 마진"]:
        alloc_df[col] = alloc_df[col].round(0).astype(int)

    alloc_df = alloc_df[DETAIL_OUTPUT_COLUMNS].copy()
    return sort_alloc_df_by_store(alloc_df)


# =========================
# 메인
# =========================
def main():
    start_time = time.perf_counter()
    print("CSV 병렬 다운로드 중...")
    csvs = download_all_csvs()

    src = read_google_sheet_csv_from_text(csvs["poizon"])
    size_chart_raw = read_google_sheet_csv_from_text(csvs["size_chart"])
    stock_prep_raw = read_google_sheet_csv_from_text(csvs["stock_prep"])
    exchange_raw = read_google_sheet_csv_from_text(csvs["exchange"])
    squareone_location_df = read_google_sheet_csv_from_text(csvs["squareone_location"])
    fila_paju_location_raw = read_google_sheet_csv_from_text(csvs["fila_paju_location"])
    fila_paju_location_df = load_selected_excel_cols(fila_paju_location_raw, ["E", "F"])

    stock_df = load_stock_prepare(stock_prep_raw)
    size_lookup = build_size_chart_lookup(size_chart_raw)
    stock_size_lookup = build_stock_size_lookup(stock_df)
    stock_brand_lookup = build_stock_brand_lookup(stock_df)
    order_df = prepare_order_source(src, size_lookup, stock_size_lookup, stock_brand_lookup)
    alloc_df = auto_allocate(order_df, stock_df)
    pick_df = build_store_pick_lists(alloc_df, stock_df, order_df, exchange_raw)

    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = [
            executor.submit(
                upload_to_google_sheet,
                pick_df,
                SUMMARY_WEB_APP_URL,
                SUMMARY_SPREADSHEET_ID,
                SUMMARY_SHEET_NAME,
            ),
            executor.submit(
                upload_to_google_sheet,
                squareone_location_df,
                SUMMARY_WEB_APP_URL,
                SUMMARY_SPREADSHEET_ID,
                SQUAREONE_LOCATION_SHEET_NAME,
            ),
            executor.submit(
                upload_to_google_sheet,
                fila_paju_location_df,
                SUMMARY_WEB_APP_URL,
                SUMMARY_SPREADSHEET_ID,
                FILA_PAJU_LOCATION_SHEET_NAME,
            ),
        ]
        for future in as_completed(futures):
            future.result()

    elapsed = time.perf_counter() - start_time
    print(f"완료 (총 소요 시간: {elapsed:.2f}초)")


if __name__ == "__main__":
    main()
