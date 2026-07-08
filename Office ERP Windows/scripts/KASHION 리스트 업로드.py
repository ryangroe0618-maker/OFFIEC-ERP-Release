# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor, as_completed
from io import StringIO
import re
import time
from datetime import datetime

import pandas as pd
import requests


SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTU9VO_90AH04WGk0CpddbLa5JHYZg4kiiG0UaHWMaGw82-hVcf5s-XcJtSGeErqMJYP22uy5jqdQ7a/pub?gid=0&single=true&output=csv"
STOCK_TRANSFORM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"
ONLINE_SIZE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=2089368433&single=true&output=csv"
EXCHANGE_RATE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=295228098&single=true&output=csv"
PLATFORM_INFO_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRWhzq2rt_O-YnOwzWnE8L_d8--NBu-EMmDkxnnVy-oATAfCjDIX976b973bY2MiO8gYPTE5WCk-tVv/pub?gid=733480714&single=true&output=csv"
SQUAREONE_LOCATION_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSZ4Mgu9j6y27nLBYU8gAhDTfy4eMpvBgvs3oorR3BUCpcgoyf6Z1SllaqsFyos8LcH5DfxoUsN4NYG/pub?gid=289091756&single=true&output=csv"
FILA_PAJU_LOCATION_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSA47SgFq9QQPg0D3AlBnpJX6q7Yx_Dh66E1ID9MlXTahJjL0FmFVtPgyTEtj4iVj7PvRkCUoCgbjkd/pub?gid=1813802704&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxzBOrRpU-zDYT_p3vf1d4oPf1Bv5QZ8cvXJW5HAqqN1p9IsnSERlHMnXi609iIampX/exec"
SPREADSHEET_ID = "1Vm5Nxs76ELKyk7QFK_9ohRT2eFpXlVUseEhY5qyCE2A"
SHEET_NAME = "OUT"
SEARCHING_SHEET_NAME = "구하는 중"
SQUAREONE_LOCATION_SHEET_NAME = "스퀘어원 제품 위치"
FILA_PAJU_LOCATION_SHEET_NAME = "휠라 파주 제품 위치"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 120)
RETRIES = 3
RETRY_SLEEP_SEC = 2

OFFICE_ALLOC_PRIORITY = [
    "사무실 - 사무실",
    "사무실 - 스퀘어원",
    "사무실 - 구월",
    "사무실 - 부천",
    "사무실 - 푸마 여주",
    "사무실 - 아디다스 키즈",
]

DYNAMIC_STORE_PRIORITY = [
    "스퀘어원",
    "부천",
    "구월",
]

STATIC_FALLBACK_STORES = [
    "휠라 파주",
    "푸마 여주",
]
RETURN_FALLBACK_STORE = "사무실 - 반품"
LEGACY_STORE_ALIASES = {
    "사무실 - 아디다스 키즈": "사무실 - S마켓",
}

ALLOC_PRIORITY = OFFICE_ALLOC_PRIORITY + DYNAMIC_STORE_PRIORITY + STATIC_FALLBACK_STORES + [RETURN_FALLBACK_STORE]
STORE_SORT_ORDER = ALLOC_PRIORITY + ["재고없음"]
OUTPUT_COLUMNS = ["날짜", "원거래 날짜", "플랫폼", "주문번호", "뒤 4자리", "브랜드", "코드", "품번", "사이즈", "수량", "매장명", "할인가", "판매가", "수수료", "수입", "환율", "KRW", "공급가", "마진", "내역"] + ALLOC_PRIORITY
STORE_ORDER_INDEX = {store: idx for idx, store in enumerate(STORE_SORT_ORDER)}
EMPTY_STOCK_ENTRY = {col: 0 for col in ALLOC_PRIORITY}


def make_session():
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def today_str() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def fetch_csv_text(session: requests.Session, url: str) -> str:
    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            response = session.get(url, timeout=DOWNLOAD_TIMEOUT)
            response.raise_for_status()
            response.encoding = "utf-8"
            return response.text
        except Exception as e:
            last_error = e
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def fetch_all_csv_texts(url_map: dict[str, str]) -> dict[str, str]:
    results = {}
    with ThreadPoolExecutor(max_workers=min(8, len(url_map))) as executor:
        future_map = {
            executor.submit(fetch_csv_text, make_session(), url): name
            for name, url in url_map.items()
        }
        for future in as_completed(future_map):
            results[future_map[future]] = future.result()
    return results


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("\\", "/")
    text = " ".join(text.split())
    return text


def compact_code_text(value) -> str:
    return clean_text(value).replace("-", "")


def code_lookup_keys(value) -> list[str]:
    keys = []
    for key in [
        clean_text(value),
        normalize_key(value),
        compact_code_text(value),
        normalize_key(compact_code_text(value)),
    ]:
        if key and key not in keys:
            keys.append(key)
    return keys


def normalize_size(value) -> str:
    text = clean_text(value)
    if ":" in text:
        text = text.split(":", 1)[1]
    return text.strip()


def size_step1(value) -> str:
    text = clean_text(value)
    if text == "":
        return ""

    if " " in text:
        text = text.rsplit(" ", 1)[-1]

    if "/" in text:
        text = text.split("/", 1)[0]

    return text


def size_step2(value) -> str:
    text = clean_text(value)
    if text == "":
        return ""

    if re.search(r"[\u4e00-\u9fff]", text):
        return "ONE"

    return text


def normalize_key(value) -> str:
    return clean_text(value).upper()


def normalize_platform(value) -> str:
    platform = clean_text(value)
    if platform == "天猫":
        return "TM"
    if platform == "京东":
        return "JD"
    return platform


def get_order_suffix(value, length: int = 4) -> str:
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
        raise ValueError(
            f"{label} 열 개수가 부족합니다. 필요한 최대 열 인덱스: {max(required_indices) + 1}, 실제 열 수: {df_raw.shape[1]}"
        )


def load_selected_excel_cols(df_raw: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    use_cols = [excel_col_to_index(col) for col in cols]
    ensure_min_columns(df_raw, use_cols, f"선택 열 {cols}")
    df = df_raw.iloc[:, use_cols].copy().fillna("")
    df.columns = cols
    return df


def to_int(value) -> int:
    if value is None or pd.isna(value):
        return 0
    text = str(value).replace(",", "").strip()
    if text == "":
        return 0
    try:
        return int(float(text))
    except Exception:
        return 0


def to_float(value) -> float:
    if value is None or pd.isna(value):
        return 0.0
    text = str(value).replace(",", "").strip()
    if text == "":
        return 0.0
    try:
        return float(text)
    except Exception:
        return 0.0


def format_number(value):
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return ""
    num = float(num)
    if num.is_integer():
        return str(int(num))
    return f"{num:.4f}".rstrip("0").rstrip(".")


def make_empty_stock_entry() -> dict:
    return EMPTY_STOCK_ENTRY.copy()


def get_stock_entry(stock_lookup: dict, code: str) -> dict:
    for key in code_lookup_keys(code):
        if key in stock_lookup:
            return stock_lookup[key]
    return stock_lookup.setdefault(clean_text(code), make_empty_stock_entry())


def make_stock_lookups(stock_df: pd.DataFrame):
    transform_lookup = {}
    stock_lookup = {}
    supply_lookup = {}
    discount_lookup = {}

    temp = stock_df.fillna("")
    columns = list(temp.columns)

    for values in temp.itertuples(index=False, name=None):
        row = dict(zip(columns, values))
        item_no = clean_text(row.get("품번", ""))
        converted_item_no = clean_text(row.get("품번_변환", ""))
        code = clean_text(row.get("코드", ""))

        if converted_item_no:
            transform_lookup[normalize_key(converted_item_no)] = {
                "품번_변환": item_no or converted_item_no,
                "브랜드": clean_text(row.get("브랜드", "")),
            }
            compact_converted_item_no = compact_code_text(converted_item_no)
            if compact_converted_item_no:
                transform_lookup.setdefault(
                    compact_converted_item_no.upper(),
                    {
                        "품번_변환": item_no or converted_item_no,
                        "브랜드": clean_text(row.get("브랜드", "")),
                    },
                )

        if code and not any(key in stock_lookup for key in code_lookup_keys(code)):
            stock_entry = {
                col: to_int(row.get(col, row.get(LEGACY_STORE_ALIASES.get(col, ""), 0)))
                for col in ALLOC_PRIORITY
            }
            for key in code_lookup_keys(code):
                stock_lookup.setdefault(key, stock_entry)

        if item_no and item_no not in supply_lookup:
            supply_lookup[item_no] = to_float(row.get("공급가", 0))
        compact_item_no = compact_code_text(item_no)
        if compact_item_no and compact_item_no not in supply_lookup and item_no in supply_lookup:
            supply_lookup[compact_item_no] = supply_lookup[item_no]
        normalized_item_no = normalize_key(item_no)
        if normalized_item_no and item_no in supply_lookup:
            supply_lookup.setdefault(normalized_item_no, supply_lookup[item_no])

        if item_no and item_no not in discount_lookup:
            discount_lookup[item_no] = to_float(row.get("할인가", 0))
        if compact_item_no and compact_item_no not in discount_lookup and item_no in discount_lookup:
            discount_lookup[compact_item_no] = discount_lookup[item_no]
        if normalized_item_no and item_no in discount_lookup:
            discount_lookup.setdefault(normalized_item_no, discount_lookup[item_no])

    return transform_lookup, stock_lookup, supply_lookup, discount_lookup


def build_variant_code_to_size_map(stock_df: pd.DataFrame) -> dict:
    code_to_size = {}

    temp = stock_df.fillna("")
    columns = list(temp.columns)

    for values in temp.itertuples(index=False, name=None):
        row = dict(zip(columns, values))
        stock_size = clean_text(row.get("사이즈", ""))
        for col in ["변환코드1", "변환코드2", "변환코드3"]:
            code = clean_text(row.get(col, ""))
            for key in code_lookup_keys(code):
                code_to_size.setdefault(key, stock_size)

    return code_to_size


def make_size_lookup(size_df: pd.DataFrame):
    lookup = {}
    valid_keys = set()

    temp = size_df.fillna("")
    columns = list(temp.columns)

    for values in temp.itertuples(index=False, name=None):
        row = dict(zip(columns, values))
        brand = normalize_key(row.get("브랜드", ""))
        size = normalize_key(row.get("사이즈", ""))
        if not brand or not size:
            continue
        valid_keys.add((brand, size))
        lookup[(brand, size)] = clean_text(row.get("사이즈 변환", ""))

    return lookup, valid_keys


def get_exchange_rate(exchange_df: pd.DataFrame, country: str) -> float:
    temp = exchange_df.fillna("").copy()
    if "국가" not in temp.columns or "환율" not in temp.columns:
        return 0.0

    matched = temp[temp["국가"].astype(str).str.strip().str.upper() == country.upper()]
    if matched.empty:
        return 0.0

    return to_float(matched.iloc[0]["환율"])


def make_delivery_lookup(platform_df: pd.DataFrame):
    lookup = {}
    temp = platform_df.fillna("").copy()

    if temp.shape[1] < 19:
        return lookup

    order_series = temp.iloc[:, 2].apply(clean_text)
    status_series = temp.iloc[:, 18].apply(clean_text)

    for order_no, delivery_text in zip(order_series, status_series):
        normalized_order_no = normalize_key(order_no)
        if normalized_order_no and normalized_order_no not in lookup:
            lookup[normalized_order_no] = delivery_text

    return lookup


def split_item_and_size(item_value, fallback_size_value):
    item_text = clean_text(item_value)
    fallback_size = normalize_size(fallback_size_value)

    if "-" not in item_text:
        return item_text, fallback_size

    left, right = item_text.rsplit("-", 1)
    left = left.strip()
    right = right.strip()

    if not left:
        return item_text, fallback_size

    derived_size = right or fallback_size
    return left, derived_size


def make_code(item_no, size_value):
    item_no = clean_text(item_no)
    size_value = clean_text(size_value)

    if item_no == "" or size_value == "":
        return ""

    return f"{item_no}{size_value}"


def get_dynamic_store_priority(stock_entry: dict) -> list[str]:
    base_order = {store: idx for idx, store in enumerate(DYNAMIC_STORE_PRIORITY)}
    return sorted(
        DYNAMIC_STORE_PRIORITY,
        key=lambda store: (
            -int(stock_entry.get(store, 0)),
            base_order[store],
        ),
    )


def get_allocation_priority(stock_entry: dict) -> list[str]:
    return (
        OFFICE_ALLOC_PRIORITY
        + get_dynamic_store_priority(stock_entry)
        + STATIC_FALLBACK_STORES
        + [RETURN_FALLBACK_STORE]
    )


def get_order_candidate_stores(order_rows: list[dict], stock_lookup: dict) -> list[str]:
    base_order = {store: idx for idx, store in enumerate(DYNAMIC_STORE_PRIORITY)}
    dynamic_totals = {store: 0 for store in DYNAMIC_STORE_PRIORITY}

    for order_row in order_rows:
        stock_entry = get_stock_entry(stock_lookup, order_row["코드"])
        for store in DYNAMIC_STORE_PRIORITY:
            dynamic_totals[store] += int(stock_entry.get(store, 0))

    dynamic_stores = sorted(
        DYNAMIC_STORE_PRIORITY,
        key=lambda store: (-dynamic_totals[store], base_order[store]),
    )

    return OFFICE_ALLOC_PRIORITY + dynamic_stores + STATIC_FALLBACK_STORES + [RETURN_FALLBACK_STORE]


def allocate_qty(stock_entry: dict, qty: int) -> list[tuple[str, int]]:
    allocations = []
    remain = qty

    for store in get_allocation_priority(stock_entry):
        if remain <= 0:
            break

        available = int(stock_entry.get(store, 0))
        if available <= 0:
            continue

        use_qty = min(available, remain)
        stock_entry[store] = available - use_qty
        remain -= use_qty
        allocations.append((store, use_qty))

    if remain > 0:
        allocations.append(("재고없음", remain))

    return allocations


def can_allocate_order_to_store(order_rows: list[dict], stock_lookup: dict, store: str) -> bool:
    required_by_code = {}

    for row in order_rows:
        code = row["코드"]
        qty = row["수량_int"]
        if code == "" or qty <= 0:
            return False
        required_by_code[code] = required_by_code.get(code, 0) + qty

    for code, required_qty in required_by_code.items():
        stock_entry = get_stock_entry(stock_lookup, code)
        if int(stock_entry.get(store, 0)) < required_qty:
            return False

    return True


def apply_order_allocation_to_store(order_rows: list[dict], stock_lookup: dict, store: str) -> list[dict]:
    allocated_rows = []

    for row in order_rows:
        stock_entry = get_stock_entry(stock_lookup, row["코드"])
        stock_entry[store] = int(stock_entry.get(store, 0)) - row["수량_int"]
        row_data = row["base_output"].copy()
        row_data["매장명"] = store
        allocated_rows.append(row_data)

    return allocated_rows


def sort_output_df(df: pd.DataFrame) -> pd.DataFrame:
    temp = df.copy()
    temp["매장명_순서"] = temp["매장명"].map(STORE_ORDER_INDEX).fillna(999).astype(int)
    temp = temp.sort_values(
        by=["플랫폼", "매장명_순서", "매장명", "주문번호", "품번", "사이즈"],
        ascending=[True, True, True, True, True, True],
    ).drop(columns=["매장명_순서"]).reset_index(drop=True)
    return temp


def build_output_df(source_df: pd.DataFrame, transform_lookup, stock_lookup, supply_lookup, discount_lookup, size_lookup, size_valid_keys, variant_code_to_size, exchange_rate, delivery_lookup) -> pd.DataFrame:
    source_df = source_df.fillna("")
    today = today_str()

    prepared_rows = []
    for row in source_df.itertuples(index=False, name=None):
        item_no, raw_size = split_item_and_size(row[5], row[7])
        stock_info = transform_lookup.get(normalize_key(item_no), {})
        converted_item_no = stock_info.get("품번_변환", "") or item_no
        brand = stock_info.get("브랜드", "")
        size1 = size_step1(raw_size)
        size2 = size_step2(size1)
        size_key = (normalize_key(brand), normalize_key(size2))
        converted_size = size_lookup.get(size_key, "") if size_key in size_valid_keys else ""
        final_converted_size = converted_size or size2
        variant_code = make_code(converted_item_no, final_converted_size)
        output_size = ""
        for key in code_lookup_keys(variant_code):
            output_size = variant_code_to_size.get(key, "")
            if output_size:
                break
        output_size = output_size or final_converted_size or size2
        code = make_code(converted_item_no, output_size)
        qty = to_int(clean_text(row[8]))
        price = to_float(clean_text(row[9]))
        actual_price = price * 0.7
        fee = price - actual_price
        discount_price = discount_lookup.get(converted_item_no, 0.0)
        krw = round(actual_price * exchange_rate)
        supply_unit_price = supply_lookup.get(converted_item_no, 0.0) / 1.1
        supply_price = supply_unit_price * qty
        margin = krw - supply_price if krw or supply_price else 0.0
        order_no = clean_text(row[0])
        normalized_order_no = normalize_key(order_no)
        stock_entry = get_stock_entry(stock_lookup, code)
        stock_snapshot = {col: stock_entry.get(col, 0) for col in ALLOC_PRIORITY}
        base_output = {
            "날짜": today,
            "원거래 날짜": "",
            "플랫폼": "KASHION",
            "주문번호": order_no,
            "뒤 4자리": get_order_suffix(order_no),
            "브랜드": brand,
            "코드": code,
            "품번": converted_item_no,
            "사이즈": output_size,
            "수량": str(qty),
            "매장명": "",
            "판매가": format_number(price),
            "수수료": format_number(fee),
            "수입": format_number(actual_price),
            "할인가": format_number(discount_price),
            "환율": format_number(exchange_rate),
            "KRW": format_number(krw),
            "공급가": format_number(supply_price),
            "마진": format_number(margin),
            "내역": delivery_lookup.get(normalized_order_no, ""),
        }
        base_output.update({
            col: ("" if stock_snapshot.get(col, 0) == 0 else stock_snapshot.get(col, 0))
            for col in ALLOC_PRIORITY
        })
        prepared_rows.append({
            "주문번호": order_no,
            "코드": code,
            "수량_int": qty,
            "base_output": base_output,
        })

    transformed_rows = []
    prepared_df = pd.DataFrame(prepared_rows)

    for _, group_df in prepared_df.groupby("주문번호", sort=False):
        order_rows = group_df.to_dict("records")
        assigned = False

        candidate_stores = get_order_candidate_stores(order_rows, stock_lookup)
        for store in candidate_stores:
            if can_allocate_order_to_store(order_rows, stock_lookup, store):
                transformed_rows.extend(apply_order_allocation_to_store(order_rows, stock_lookup, store))
                assigned = True
                break

        if assigned:
            continue

        for order_row in order_rows:
            stock_entry = get_stock_entry(stock_lookup, order_row["코드"])
            allocations = allocate_qty(stock_entry, order_row["수량_int"])

            for store, alloc_qty in allocations:
                row_data = order_row["base_output"].copy()
                row_data["매장명"] = store
                row_data["수량"] = str(alloc_qty)
                transformed_rows.append(row_data)

    result_df = pd.DataFrame(transformed_rows, columns=OUTPUT_COLUMNS)
    return sort_output_df(result_df)


def build_searching_df(platform_reference_df: pd.DataFrame) -> pd.DataFrame:
    if platform_reference_df.shape[1] < 19:
        return pd.DataFrame()

    temp = platform_reference_df.copy()
    temp.columns = [clean_text(col) for col in temp.columns]
    temp["__source_order"] = range(len(temp))

    order_series = get_first_column_series(temp, ["주문번호"]).apply(clean_text)
    code_series = get_first_column_series(temp, ["코드"]).apply(clean_text)
    item_series = get_first_column_series(temp, ["품번", "품번 변환"]).apply(clean_text)
    size_series = get_first_column_series(temp, ["사이즈", "사이즈 변환"]).apply(clean_text)
    product_key_series = code_series.where(code_series.ne(""), item_series + "|" + size_series)
    temp["__latest_key"] = order_series + "|" + product_key_series
    original_date_lookup = (
        temp.drop_duplicates("__latest_key", keep="first")
        .set_index("__latest_key")["날짜"]
        .to_dict()
    )

    newest_first = temp.iloc[::-1].copy()

    newest_first = newest_first[~newest_first["__latest_key"].duplicated()].copy()
    newest_first["__original_date"] = (
        newest_first["__latest_key"]
        .map(original_date_lookup)
        .fillna(get_first_column_series(newest_first, ["날짜"]))
    )

    status_series = get_first_column_series(newest_first, ["내역"]).apply(clean_text)
    searching_values = {"调货中"}
    mask = status_series.isin(searching_values)
    return newest_first[mask].copy().reset_index(drop=True)


def get_first_column_series(df: pd.DataFrame, column_names: list[str]) -> pd.Series:
    for column_name in column_names:
        if column_name in df.columns:
            return df[column_name].fillna("").astype(str)
    return pd.Series([""] * len(df), index=df.index, dtype=str)


def build_searching_out_df(searching_df: pd.DataFrame) -> pd.DataFrame:
    out_df = pd.DataFrame("", index=searching_df.index, columns=OUTPUT_COLUMNS)
    if searching_df.empty:
        return out_df

    column_aliases = {
        "판매가": ["판매가", "총 판매가"],
        "수수료": ["수수료", "총 수수료"],
        "수입": ["수입", "총 수입", "실 판매가"],
        "KRW": ["KRW", "총 KRW"],
        "공급가": ["공급가", "총 공급가"],
    }

    for output_column in OUTPUT_COLUMNS:
        aliases = column_aliases.get(output_column, [output_column])
        out_df[output_column] = get_first_column_series(searching_df, aliases)

    out_df["원거래 날짜"] = get_first_column_series(searching_df, ["__original_date", "날짜"]).reset_index(drop=True)
    out_df["날짜"] = today_str()

    status_series = get_first_column_series(searching_df, ["내역"]).apply(clean_text)
    out_df.loc[status_series.eq("调货中"), "매장명"] = "재고없음"

    return out_df.fillna("").astype(str).reset_index(drop=True)


def upload_to_google_sheet(df: pd.DataFrame, sheet_name: str):
    helper_columns = [col for col in df.columns if str(col).startswith("__")]
    upload_df = df.drop(columns=helper_columns, errors="ignore")
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": sheet_name,
        "values": [upload_df.columns.tolist()] + upload_df.fillna("").astype(str).values.tolist(),
    }

    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            print(f"업로드 완료: {sheet_name} / {len(df)}행")
            print("응답:", response.text)
            return
        except Exception as e:
            last_error = e
            print(f"업로드 실패 ({sheet_name}, {attempt}/{RETRIES}): {e}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def upload_multiple_sheets(upload_jobs: list[tuple[pd.DataFrame, str]]):
    with ThreadPoolExecutor(max_workers=min(4, len(upload_jobs))) as executor:
        futures = [
            executor.submit(upload_to_google_sheet, df, sheet_name)
            for df, sheet_name in upload_jobs
        ]
        for future in as_completed(futures):
            future.result()


def main():
    start_time = time.perf_counter()
    csv_texts = fetch_all_csv_texts({
        "source": SOURCE_CSV_URL,
        "stock": STOCK_TRANSFORM_CSV_URL,
        "size": ONLINE_SIZE_CSV_URL,
        "exchange": EXCHANGE_RATE_CSV_URL,
        "platform": PLATFORM_INFO_CSV_URL,
        "squareone_location": SQUAREONE_LOCATION_CSV_URL,
        "fila_paju_location": FILA_PAJU_LOCATION_CSV_URL,
    })

    source_df = pd.read_csv(StringIO(csv_texts["source"]), dtype=str).fillna("")
    stock_df = pd.read_csv(StringIO(csv_texts["stock"]), dtype=str).fillna("")
    size_df = pd.read_csv(StringIO(csv_texts["size"]), dtype=str).fillna("")
    exchange_df = pd.read_csv(StringIO(csv_texts["exchange"]), dtype=str).fillna("")
    platform_reference_df = pd.read_csv(StringIO(csv_texts["platform"]), dtype=str).fillna("")
    squareone_location_df = pd.read_csv(StringIO(csv_texts["squareone_location"]), dtype=str).fillna("")
    fila_paju_location_raw = pd.read_csv(StringIO(csv_texts["fila_paju_location"]), dtype=str).fillna("")
    fila_paju_location_df = load_selected_excel_cols(fila_paju_location_raw, ["E", "F"])

    if source_df.shape[1] < 10:
        raise ValueError("원본 CSV 컬럼 수가 부족합니다. A~J 열을 확인해주세요.")

    transform_lookup, stock_lookup, supply_lookup, discount_lookup = make_stock_lookups(stock_df)
    size_lookup, size_valid_keys = make_size_lookup(size_df)
    variant_code_to_size = build_variant_code_to_size_map(stock_df)
    exchange_rate = get_exchange_rate(exchange_df, "CNY")
    delivery_lookup = make_delivery_lookup(platform_reference_df)
    output_df = build_output_df(
        source_df.iloc[:, :10].copy(),
        transform_lookup,
        stock_lookup,
        supply_lookup,
        discount_lookup,
        size_lookup,
        size_valid_keys,
        variant_code_to_size,
        exchange_rate,
        delivery_lookup,
    )
    searching_df = build_searching_df(platform_reference_df)
    searching_out_df = build_searching_out_df(searching_df)
    combined_output_df = pd.concat([output_df, searching_out_df], ignore_index=True)
    upload_multiple_sheets([
        (combined_output_df, SHEET_NAME),
        (searching_df, SEARCHING_SHEET_NAME),
        (squareone_location_df, SQUAREONE_LOCATION_SHEET_NAME),
        (fila_paju_location_df, FILA_PAJU_LOCATION_SHEET_NAME),
    ])
    elapsed = time.perf_counter() - start_time
    print(f"완료 (총 소요 시간: {elapsed:.2f}초)")


if __name__ == "__main__":
    main()
