# -*- coding: utf-8 -*-

from io import StringIO
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests


INPUT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSm0aI5pid38q7OLU_E1DNM9j8iULBFUPlbIPetHbvbDtgEZmmjPVo6UyQVgcSb0KDRxNCN0fdvWCpI/pub?gid=0&single=true&output=csv"
STOCK_PREP_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"
SQUAREONE_LOCATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSZ4Mgu9j6y27nLBYU8gAhDTfy4eMpvBgvs3oorR3BUCpcgoyf6Z1SllaqsFyos8LcH5DfxoUsN4NYG/pub?gid=289091756&single=true&output=csv"
FILA_LOCATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSA47SgFq9QQPg0D3AlBnpJX6q7Yx_Dh66E1ID9MlXTahJjL0FmFVtPgyTEtj4iVj7PvRkCUoCgbjkd/pub?gid=1813802704&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxH25CXp15V5tD6wslmAk3odB37hoMtAPEv8vddwJjEUUF2cWuZ0yZZMSknVzyuI4IV/exec"
SPREADSHEET_ID = "1pi4Z6rxa6734VFRCfFFGc_eRQ_vtJYgpS74yJJwSk54"
TARGET_SHEET_NAME = "OUT"
SQUAREONE_LOCATION_SHEET_NAME = "스퀘어원 제품 위치"
FILA_LOCATION_SHEET_NAME = "휠라 제품 위치"

DOWNLOAD_RETRIES = 3
DOWNLOAD_TIMEOUT = (10, 60)
DOWNLOAD_SLEEP_SEC = 2
UPLOAD_RETRIES = 3
UPLOAD_TIMEOUT = (10, 300)
UPLOAD_SLEEP_SEC = 3

OFFICE_ALLOC_PRIORITY = [
    "사무실 - 사무실",
    "사무실 - 구월",
    "사무실 - 부천",
    "사무실 - 스퀘어원",
    "사무실 - 아디다스 키즈",
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

STORE_ALLOC_PRIORITY = SECONDARY_DYNAMIC_STORES + SECONDARY_FALLBACK_STORES
STOCK_REFERENCE_COLUMNS = OFFICE_ALLOC_PRIORITY + STORE_ALLOC_PRIORITY + [FILA_FIXED_STORE, RETURN_FALLBACK_STORE]
VALID_STORE_VALUES = STOCK_REFERENCE_COLUMNS + ["재고없음"]
STORE_ORDER_MAP = {store: idx for idx, store in enumerate(VALID_STORE_VALUES)}
PRICE_SOURCE_COLUMNS = ["최초가", "할인율", "할인가", "공급가"]
SALES_OUTPUT_COLUMNS = ["할인율", "최초가", "할인가", "총 판매가", "총 수수료", "총 수입", "환율", "총 KRW", "총 공급가", "마진", "내역"]
OUTPUT_COLUMNS = ["날짜", "플랫폼", "주문번호", "뒷 4자리", "브랜드", "코드", "품번", "사이즈", "수량", "매장명"] + SALES_OUTPUT_COLUMNS + STOCK_REFERENCE_COLUMNS


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(message, flush=True)


def make_session():
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(session, url, retries=DOWNLOAD_RETRIES, timeout=DOWNLOAD_TIMEOUT):
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            response = session.get(url, timeout=timeout)
            response.raise_for_status()
            response.encoding = "utf-8"
            return response.text
        except Exception as exc:
            last_error = exc
            log(f"다운로드 실패 ({attempt}/{retries}): {url}")
            if attempt < retries:
                time.sleep(DOWNLOAD_SLEEP_SEC)
    raise last_error


def read_google_sheet_csv_from_text(csv_text: str) -> pd.DataFrame:
    return pd.read_csv(StringIO(csv_text), dtype=str).fillna("")


def read_google_sheet_csv(url: str) -> pd.DataFrame:
    return read_google_sheet_csv_from_text(fetch_csv_text(make_session(), url))


def download_csvs_parallel(named_urls: dict[str, str]) -> dict[str, pd.DataFrame]:
    results = {}
    with ThreadPoolExecutor(max_workers=len(named_urls)) as executor:
        futures = {executor.submit(read_google_sheet_csv, url): name for name, url in named_urls.items()}
        for future in as_completed(futures):
            results[futures[future]] = future.result()
    return results


def upload_to_google_sheet(df: pd.DataFrame, web_app_url: str, spreadsheet_id: str, sheet_name: str):
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()

    payload = {
        "spreadsheetId": spreadsheet_id,
        "sheetName": sheet_name,
        "values": values,
    }

    log(f"구글 시트 업로드 시작: {sheet_name} / {len(df)}행")
    last_error = None

    for attempt in range(1, UPLOAD_RETRIES + 1):
        try:
            response = requests.post(web_app_url, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            log("구글 시트 업로드 완료")
            log(f"응답: {response.text}")
            return
        except requests.exceptions.RequestException as exc:
            last_error = exc
            log(f"구글 시트 업로드 실패 ({sheet_name}, {attempt}/{UPLOAD_RETRIES}): {exc}")
            if attempt < UPLOAD_RETRIES:
                time.sleep(UPLOAD_SLEEP_SEC)

    raise last_error


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("\\", "/")
    text = re.sub(r"\s+", " ", text)
    return text


def compact_code_text(value) -> str:
    return clean_text(value).replace("-", "")


def to_number(value, default=0):
    text = clean_text(value).replace(",", "")
    if text == "":
        return default
    try:
        return float(text)
    except Exception:
        return default


def parse_discount_rate(value) -> float:
    text = clean_text(value).replace(",", "")
    if text == "":
        return 0
    if text.endswith("%"):
        return to_number(text[:-1], 0) / 100
    number = to_number(text, 0)
    return number / 100 if number > 1 else number


def calculate_live_sale_price(original_price, discount_rate) -> int:
    original = to_number(original_price, 0)
    rate = parse_discount_rate(discount_rate)
    if rate == 0:
        multiplier = 0.72
    elif rate < 0.30:
        multiplier = 0.62
    else:
        multiplier = 0.55
    return int(round(original * multiplier, 0))


def parse_discount_rate_series(series: pd.Series) -> pd.Series:
    text = series.fillna("").astype(str).str.strip().str.replace(",", "", regex=False)
    percent_mask = text.str.endswith("%", na=False)
    numeric = pd.to_numeric(text.str.rstrip("%"), errors="coerce").fillna(0)
    rate = numeric.where(~percent_mask, numeric / 100)
    rate = rate.where(rate <= 1, rate / 100)
    return rate


def calculate_live_sale_price_series(original_price_series: pd.Series, discount_rate_series: pd.Series) -> pd.Series:
    original = pd.to_numeric(original_price_series.fillna("").astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    rate = parse_discount_rate_series(discount_rate_series)
    multiplier = pd.Series(0.55, index=original.index)
    multiplier = multiplier.mask(rate.eq(0), 0.72)
    multiplier = multiplier.mask(rate.gt(0) & rate.lt(0.30), 0.62)
    return (original * multiplier).round(0).astype(int)


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


def load_selected_excel_cols(df_raw: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    use_cols = [excel_col_to_index(col) for col in cols]
    ensure_min_columns(df_raw, use_cols, f"선택 열 {cols}")
    df = df_raw.iloc[:, use_cols].copy().fillna("")
    df.columns = cols
    return df


def normalize_brand_series(series: pd.Series) -> pd.Series:
    source = (
        series.fillna("").astype(str).str.strip().str.upper()
        .str.replace("（", "(", regex=False)
        .str.replace("）", ")", regex=False)
        .str.replace("\\", "/", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )

    result = pd.Series("", index=series.index, dtype=str)
    result = result.mask(source.str.contains("NORTH FACE", na=False), "THE NORTH FACE")
    result = result.mask(source.str.contains("NIKE|나이키", na=False), "NIKE")
    result = result.mask(source.str.contains("ASICS|아식스", na=False), "ASICS")
    result = result.mask(source.str.contains("CONVERSE|컨버스", na=False), "CONVERSE")
    result = result.mask(source.str.contains("ADIDAS", na=False), "ADIDAS")
    result = result.mask(source.str.contains("FILA", na=False), "FILA")
    result = result.mask(source.str.contains("PUMA", na=False), "PUMA")
    return result


def transform_item_no_series(item_series: pd.Series, brand_series: pd.Series) -> pd.Series:
    result = (
        item_series.fillna("").astype(str).str.strip()
        .str.replace("（", "(", regex=False)
        .str.replace("）", ")", regex=False)
        .str.replace("\\", "/", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )

    puma_mask = brand_series.eq("PUMA")
    tnf_fila_mask = brand_series.isin(["THE NORTH FACE", "FILA"])

    result = result.mask(
        puma_mask,
        result.where(
            ~puma_mask,
            result
            .str.replace("(黑色标)", "", regex=False)
            .str.replace("（黑色标）", "", regex=False)
            .str.replace("鞋", "", regex=False),
        ),
    )

    tnf_fila_values = result.where(~tnf_fila_mask, result)
    tnf_fila_values = tnf_fila_values.str.replace(r"^1100", "", regex=True)
    tnf_fila_values = tnf_fila_values.str.replace("_", "", regex=False).str.replace("-", "", regex=False)
    result = result.mask(tnf_fila_mask, tnf_fila_values)

    return result.str.strip()


def normalize_size_series(series: pd.Series) -> pd.Series:
    result = (
        series.fillna("").astype(str).str.strip()
        .str.replace("（", "(", regex=False)
        .str.replace("）", ")", regex=False)
        .str.replace("\\", "/", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )

    has_space = result.str.contains(" ", regex=False)
    result = result.mask(has_space, result.str.rsplit(" ", n=1).str[-1])

    has_slash = result.str.contains("/", regex=False)
    result = result.mask(has_slash, result.str.split("/", n=1).str[0])

    numeric = pd.to_numeric(result, errors="coerce")
    int_mask = numeric.notna() & (numeric % 1 == 0)
    result = result.mask(int_mask, numeric.fillna(0).astype(int).astype(str))

    return result.str.upper()


def load_input_orders(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    if df.shape[1] < 4:
        raise DataValidationError(f"입력 시트 열 개수가 부족합니다. 실제 열 수: {df.shape[1]}")

    df = df.iloc[:, :4].copy()
    df.columns = ["브랜드원본", "품번원본", "사이즈원본", "수량"]
    df = df.fillna("").astype(str).apply(lambda col: col.str.strip())

    df["브랜드"] = normalize_brand_series(df["브랜드원본"])
    df["품번"] = transform_item_no_series(df["품번원본"], df["브랜드"])
    df["사이즈"] = normalize_size_series(df["사이즈원본"])
    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)

    df = df[(df["품번"] != "") & (df["사이즈"] != "") & (df["수량"] > 0)].copy()
    return df[["브랜드", "품번", "사이즈", "수량"]].reset_index(drop=True)


def load_stock_prepare(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()

    required = ["브랜드", "코드", "품번", "사이즈", "변환코드1", "변환코드2", "변환코드3"]
    for column in required:
        if column not in df.columns:
            raise DataValidationError(f"분배준비 시트에 필수 열이 없습니다: {column}")

    for column in STOCK_REFERENCE_COLUMNS:
        if column not in df.columns:
            df[column] = 0
    for column in PRICE_SOURCE_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    keep_columns = required + PRICE_SOURCE_COLUMNS + STOCK_REFERENCE_COLUMNS
    df = df[keep_columns].copy()

    text_columns = ["브랜드", "코드", "품번", "사이즈", "변환코드1", "변환코드2", "변환코드3"]
    for column in text_columns:
        df[column] = df[column].apply(clean_text)
    for column in PRICE_SOURCE_COLUMNS:
        df[column] = df[column].apply(clean_text)

    df["브랜드"] = normalize_brand_series(df["브랜드"])
    df["품번"] = transform_item_no_series(df["품번"], df["브랜드"])

    for column in STOCK_REFERENCE_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0).astype(int)

    return df


def build_variant_code_to_size_map(stock_df: pd.DataFrame) -> dict:
    code_to_size = {}

    for row in stock_df[["사이즈", "변환코드1", "변환코드2", "변환코드3"]].itertuples(index=False, name=None):
        stock_size = clean_text(row[0])
        for code in row[1:]:
            variant_code = clean_text(code)
            if variant_code and variant_code not in code_to_size:
                code_to_size[variant_code] = stock_size
            compact_variant_code = compact_code_text(code)
            if compact_variant_code and compact_variant_code not in code_to_size:
                code_to_size[compact_variant_code] = stock_size

    return code_to_size


def resolve_final_codes(order_df: pd.DataFrame, stock_df: pd.DataFrame) -> pd.DataFrame:
    result = order_df.copy()
    variant_code_to_size = build_variant_code_to_size_map(stock_df)

    result["사이즈_숫자영문"] = result["사이즈"].fillna("").astype(str).str.upper().str.replace(r"[^0-9A-Z]", "", regex=True)
    result["후보1"] = result["품번"].astype(str) + result["사이즈"].astype(str)
    result["후보2"] = result["품번"].astype(str) + result["사이즈"].astype(str).str.upper()
    result["후보3"] = result["품번"].astype(str) + result["사이즈_숫자영문"].astype(str)

    result["최종사이즈"] = (
        result["후보1"].map(variant_code_to_size)
        .fillna(result["후보2"].map(variant_code_to_size))
        .fillna(result["후보3"].map(variant_code_to_size))
        .fillna("")
    )
    result["최종사이즈"] = result["최종사이즈"].where(result["최종사이즈"] != "", result["사이즈"])
    result["코드"] = result["품번"].astype(str) + result["최종사이즈"].astype(str)
    result["사이즈"] = result["최종사이즈"]

    return result[["브랜드", "코드", "품번", "사이즈", "수량"]].copy()


def allocate_round_robin(stock_entry: dict, stores: list[str], remain: int, start_index: int = 0) -> tuple[list[tuple[str, int]], int]:
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
            allocations.append((store, 1))
            progressed = True

        if not progressed:
            break

    return allocations, current_index


def rotate_stores_by_stock(stock_entry: dict, stores: list[str], fallback_index: int = 0) -> list[str]:
    if not stores:
        return []

    available_pairs = [
        (idx, int(stock_entry.get(store, 0)))
        for idx, store in enumerate(stores)
    ]
    max_available = max((available for _, available in available_pairs), default=0)

    if max_available <= 0:
        start_index = fallback_index % len(stores)
    else:
        start_index = next(
            idx
            for idx, available in available_pairs
            if available == max_available
        )

    return stores[start_index:] + stores[:start_index]


def allocate_by_priority(stock_entry: dict, stores: list[str], remain: int) -> list[tuple[str, int]]:
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


def allocate_secondary_stores(stock_entry: dict, remain: int, start_index: int = 0) -> tuple[list[tuple[str, int]], int]:
    allocations = []
    dynamic_stores = rotate_stores_by_stock(
        stock_entry,
        SECONDARY_DYNAMIC_STORES,
        start_index,
    )

    dynamic_allocations, next_dynamic_index = allocate_round_robin(
        stock_entry,
        dynamic_stores,
        remain,
        0,
    )
    allocations.extend(dynamic_allocations)
    remain -= sum(qty for _, qty in dynamic_allocations)

    if remain > 0:
        fallback_stores = rotate_stores_by_stock(stock_entry, SECONDARY_FALLBACK_STORES, 0)
        fallback_allocations, _ = allocate_round_robin(
            stock_entry,
            fallback_stores,
            remain,
            0,
        )
        allocations.extend(fallback_allocations)
        remain -= sum(qty for _, qty in fallback_allocations)

    return allocations, next_dynamic_index


def auto_allocate(order_df: pd.DataFrame, stock_df: pd.DataFrame) -> pd.DataFrame:
    alloc_rows = []
    office_rr_state = {}
    store_rr_state = {}
    stock_base = (
        stock_df[["코드"] + STOCK_REFERENCE_COLUMNS]
        .drop_duplicates(subset=["코드"], keep="first")
        .set_index("코드")[STOCK_REFERENCE_COLUMNS]
        .to_dict("index")
    )
    stock_lookup = stock_base
    stock_work = {code: stores.copy() for code, stores in stock_base.items()}
    compact_stock_work = {}
    compact_stock_lookup = {}
    for code, stores in stock_work.items():
        compact_code = compact_code_text(code)
        if compact_code and compact_code not in stock_work:
            compact_stock_work[compact_code] = stores.copy()
        if compact_code and compact_code not in stock_lookup:
            compact_stock_lookup[compact_code] = stock_lookup[code]
    stock_work.update(compact_stock_work)
    stock_lookup.update(compact_stock_lookup)
    stock_index = set(stock_work.keys())

    for row in order_df.itertuples(index=False):
        brand = row.브랜드
        code = row.코드
        item_no = row.품번
        size = row.사이즈
        requested_qty = int(row.수량)

        if code == "" or code not in stock_index:
            alloc_rows.append({
                "브랜드": brand,
                "코드": code,
                "품번": item_no,
                "사이즈": size,
                "수량": requested_qty,
                "매장명": "재고없음",
            })
            continue

        remain = requested_qty
        stock_entry = stock_work[code]
        office_stores = rotate_stores_by_stock(
            stock_entry,
            OFFICE_ALLOC_PRIORITY,
            office_rr_state.get(code, 0),
        )

        rr_allocations, next_rr_index = allocate_round_robin(
            stock_entry,
            office_stores,
            remain,
        )
        office_rr_state[code] = next_rr_index % len(office_stores) if office_stores else 0

        for store, use_qty in rr_allocations:
            alloc_rows.append({
                "브랜드": brand,
                "코드": code,
                "품번": item_no,
                "사이즈": size,
                "수량": use_qty,
                "매장명": store,
            })
        remain -= sum(qty for _, qty in rr_allocations)

        if remain > 0 and brand == "FILA":
            available = int(stock_entry.get(FILA_FIXED_STORE, 0))
            use_qty = min(available, remain)
            if use_qty > 0:
                stock_entry[FILA_FIXED_STORE] = available - use_qty
                alloc_rows.append({
                    "브랜드": brand,
                    "코드": code,
                    "품번": item_no,
                    "사이즈": size,
                    "수량": use_qty,
                    "매장명": FILA_FIXED_STORE,
                })
                remain -= use_qty

        if remain > 0:
            secondary_allocations, next_store_index = allocate_secondary_stores(
                stock_entry,
                remain,
                store_rr_state.get(code, 0),
            )
            store_rr_state[code] = next_store_index

            for store, use_qty in secondary_allocations:
                alloc_rows.append({
                    "브랜드": brand,
                    "코드": code,
                    "품번": item_no,
                    "사이즈": size,
                    "수량": use_qty,
                    "매장명": store,
                })
            remain -= sum(qty for _, qty in secondary_allocations)

        if remain > 0:
            return_allocations = allocate_by_priority(
                stock_entry,
                [RETURN_FALLBACK_STORE],
                remain,
            )
            for store, use_qty in return_allocations:
                alloc_rows.append({
                    "브랜드": brand,
                    "코드": code,
                    "품번": item_no,
                    "사이즈": size,
                    "수량": use_qty,
                    "매장명": store,
                })
            remain -= sum(qty for _, qty in return_allocations)

        if remain > 0:
            alloc_rows.append({
                "브랜드": brand,
                "코드": code,
                "품번": item_no,
                "사이즈": size,
                "수량": remain,
                "매장명": "재고없음",
            })

    alloc_df = pd.DataFrame(alloc_rows)
    if alloc_df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    alloc_df = (
        alloc_df.groupby(["브랜드", "코드", "품번", "사이즈", "매장명"], as_index=False)["수량"]
        .sum()
    )

    stock_lookup_df = (
        pd.DataFrame.from_dict(stock_lookup, orient="index")
        .reindex(columns=STOCK_REFERENCE_COLUMNS)
        .reset_index()
        .rename(columns={"index": "코드"})
    )
    alloc_df = alloc_df.merge(stock_lookup_df, on="코드", how="left")
    for column in STOCK_REFERENCE_COLUMNS:
        alloc_df[column] = pd.to_numeric(alloc_df[column], errors="coerce").fillna(0).astype(int).replace(0, "")

    alloc_df["매장순서"] = alloc_df["매장명"].map(STORE_ORDER_MAP).fillna(999).astype(int)
    alloc_df = alloc_df.sort_values(
        by=["매장순서", "매장명", "브랜드", "품번", "사이즈"],
        ascending=[True, True, True, True, True],
    ).drop(columns=["매장순서"]).reset_index(drop=True)

    repeat_counts = pd.to_numeric(alloc_df["수량"], errors="coerce").fillna(0).astype(int).clip(lower=0)
    alloc_df = alloc_df.loc[alloc_df.index.repeat(repeat_counts)].reset_index(drop=True)
    alloc_df["수량"] = 1
    alloc_df["날짜"] = time.strftime("%Y-%m-%d")
    alloc_df["플랫폼"] = "LIVE"
    alloc_df["주문번호"] = ""
    alloc_df["뒷 4자리"] = ""
    price_lookup_df = stock_df[["품번"] + PRICE_SOURCE_COLUMNS].drop_duplicates(subset=["품번"], keep="first")
    alloc_df = alloc_df.merge(price_lookup_df, on="품번", how="left")
    for column in PRICE_SOURCE_COLUMNS:
        alloc_df[column] = alloc_df[column].fillna("")
    alloc_df["총 판매가"] = calculate_live_sale_price_series(alloc_df["최초가"], alloc_df["할인율"])
    alloc_df["총 수수료"] = ""
    alloc_df["총 수입"] = alloc_df["총 판매가"]
    alloc_df["환율"] = ""
    alloc_df["총 KRW"] = alloc_df["총 판매가"]
    alloc_df["총 공급가"] = alloc_df["공급가"]
    supply = pd.to_numeric(alloc_df["총 공급가"].fillna("").astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    alloc_df["마진"] = (alloc_df["총 KRW"] - supply).round(0).astype(int)
    alloc_df["내역"] = ""

    return alloc_df[OUTPUT_COLUMNS].copy()


def build_allocation_result() -> pd.DataFrame:
    log("입력/분배준비 시트 다운로드 중...")
    downloads = download_csvs_parallel({
        "input": INPUT_CSV_URL,
        "stock": STOCK_PREP_URL,
    })
    input_raw = downloads["input"]
    stock_raw = downloads["stock"]

    order_df = load_input_orders(input_raw)
    stock_df = load_stock_prepare(stock_raw)
    prepared_order_df = resolve_final_codes(order_df, stock_df)

    log(f"주문 {len(prepared_order_df)}건 자동분배 시작")
    alloc_df = auto_allocate(prepared_order_df, stock_df)
    log(f"자동분배 완료: {len(alloc_df)}행")
    return alloc_df


def main():
    start_time = time.perf_counter()

    log("구글 시트 다운로드 중...")
    downloads = download_csvs_parallel({
        "input": INPUT_CSV_URL,
        "stock": STOCK_PREP_URL,
        "squareone_location": SQUAREONE_LOCATION_URL,
        "fila_location": FILA_LOCATION_URL,
    })
    input_raw = downloads["input"]
    stock_raw = downloads["stock"]
    squareone_location_df = downloads["squareone_location"]
    fila_location_raw = downloads["fila_location"]
    fila_location_df = load_selected_excel_cols(fila_location_raw, ["E", "F"])

    order_df = load_input_orders(input_raw)
    stock_df = load_stock_prepare(stock_raw)
    prepared_order_df = resolve_final_codes(order_df, stock_df)

    log(f"주문 {len(prepared_order_df)}건 자동분배 시작")
    alloc_df = auto_allocate(prepared_order_df, stock_df)
    log(f"자동분배 완료: {len(alloc_df)}행")

    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = [
            executor.submit(
                upload_to_google_sheet,
                alloc_df,
                WEB_APP_URL,
                SPREADSHEET_ID,
                TARGET_SHEET_NAME,
            ),
            executor.submit(
                upload_to_google_sheet,
                squareone_location_df,
                WEB_APP_URL,
                SPREADSHEET_ID,
                SQUAREONE_LOCATION_SHEET_NAME,
            ),
            executor.submit(
                upload_to_google_sheet,
                fila_location_df,
                WEB_APP_URL,
                SPREADSHEET_ID,
                FILA_LOCATION_SHEET_NAME,
            ),
        ]
        for future in as_completed(futures):
            future.result()

    elapsed = time.perf_counter() - start_time
    log(f"완료 (총 소요 시간: {elapsed:.2f}초)")


if __name__ == "__main__":
    main()
