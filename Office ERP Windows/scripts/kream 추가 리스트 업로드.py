# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from io import StringIO
import re
import time

import pandas as pd
import requests


SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQSGUP7ZizELC9jJqGdxawP_HRsxo-fUZ_5BgMx8NUaEPTJYhH31iiH2-_NEE4Ff9UW3OATcxopoKEU/pub?gid=164516117&single=true&output=csv"
ONLINE_SIZE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=2089368433&single=true&output=csv"
STOCK_TRANSFORM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzsa1Lf-GvPNMYJ6DZ3pw8rbFvVQ1SF9P11xpIogUoPQTt4PG6PXJzXrZyBZeP-xTkwtw/exec"
SPREADSHEET_ID = "1qdk3Pej-zz89S1ORM3zb0SkQFTL6vMNQoCdncj-pkGY"
TARGET_SHEET_NAME = "OUT"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 300)
RETRIES = 3
RETRY_SLEEP_SEC = 2

OFFICE_ALLOC_PRIORITY = [
    "사무실 - 사무실",
    "사무실 - 구월",
    "사무실 - 부천",
    "사무실 - 스퀘어원",
    "사무실 - 아디다스 키즈",
]
SECONDARY_DYNAMIC_STORES = ["스퀘어원", "부천", "구월"]
SECONDARY_FALLBACK_STORES = ["푸마 여주"]
FILA_FIXED_STORE = "휠라 파주"
RETURN_FALLBACK_STORE = "사무실 - 반품"
FINAL_FALLBACK_STORE = "사무실 - 푸마 여주"
LEGACY_STORE_ALIASES = {
    "사무실 - 아디다스 키즈": "사무실 - S마켓",
}
ALLOC_PRIORITY = (
    OFFICE_ALLOC_PRIORITY
    + SECONDARY_DYNAMIC_STORES
    + SECONDARY_FALLBACK_STORES
    + [FILA_FIXED_STORE, RETURN_FALLBACK_STORE, FINAL_FALLBACK_STORE]
)
STOCK_REFERENCE_COLUMNS = ALLOC_PRIORITY

OUTPUT_COLUMNS = ["날짜", "구분", "플랫폼", "주문번호", "운송장번호", "브랜드", "코드", "품번", "사이즈", "수량", "매장명", "할인가", "총 판매가", "총 수수료", "총 수입", "환율", "총 KRW", "총 공급가", "마진"]
INTERNAL_OUTPUT_COLUMNS = OUTPUT_COLUMNS + ["공급가", "수입", "판매가"]
DISPLAY_STOCK_COLUMNS = [
    "사무실 - 사무실",
    "사무실 - 구월",
    "사무실 - 부천",
    "사무실 - 스퀘어원",
    "사무실 - 아디다스 키즈",
    "사무실 - 푸마 여주",
    "스퀘어원",
    "부천",
    "구월",
    "푸마 여주",
    "휠라 파주",
    "사무실 - 반품",
]
OUTPUT_COLUMNS = OUTPUT_COLUMNS + DISPLAY_STOCK_COLUMNS


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[KREAM 추가] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


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
    return re.sub(r"\s+", " ", text)


def compact_code_text(value) -> str:
    return clean_text(value).replace("-", "")


def normalize_size_match_text(value) -> str:
    text = clean_text(value)
    text = re.sub(r"\.0+$", "", text)
    return re.sub(r"^0+(?=\d)", "", text)


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
    result = result.mask(s.str.contains("CONVERSE|컨버스", na=False), "CONVERSE")
    result = result.mask(s.str.contains("ADIDAS", na=False), "ADIDAS")
    result = result.mask(s.str.contains("FILA", na=False), "FILA")
    result = result.mask(s.str.contains("PUMA", na=False), "PUMA")
    return result


def transform_item_no_series(item_series: pd.Series, brand_series: pd.Series) -> pd.Series:
    result = (
        item_series.fillna("").astype(str).str.strip()
        .str.replace("（", "(", regex=False)
        .str.replace("）", ")", regex=False)
        .str.replace("\\", "/", regex=False)
        .str.replace(r"\s+", "", regex=True)
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
            .str.replace("鞋", "", regex=False)
        )
    )

    tnf_fila_values = result.where(~tnf_fila_mask, result)
    tnf_fila_values = tnf_fila_values.str.replace(r"^1100", "", regex=True)
    tnf_fila_values = tnf_fila_values.str.replace("_", "", regex=False).str.replace("-", "", regex=False)
    result = result.mask(tnf_fila_mask, tnf_fila_values)
    return result.str.strip()


def normalize_size_formula_text(value) -> str:
    text = clean_text(value)
    if text == "":
        return ""
    if "(" in text and ")" in text and text.find("(") < text.find(")"):
        text = text[text.find("(") + 1:text.find(")", text.find("(") + 1)]
    elif text.upper().startswith("W"):
        text = text[1:]
    return "".join(re.findall(r"[A-Za-z0-9.]+", text))


def size_step1_series(series: pd.Series) -> pd.Series:
    s = series.fillna("").astype(str).apply(normalize_size_formula_text)
    has_space = s.str.contains(" ", regex=False)
    s = s.mask(has_space, s.str.rsplit(" ", n=1).str[-1])
    has_slash = s.str.contains("/", regex=False)
    s = s.mask(has_slash, s.str.split("/", n=1).str[0])
    return s.apply(clean_text)


def size_step2_series(series: pd.Series) -> pd.Series:
    s = (
        series.fillna("").astype(str).str.strip()
        .str.replace("（", "(", regex=False)
        .str.replace("）", ")", regex=False)
        .str.replace("\\", "/", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )
    return s.mask(s.str.contains(r"[\u4e00-\u9fff]", na=False), "ONE")


def to_number_text(value) -> str:
    text = clean_text(value).replace(",", "")
    if text == "":
        return ""
    number = pd.to_numeric(text, errors="coerce")
    if pd.isna(number):
        return clean_text(value)
    return str(int(round(float(number))))


def load_online_size_map(df_raw: pd.DataFrame) -> pd.DataFrame:
    if df_raw.shape[1] < 3:
        raise DataValidationError("온라인 사이즈표 열 수가 부족합니다.")
    df = df_raw.iloc[:, :3].copy()
    df.columns = ["브랜드", "사이즈", "사이즈변환"]
    df["브랜드"] = normalize_brand_series(df["브랜드"])
    df["사이즈"] = df["사이즈"].apply(normalize_size_match_text)
    df["사이즈변환"] = df["사이즈변환"].apply(normalize_size_match_text)
    return df


def build_online_size_lookup(size_map_df: pd.DataFrame) -> dict[tuple[str, str], str]:
    temp = size_map_df[["브랜드", "사이즈", "사이즈변환"]].copy()
    temp["key"] = list(zip(temp["브랜드"], temp["사이즈"]))
    temp = temp.drop_duplicates(subset=["key"], keep="first")
    return dict(zip(temp["key"], temp["사이즈변환"]))


def load_stock_prepare(df_raw: pd.DataFrame) -> pd.DataFrame:
    df_source = df_raw.copy()
    for col, legacy_col in LEGACY_STORE_ALIASES.items():
        if col not in df_source.columns and legacy_col in df_source.columns:
            df_source[col] = df_source[legacy_col]

    required = ["코드", "품번", "사이즈", "변환코드1", "변환코드2", "변환코드3", "할인가", "공급가"] + STOCK_REFERENCE_COLUMNS
    missing = [col for col in required if col not in df_source.columns]
    if missing:
        raise DataValidationError(f"분배준비 시트 필수 열이 없습니다: {missing}")

    df = df_source[required].copy()
    text_cols = ["코드", "품번", "사이즈", "변환코드1", "변환코드2", "변환코드3"]
    for col in text_cols:
        df[col] = df[col].apply(clean_text)
    for col in ["할인가", "공급가"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in STOCK_REFERENCE_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    return df


def build_variant_code_to_size_map(stock_df: pd.DataFrame) -> dict[str, str]:
    code_to_size = {}
    for row in stock_df[["사이즈", "변환코드1", "변환코드2", "변환코드3"]].itertuples(index=False, name=None):
        stock_size = normalize_size_match_text(row[0])
        for code in row[1:]:
            code = clean_text(code)
            compact_code = compact_code_text(code)
            if code and code not in code_to_size:
                code_to_size[code] = stock_size
            if compact_code and compact_code not in code_to_size:
                code_to_size[compact_code] = stock_size
    return code_to_size


def build_stock_code_lookup(stock_df: pd.DataFrame) -> dict[str, dict]:
    lookup = {}
    keep_cols = ["코드"] + STOCK_REFERENCE_COLUMNS
    rows = stock_df[keep_cols].drop_duplicates(subset=["코드"], keep="first").to_dict("records")
    for row in rows:
        code = clean_text(row.get("코드", ""))
        stock_entry = {col: int(pd.to_numeric(row.get(col, 0), errors="coerce") or 0) for col in STOCK_REFERENCE_COLUMNS}
        if code:
            lookup[code] = stock_entry.copy()
        compact_code = compact_code_text(code)
        if compact_code and compact_code not in lookup:
            lookup[compact_code] = stock_entry.copy()
    return lookup


def build_price_lookup(stock_df: pd.DataFrame, column_name: str) -> dict[str, float]:
    temp = stock_df[["품번", column_name]].copy()
    temp["품번"] = temp["품번"].apply(clean_text)
    temp[column_name] = pd.to_numeric(temp[column_name], errors="coerce").fillna(0)
    temp = temp.drop_duplicates(subset=["품번"], keep="first")
    lookup = dict(zip(temp["품번"], temp[column_name]))
    for item_no, value in temp[["품번", column_name]].itertuples(index=False, name=None):
        compact_item_no = compact_code_text(item_no)
        if compact_item_no and compact_item_no not in lookup:
            lookup[compact_item_no] = value
    return lookup


def build_display_stock_lookup(stock_df: pd.DataFrame) -> dict[str, dict]:
    keep_cols = ["코드"] + DISPLAY_STOCK_COLUMNS
    rows = stock_df[keep_cols].drop_duplicates(subset=["코드"], keep="first").to_dict("records")
    lookup = {}
    for row in rows:
        code = clean_text(row.get("코드", ""))
        stock_entry = {col: int(pd.to_numeric(row.get(col, 0), errors="coerce") or 0) for col in DISPLAY_STOCK_COLUMNS}
        if code:
            lookup[code] = stock_entry.copy()
        compact_code = compact_code_text(code)
        if compact_code and compact_code not in lookup:
            lookup[compact_code] = stock_entry.copy()
    return lookup


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
            progressed = True
            allocations.append((store, 1))
        if not progressed:
            break
    return allocations, current_index


def allocate_by_priority(stock_entry: dict, stores: list[str], remain: int) -> list[tuple[str, int]]:
    allocations = []
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


def allocate_secondary_stores(stock_entry: dict, remain: int) -> list[tuple[str, int]]:
    allocations = []
    dynamic_stores = sorted(
        SECONDARY_DYNAMIC_STORES,
        key=lambda store: (-int(stock_entry.get(store, 0)), SECONDARY_DYNAMIC_STORES.index(store)),
    )
    allocations.extend(allocate_by_priority(stock_entry, dynamic_stores, remain))
    remain -= sum(qty for _, qty in allocations)
    if remain > 0:
        allocations.extend(allocate_by_priority(stock_entry, SECONDARY_FALLBACK_STORES, remain))
    return allocations


def load_all_sources() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    csvs = fetch_all_csv_texts(
        {
            "source": SOURCE_CSV_URL,
            "online_size": ONLINE_SIZE_CSV_URL,
            "stock": STOCK_TRANSFORM_CSV_URL,
        }
    )
    source_df = pd.read_csv(StringIO(csvs["source"]), dtype=str, keep_default_na=False).fillna("")
    size_df = pd.read_csv(StringIO(csvs["online_size"]), dtype=str, keep_default_na=False).fillna("")
    stock_df = pd.read_csv(StringIO(csvs["stock"]), dtype=str, keep_default_na=False).fillna("")
    return source_df, size_df, stock_df


def prepare_order_source(source_df: pd.DataFrame, size_map_df: pd.DataFrame, stock_df: pd.DataFrame) -> pd.DataFrame:
    required_columns = ["주문/보관번호", "모델품번", "상품명(영문 상품명)", "옵션", "거래금액", "정산 예정 금액", "운송장번호"]
    missing = [col for col in required_columns if col not in source_df.columns]
    if missing:
        raise DataValidationError(f"KREAM 원본 시트 필수 열이 없습니다: {missing}")

    src = pd.DataFrame(
        {
            "주문번호": source_df["주문/보관번호"].apply(clean_text),
            "품번원본": source_df["모델품번"].apply(clean_text),
            "브랜드원본": source_df["상품명(영문 상품명)"].apply(clean_text),
            "사이즈원본": source_df["옵션"].apply(clean_text),
            "판매가": source_df["거래금액"].apply(to_number_text),
            "수입": source_df["정산 예정 금액"].apply(to_number_text),
            "운송장번호": source_df["운송장번호"].apply(clean_text),
            "수량": 1,
        }
    )
    src = src[src["주문번호"] != ""].reset_index(drop=True)

    src["브랜드"] = normalize_brand_series(src["브랜드원본"])
    src["품번"] = transform_item_no_series(src["품번원본"], src["브랜드"])
    src["사이즈1"] = size_step1_series(src["사이즈원본"])
    src["사이즈2"] = size_step2_series(src["사이즈1"]).apply(normalize_size_match_text)

    size_lookup = build_online_size_lookup(load_online_size_map(size_map_df))
    src["사이즈3"] = list(zip(src["브랜드"], src["사이즈2"]))
    src["사이즈3"] = src["사이즈3"].map(size_lookup).fillna("")
    src["사이즈3"] = src["사이즈3"].where(src["사이즈3"].astype(str).str.strip() != "", src["사이즈2"])

    variant_code_to_size = build_variant_code_to_size_map(stock_df)
    src["코드_원본기준"] = src["품번"] + src["사이즈2"]
    src["코드_원본보조"] = src["품번"] + src["사이즈2"].apply(normalize_size_match_text)
    src["코드_3차기준"] = src["품번"] + src["사이즈3"]
    src["코드_3차보조"] = src["품번"] + src["사이즈3"].apply(normalize_size_match_text)

    src["최종사이즈"] = (
        src["코드_원본기준"].map(variant_code_to_size)
        .fillna(src["코드_원본보조"].map(variant_code_to_size))
        .fillna(src["코드_3차기준"].map(variant_code_to_size))
        .fillna(src["코드_3차보조"].map(variant_code_to_size))
        .fillna("")
    )
    src["최종사이즈"] = src["최종사이즈"].where(src["최종사이즈"].astype(str).str.strip() != "", src["사이즈3"])
    src["매칭사이즈"] = src["최종사이즈"]
    src["출력사이즈"] = src["최종사이즈"]
    src["매칭코드"] = src["품번"] + src["매칭사이즈"]
    src["최종코드"] = src["품번"] + src["출력사이즈"]

    stock_code_set = set(build_stock_code_lookup(stock_df).keys())
    src["매칭코드"] = src["매칭코드"].where(src["매칭코드"].isin(stock_code_set), src["매칭코드"].apply(compact_code_text))
    src = src[["주문번호", "브랜드", "품번", "출력사이즈", "최종코드", "매칭코드", "판매가", "수입", "운송장번호", "수량"]].copy()
    src["수량"] = pd.to_numeric(src["수량"], errors="coerce").fillna(0).astype(int)
    src["판매가"] = pd.to_numeric(src["판매가"], errors="coerce").fillna(0).round(0).astype(int).astype(str)
    src["수입"] = pd.to_numeric(src["수입"], errors="coerce").fillna(0).round(0).astype(int).astype(str)
    return src.reset_index(drop=True)


def auto_allocate(order_df: pd.DataFrame, stock_df: pd.DataFrame) -> pd.DataFrame:
    alloc_rows = []
    office_rr_state = {}
    stock_lookup = build_stock_code_lookup(stock_df)
    display_stock_lookup = build_display_stock_lookup(stock_df)
    discount_lookup = build_price_lookup(stock_df, "할인가")
    supply_lookup = build_price_lookup(stock_df, "공급가")
    stock_index = set(stock_lookup.keys())

    for row in order_df.itertuples(index=False):
        match_code = clean_text(row.매칭코드)
        final_code = clean_text(row.최종코드)
        remain = int(row.수량)

        if match_code == "" or match_code not in stock_index:
            display_stock = display_stock_lookup.get(match_code) or display_stock_lookup.get(compact_code_text(match_code)) or {col: 0 for col in DISPLAY_STOCK_COLUMNS}
            item_no = clean_text(row.품번)
            alloc_rows.append(
                {
                    "날짜": datetime.now().strftime("%Y-%m-%d"),
                    "구분": "추가",
                    "플랫폼": "KREAM",
                    "주문번호": clean_text(row.주문번호),
                    "운송장번호": clean_text(row.운송장번호),
                    "브랜드": clean_text(row.브랜드),
                    "코드": final_code,
                    "품번": item_no,
                    "사이즈": clean_text(row.출력사이즈),
                    "수량": remain,
                    "매장명": "재고없음",
                    "할인가": int(round(float(discount_lookup.get(item_no, discount_lookup.get(compact_code_text(item_no), 0)) or 0))),
                    "공급가": int(round(float(supply_lookup.get(item_no, supply_lookup.get(compact_code_text(item_no), 0)) or 0))),
                    "판매가": row.판매가,
                    "수입": row.수입,
                    "총 판매가": row.판매가,
                    "총 수수료": str(int(round(float(pd.to_numeric(row.판매가, errors="coerce") or 0) - float(pd.to_numeric(row.수입, errors="coerce") or 0)))),
                    "총 수입": row.수입,
                    "환율": "",
                    "총 KRW": row.수입,
                    "총 공급가": str(int(round(float(supply_lookup.get(item_no, supply_lookup.get(compact_code_text(item_no), 0)) or 0) * int(remain)))),
                    "마진": str(int(round(float(pd.to_numeric(row.수입, errors="coerce") or 0) - (float(supply_lookup.get(item_no, supply_lookup.get(compact_code_text(item_no), 0)) or 0) * int(remain))))),
                    **display_stock,
                }
            )
            continue

        stock_entry = stock_lookup[match_code]
        rr_start_index = office_rr_state.get(match_code, 0)
        allocations, next_rr_index = allocate_round_robin(stock_entry, OFFICE_ALLOC_PRIORITY, remain, rr_start_index)
        office_rr_state[match_code] = next_rr_index
        remain -= sum(qty for _, qty in allocations)

        if remain > 0 and row.브랜드 == "FILA":
            allocations.extend(allocate_by_priority(stock_entry, [FILA_FIXED_STORE], remain))
            remain -= sum(qty for _, qty in allocations if _ == FILA_FIXED_STORE)

        if remain > 0:
            secondary_allocs = allocate_secondary_stores(stock_entry, remain)
            allocations.extend(secondary_allocs)
            remain -= sum(qty for _, qty in secondary_allocs)

        if remain > 0:
            return_allocs = allocate_by_priority(stock_entry, [RETURN_FALLBACK_STORE], remain)
            allocations.extend(return_allocs)
            remain -= sum(qty for _, qty in return_allocs)

        if remain > 0:
            final_allocs = allocate_by_priority(stock_entry, [FINAL_FALLBACK_STORE], remain)
            allocations.extend(final_allocs)
            remain -= sum(qty for _, qty in final_allocs)

        if not allocations:
            allocations = [("재고없음", int(row.수량))]
            remain = 0

        if remain > 0:
            allocations.append(("재고없음", remain))

        for store, qty in allocations:
            display_stock = display_stock_lookup.get(match_code) or display_stock_lookup.get(compact_code_text(match_code)) or {col: 0 for col in DISPLAY_STOCK_COLUMNS}
            item_no = clean_text(row.품번)
            price = float(pd.to_numeric(row.판매가, errors="coerce") or 0)
            income = float(pd.to_numeric(row.수입, errors="coerce") or 0)
            supply_price = float(supply_lookup.get(item_no, supply_lookup.get(compact_code_text(item_no), 0)) or 0)
            total_sales = int(round(price * qty))
            total_income = int(round(income * qty))
            total_supply = int(round(supply_price * qty))
            alloc_rows.append(
                {
                    "날짜": datetime.now().strftime("%Y-%m-%d"),
                    "구분": "추가",
                    "플랫폼": "KREAM",
                    "주문번호": clean_text(row.주문번호),
                    "운송장번호": clean_text(row.운송장번호),
                    "브랜드": clean_text(row.브랜드),
                    "코드": final_code,
                    "품번": item_no,
                    "사이즈": clean_text(row.출력사이즈),
                    "수량": str(int(qty)),
                    "매장명": store,
                    "할인가": int(round(float(discount_lookup.get(item_no, discount_lookup.get(compact_code_text(item_no), 0)) or 0))),
                    "공급가": int(round(supply_price)),
                    "판매가": row.판매가,
                    "수입": row.수입,
                    "총 판매가": str(total_sales),
                    "총 수수료": str(total_sales - total_income),
                    "총 수입": str(total_income),
                    "환율": "",
                    "총 KRW": str(total_income),
                    "총 공급가": str(total_supply),
                    "마진": str(total_income - total_supply),
                    **display_stock,
                }
            )

    result_df = pd.DataFrame(alloc_rows)
    if not result_df.empty:
        for col in ["날짜", "구분", "플랫폼", "주문번호", "운송장번호", "브랜드", "코드", "품번", "사이즈", "매장명", "환율"]:
            result_df[col] = result_df[col].apply(clean_text)
        for col in ["수량", "할인가", "공급가", "판매가", "총 판매가", "수입", "총 수입", "총 KRW", "총 공급가", "총 수수료", "마진"]:
            result_df[f"{col}_num"] = pd.to_numeric(result_df[col], errors="coerce").fillna(0)
        for col in DISPLAY_STOCK_COLUMNS:
            result_df[col] = pd.to_numeric(result_df[col], errors="coerce").fillna(0).astype(int)
        result_df["할인가"] = result_df["할인가_num"].round(0).astype(int).apply(lambda x: "" if int(x) == 0 else str(int(x)))
        result_df["판매가"] = result_df["판매가_num"].round(0).astype(int).astype(str)
        result_df["총 판매가"] = result_df["총 판매가_num"].round(0).astype(int).astype(str)
        result_df["총 수입"] = result_df["총 수입_num"].round(0).astype(int).astype(str)
        result_df["총 KRW"] = result_df["총 KRW_num"].round(0).astype(int).astype(str)
        result_df["총 공급가"] = result_df["총 공급가_num"].round(0).astype(int).astype(str)
        result_df["총 수수료"] = result_df["총 수수료_num"].round(0).astype(int).astype(str)
        result_df["마진"] = result_df["마진_num"].round(0).astype(int).astype(str)
        result_df["수량"] = result_df["수량_num"].apply(lambda x: "" if int(x) == 0 else str(int(x)))
        for col in DISPLAY_STOCK_COLUMNS:
            result_df[col] = result_df[col].apply(lambda x: "" if int(x) == 0 else str(int(x)))
        result_df = result_df[OUTPUT_COLUMNS]
    else:
        result_df = pd.DataFrame(columns=OUTPUT_COLUMNS)
    store_order = {name: idx for idx, name in enumerate(ALLOC_PRIORITY + ["재고없음"])}
    result_df["_sort"] = result_df["매장명"].map(store_order).fillna(999).astype(int)
    result_df = result_df.sort_values(by=["_sort", "매장명", "운송장번호", "코드"], ascending=[True, True, True, True]).drop(columns=["_sort"]).reset_index(drop=True)
    return result_df


def upload_to_google_sheet(df: pd.DataFrame):
    if df.empty:
        raise DataValidationError("업로드할 데이터가 없습니다.")
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": [df.columns.tolist()] + df.fillna("").astype(str).values.tolist(),
        "append": True,
        "clear": False,
    }
    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            log(f"업로드 완료 -> {TARGET_SHEET_NAME} / {len(df)}행")
            return
        except Exception as e:
            last_error = e
            log(f"업로드 실패 ({attempt}/{RETRIES}) -> {e}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def main():
    start_time = time.perf_counter()
    log("KREAM 추가 리스트 다운로드 시작")
    source_df, size_map_df, stock_raw_df = load_all_sources()
    stock_df = load_stock_prepare(stock_raw_df)
    order_df = prepare_order_source(source_df, size_map_df, stock_df)
    output_df = auto_allocate(order_df, stock_df)
    log(f"출력 데이터 생성 완료 -> {output_df.shape[0]}행")
    upload_to_google_sheet(output_df)
    elapsed = time.perf_counter() - start_time
    log(f"완료 ({elapsed:.2f}초)")


if __name__ == "__main__":
    main()
