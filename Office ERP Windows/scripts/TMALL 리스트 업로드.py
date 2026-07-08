# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor, as_completed
from io import StringIO
import time
from datetime import datetime

import pandas as pd
import requests


SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSZMdVGPE8cfhFb-81CL6IS-xkP2s9HToyX8EuKE2F-B7djTz-pt7-DAhyQ24yvRdSDCaXpDk36266b/pub?gid=0&single=true&output=csv"
STOCK_TRANSFORM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"
ONLINE_SIZE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=2089368433&single=true&output=csv"
EXCHANGE_RATE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=295228098&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwacw7gEGNv2wnhnpSzVeLThZz2pLKf3MLlRCRoB3HdzIHsZsFMHg1BDel5vpzK4pi5PA/exec"
SPREADSHEET_ID = "1C2QmFLdZJkebFfbcaZeR7VH_ku1Jc8wOJm_5xepjfXU"
TARGET_SHEET_NAME = "OUT"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 120)
RETRIES = 3
RETRY_SLEEP_SEC = 2

OFFICE_ALLOC_PRIORITY = [
    "사무실 - 사무실",
    "사무실 - 구월",
    "사무실 - 부천",
    "사무실 - 스퀘어원",
    "사무실 - S마켓",
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
ALLOC_PRIORITY = OFFICE_ALLOC_PRIORITY + DYNAMIC_STORE_PRIORITY + STATIC_FALLBACK_STORES + [RETURN_FALLBACK_STORE]
STORE_SORT_ORDER = ALLOC_PRIORITY + ["재고없음"]
STORE_ORDER_INDEX = {store: idx for idx, store in enumerate(STORE_SORT_ORDER)}
EMPTY_STOCK_ENTRY = {col: 0 for col in ALLOC_PRIORITY}

SOURCE_COLUMNS = ["A", "Q", "S", "V", "AB", "AC", "AJ", "AO"]
OUTPUT_COLUMNS = [
    "날짜",
    "플랫폼",
    "주문번호",
    "운송장",
    "뒤 4자리",
    "운송장 링크",
    "브랜드",
    "코드",
    "품번",
    "사이즈",
    "수량",
    "매장명",
    "할인가",
    "판매가",
    "수수료",
    "수입",
    "환율",
    "KRW",
    "공급가",
    "마진",
] + ALLOC_PRIORITY


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(message, flush=True)


def make_session():
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


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


def today_str() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def normalize_key(value) -> str:
    return clean_text(value).upper()


def to_int(value) -> int:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return 0
    return int(num)


def to_float(value) -> float:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return 0.0
    return float(num)


def format_number(value):
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return ""
    num = float(num)
    if num.is_integer():
        return str(int(num))
    return f"{num:.4f}".rstrip("0").rstrip(".")


def get_suffix(value, length: int = 4) -> str:
    text = clean_text(value)
    if text == "":
        return ""
    return text[-length:]


def make_empty_stock_entry() -> dict:
    return EMPTY_STOCK_ENTRY.copy()


def get_stock_entry(stock_lookup: dict, code: str) -> dict:
    return stock_lookup.setdefault(code, make_empty_stock_entry())


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
            log(f"다운로드 실패 ({attempt}/{RETRIES}): {e}")
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


def load_selected_excel_cols(df_raw: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    use_cols = [excel_col_to_index(col) for col in cols]
    ensure_min_columns(df_raw, use_cols, f"선택 열 {cols}")
    df = df_raw.iloc[:, use_cols].copy().fillna("")
    df.columns = cols
    return df


def make_transform_lookup(csv_text: str) -> tuple[dict[str, str], dict[str, str], dict[str, str], dict[str, str], dict[str, dict], dict[str, float], dict[str, float]]:
    df_raw = pd.read_csv(StringIO(csv_text), dtype=str).fillna("")
    selected = load_selected_excel_cols(
        df_raw,
        ["A", "B", "D", "E", "H", "M", "N", "O", "S", "T"]
    )

    item_lookup = {}
    brand_lookup = {}
    code_lookup = {}
    final_size_lookup = {}
    stock_lookup = {}
    supply_lookup = {}
    discount_lookup = {}
    for brand, code_value, item_no, converted_item_no, final_size_value, code1, code2, code3, discount_value, supply_value in zip(
        selected["A"], selected["B"], selected["D"], selected["E"], selected["H"], selected["M"], selected["N"], selected["O"], selected["S"], selected["T"]
    ):
        key = clean_text(converted_item_no)
        item_value = clean_text(item_no)
        brand_value = clean_text(brand)
        code_value = clean_text(code_value)
        final_size_value = clean_text(final_size_value)
        discount_num = to_float(discount_value)
        supply_num = to_float(supply_value)
        if key and key not in item_lookup:
            item_lookup[key] = item_value
        compact_key = compact_code_text(converted_item_no)
        if compact_key and compact_key not in item_lookup:
            item_lookup[compact_key] = item_value
        if item_value and item_value not in brand_lookup:
            brand_lookup[item_value] = brand_value
        compact_item_value = compact_code_text(item_value)
        if compact_item_value and compact_item_value not in brand_lookup:
            brand_lookup[compact_item_value] = brand_value
        if item_value and item_value not in discount_lookup:
            discount_lookup[item_value] = discount_num
        if compact_item_value and compact_item_value not in discount_lookup:
            discount_lookup[compact_item_value] = discount_num
        if item_value and item_value not in supply_lookup:
            supply_lookup[item_value] = supply_num / 1.1 if supply_num else 0.0
        if compact_item_value and compact_item_value not in supply_lookup:
            supply_lookup[compact_item_value] = supply_num / 1.1 if supply_num else 0.0
        for code_key in [code1, code2, code3]:
            normalized_code_key = clean_text(code_key)
            if normalized_code_key and normalized_code_key not in code_lookup:
                code_lookup[normalized_code_key] = code_value
            if normalized_code_key and normalized_code_key not in final_size_lookup:
                final_size_lookup[normalized_code_key] = final_size_value
            compact_code_key = compact_code_text(code_key)
            if compact_code_key and compact_code_key not in code_lookup:
                code_lookup[compact_code_key] = code_value
            if compact_code_key and compact_code_key not in final_size_lookup:
                final_size_lookup[compact_code_key] = final_size_value

    temp = df_raw.fillna("")
    code_idx = excel_col_to_index("B")
    if temp.shape[1] > code_idx:
        code_series = temp.iloc[:, code_idx].apply(clean_text)
        for row_idx, code_value in code_series.items():
            if not code_value or code_value in stock_lookup:
                continue
            row = temp.iloc[row_idx]
            stock_lookup[code_value] = {
                col: to_int(row[col]) if col in row.index else 0
                for col in ALLOC_PRIORITY
            }
            compact_code_value = compact_code_text(code_value)
            if compact_code_value and compact_code_value not in stock_lookup:
                stock_lookup[compact_code_value] = stock_lookup[code_value].copy()

    return item_lookup, brand_lookup, code_lookup, final_size_lookup, stock_lookup, supply_lookup, discount_lookup


def make_size_lookup(csv_text: str) -> dict[tuple[str, str], str]:
    df_raw = pd.read_csv(StringIO(csv_text), dtype=str).fillna("")

    if {"브랜드", "사이즈", "사이즈 변환"}.issubset(df_raw.columns):
        temp = df_raw[["브랜드", "사이즈", "사이즈 변환"]].copy()
    elif {"브랜드", "사이즈", "EU"}.issubset(df_raw.columns):
        temp = df_raw[["브랜드", "사이즈", "EU"]].copy()
        temp.columns = ["브랜드", "사이즈", "사이즈 변환"]
    else:
        if df_raw.shape[1] < 3:
            raise DataValidationError("온라인 사이즈표 열 수가 부족합니다.")
        temp = df_raw.iloc[:, :3].copy()
        temp.columns = ["브랜드", "사이즈", "사이즈 변환"]

    lookup = {}
    for brand, size, converted_size in zip(temp["브랜드"], temp["사이즈"], temp["사이즈 변환"]):
        key = (normalize_key(brand), normalize_key(size))
        value = clean_text(converted_size)
        if key[0] and key[1] and key not in lookup:
            lookup[key] = value
        compact_size = compact_code_text(size)
        if key[0] and compact_size:
            lookup.setdefault((key[0], compact_size.upper()), value)
    return lookup


def get_exchange_rate(csv_text: str) -> float:
    df_raw = pd.read_csv(StringIO(csv_text), dtype=str, header=None).fillna("")
    if df_raw.shape[0] < 3 or df_raw.shape[1] < 2:
        raise DataValidationError("환율 시트에서 B3 값을 읽을 수 없습니다.")
    return to_float(df_raw.iat[2, 1])


def get_dynamic_store_priority(stock_entry: dict) -> list[str]:
    base_order = {store: idx for idx, store in enumerate(DYNAMIC_STORE_PRIORITY)}
    return sorted(
        DYNAMIC_STORE_PRIORITY,
        key=lambda store: (-int(stock_entry.get(store, 0)), base_order[store]),
    )


def get_allocation_priority(stock_entry: dict) -> list[str]:
    return OFFICE_ALLOC_PRIORITY + get_dynamic_store_priority(stock_entry) + STATIC_FALLBACK_STORES + [RETURN_FALLBACK_STORE]


def get_order_candidate_stores(order_rows: list[dict], stock_lookup: dict) -> list[str]:
    base_order = {store: idx for idx, store in enumerate(DYNAMIC_STORE_PRIORITY)}
    dynamic_totals = {store: 0 for store in DYNAMIC_STORE_PRIORITY}

    for order_row in order_rows:
        stock_entry = get_stock_entry(stock_lookup, order_row["코드 변환"])
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
        code = row["코드 변환"]
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
        stock_entry = get_stock_entry(stock_lookup, row["코드 변환"])
        stock_entry[store] = int(stock_entry.get(store, 0)) - row["수량_int"]
        row_data = row["base_output"].copy()
        row_data["매장명"] = store
        allocated_rows.append(row_data)
    return allocated_rows


def sort_output_df(df: pd.DataFrame) -> pd.DataFrame:
    temp = df.copy()
    temp["매장명_순서"] = temp["매장명"].map(STORE_ORDER_INDEX).fillna(999).astype(int)
    temp = temp.sort_values(
        by=["매장명_순서", "매장명", "주문번호", "코드"],
        ascending=[True, True, True, True],
    ).drop(columns=["매장명_순서"]).reset_index(drop=True)
    return temp


def build_output_df(
    csv_text: str,
    transform_lookup: dict[str, str],
    brand_lookup: dict[str, str],
    size_lookup: dict[tuple[str, str], str],
    code_lookup: dict[str, str],
    final_size_lookup: dict[str, str],
    stock_lookup: dict[str, dict],
    supply_lookup: dict[str, float],
    discount_lookup: dict[str, float],
    exchange_rate: float,
) -> pd.DataFrame:
    df_raw = pd.read_csv(StringIO(csv_text), dtype=str).fillna("")
    selected = load_selected_excel_cols(df_raw, SOURCE_COLUMNS)

    item_series = selected["V"].apply(clean_text)
    split_item = item_series.str.split("-", n=1, expand=True)
    item_no_series = split_item[0].fillna("").apply(clean_text)
    size_series = split_item[1].fillna("").apply(clean_text) if split_item.shape[1] > 1 else pd.Series("", index=selected.index)
    converted_item_series = item_no_series.map(lambda value: transform_lookup.get(value, value))
    brand_series = converted_item_series.map(lambda value: brand_lookup.get(value, ""))
    converted_size_series = pd.Series(
        [
            size_lookup.get((normalize_key(brand), normalize_key(size)), "")
            for brand, size in zip(brand_series, size_series)
        ],
        index=selected.index,
    )
    converted_size_series = converted_size_series.where(
        converted_size_series.fillna("").astype(str).str.strip() != "",
        size_series,
    )
    raw_item_clean = item_no_series.fillna("").astype(str).str.strip()
    converted_size_clean = converted_size_series.fillna("").astype(str).str.strip()
    code_series = (
        raw_item_clean
        + converted_size_clean
    ).where(
        (raw_item_clean != "")
        & (converted_size_clean != ""),
        "",
    )
    converted_code_series = code_series.map(lambda value: code_lookup.get(clean_text(value), ""))
    final_size_series = code_series.map(lambda value: final_size_lookup.get(clean_text(value), ""))
    sale_price_series = selected["AJ"].apply(clean_text)
    fee_series = pd.Series([""] * len(selected), index=selected.index)
    income_series = sale_price_series.copy()
    discount_price_series = converted_item_series.apply(
        lambda value: format_number(discount_lookup.get(clean_text(value), 0.0))
    )
    krw_series = income_series.apply(lambda value: round(to_float(value) * exchange_rate) if clean_text(value) != "" else "")
    supply_series = converted_item_series.apply(
        lambda value: format_number(supply_lookup.get(clean_text(value), 0.0))
    )
    margin_series = pd.Series(
        [
            format_number(to_float(krw) - to_float(supply))
            if clean_text(krw) != "" or clean_text(supply) != ""
            else ""
            for krw, supply in zip(krw_series, supply_series)
        ],
        index=selected.index,
    )

    filtered_df = pd.DataFrame({
        "날짜": today_str(),
        "플랫폼": "TMALL",
        "주문번호": selected["A"].apply(clean_text),
        "운송장": selected["Q"].apply(clean_text),
        "운송장 링크": selected["S"].apply(clean_text),
        "품번": item_no_series,
        "품번 변환": converted_item_series,
        "브랜드": brand_series,
        "사이즈": size_series,
        "사이즈 변환": converted_size_series,
        "사이즈 변환 최종": final_size_series,
        "코드": code_series,
        "코드 변환": converted_code_series,
        "출고 상태": selected["AB"].apply(clean_text),
        "수량": selected["AC"].apply(clean_text),
        "판매가": sale_price_series,
        "수수료": fee_series,
        "수입": income_series,
        "할인가": discount_price_series,
        "환율": format_number(exchange_rate),
        "KRW": krw_series.apply(format_number),
        "공급가": supply_series,
        "마진": margin_series,
        "운송장 신청 여부": selected["AO"].apply(clean_text),
    })

    status_mask = filtered_df["출고 상태"].eq("待发货")
    request_mask = filtered_df["운송장 신청 여부"].eq("已完成")
    filtered_df = filtered_df[status_mask & request_mask].copy()

    non_empty_mask = filtered_df.astype(str).apply(lambda row: "".join(row).strip() != "", axis=1)
    filtered_df = filtered_df[non_empty_mask].reset_index(drop=True)

    if filtered_df.empty:
        raise DataValidationError("조건에 맞는 업로드 데이터가 없습니다.")

    prepared_rows = []
    for row in filtered_df.to_dict("records"):
        alloc_code = clean_text(row.get("코드 변환", "")) or clean_text(row.get("코드", ""))
        qty = to_int(row.get("수량", ""))
        stock_entry = get_stock_entry(stock_lookup, alloc_code) if alloc_code else make_empty_stock_entry()
        stock_snapshot = {col: stock_entry.get(col, 0) for col in ALLOC_PRIORITY}

        base_output = {
            "날짜": row.get("날짜", ""),
            "플랫폼": row.get("플랫폼", ""),
            "주문번호": row.get("주문번호", ""),
            "운송장": row.get("운송장", ""),
            "뒤 4자리": get_suffix(row.get("운송장", "")),
            "운송장 링크": row.get("운송장 링크", ""),
            "브랜드": row.get("브랜드", ""),
            "코드": alloc_code,
            "품번": row.get("품번 변환", ""),
            "사이즈": row.get("사이즈 변환 최종", ""),
            "수량": str(qty),
            "매장명": "",
            "할인가": row.get("할인가", ""),
            "판매가": row.get("판매가", ""),
            "수수료": row.get("수수료", ""),
            "수입": row.get("수입", ""),
            "환율": row.get("환율", ""),
            "KRW": row.get("KRW", ""),
            "공급가": row.get("공급가", ""),
            "마진": row.get("마진", ""),
        }
        base_output.update({
            col: ("" if stock_snapshot.get(col, 0) == 0 else stock_snapshot.get(col, 0))
            for col in ALLOC_PRIORITY
        })
        prepared_rows.append({
            "주문번호": clean_text(row.get("주문번호", "")),
            "코드 변환": alloc_code,
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
            stock_entry = get_stock_entry(stock_lookup, order_row["코드 변환"]) if order_row["코드 변환"] else make_empty_stock_entry()
            allocations = allocate_qty(stock_entry, order_row["수량_int"])
            for store, alloc_qty in allocations:
                row_data = order_row["base_output"].copy()
                row_data["매장명"] = store
                row_data["수량"] = str(alloc_qty)
                transformed_rows.append(row_data)

    result_df = pd.DataFrame(transformed_rows, columns=OUTPUT_COLUMNS)
    return sort_output_df(result_df)


def upload_to_google_sheet(df: pd.DataFrame):
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": [df.columns.tolist()] + df.fillna("").astype(str).values.tolist(),
    }

    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            response_text = (response.text or "").strip()
            log(f"업로드 완료: {TARGET_SHEET_NAME} / {len(df)}행")
            log(f"응답: {response_text}")
            return
        except Exception as e:
            last_error = e
            log(f"업로드 실패 ({attempt}/{RETRIES}): {e}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def main():
    start_time = time.perf_counter()
    log("TMALL 관련 시트 다운로드 시작")
    csv_texts = fetch_all_csv_texts({
        "source": SOURCE_CSV_URL,
        "transform": STOCK_TRANSFORM_CSV_URL,
        "size": ONLINE_SIZE_CSV_URL,
        "exchange": EXCHANGE_RATE_CSV_URL,
    })

    transform_lookup, brand_lookup, code_lookup, final_size_lookup, stock_lookup, supply_lookup, discount_lookup = make_transform_lookup(csv_texts["transform"])
    size_lookup = make_size_lookup(csv_texts["size"])
    exchange_rate = get_exchange_rate(csv_texts["exchange"])

    log("TMALL 원본 시트 가공 시작")
    output_df = build_output_df(
        csv_texts["source"],
        transform_lookup,
        brand_lookup,
        size_lookup,
        code_lookup,
        final_size_lookup,
        stock_lookup,
        supply_lookup,
        discount_lookup,
        exchange_rate,
    )

    log("TMALL OUT 업로드 시작")
    upload_to_google_sheet(output_df)
    elapsed = time.perf_counter() - start_time
    log(f"완료 (총 소요 시간: {elapsed:.2f}초)")


if __name__ == "__main__":
    main()
