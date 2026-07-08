# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from io import StringIO
import time

import pandas as pd
import requests


SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSfBvby00YLSVYN-dPya7lNeGxtvDHfDDFiW0FwrhW3dHpIIYvupt2yW-t-QNZQhlRjJ98dHIWdNaMC/pub?gid=0&single=true&output=csv"
STOCK_TRANSFORM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"
SUPPLY_RATE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTNvyacPTFHHaoLymTqdhyO9rQErpoNFATDJJA_WbSCMYNrqjCaqPS_lUA9JQEqyo6PWWiKO-3Iwfwx/pub?gid=0&single=true&output=csv"
SQUAREONE_LOCATION_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSZ4Mgu9j6y27nLBYU8gAhDTfy4eMpvBgvs3oorR3BUCpcgoyf6Z1SllaqsFyos8LcH5DfxoUsN4NYG/pub?gid=289091756&single=true&output=csv"
FILA_PAJU_LOCATION_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSA47SgFq9QQPg0D3AlBnpJX6q7Yx_Dh66E1ID9MlXTahJjL0FmFVtPgyTEtj4iVj7PvRkCUoCgbjkd/pub?gid=1813802704&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzjGo2ZxLTaaIyUttkcrSThwS6mZA6FB9nw47yJSghFScG-SgS1W7j_JZTykWoXr8kV/exec"
SPREADSHEET_ID = "1jHdxOK_Y9MrLyTZD1Fj62g21zCHG0LL9N_X72HvWz2k"
TARGET_SHEET_NAME = "OUT"
SQUAREONE_LOCATION_SHEET_NAME = "스퀘어원 제품 위치"
FILA_PAJU_LOCATION_SHEET_NAME = "휠라 파주 제품 위치"

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
STOCK_REFERENCE_COLUMNS = (
    OFFICE_ALLOC_PRIORITY
    + SECONDARY_DYNAMIC_STORES
    + SECONDARY_FALLBACK_STORES
    + [FILA_FIXED_STORE, RETURN_FALLBACK_STORE, FINAL_FALLBACK_STORE]
)
VALID_STORE_VALUES = STOCK_REFERENCE_COLUMNS + ["재고없음"]
STORE_SORT_ORDER = VALID_STORE_VALUES
LOOKUP_STOCK_COLUMNS = ["현재고", "사무실"] + STOCK_REFERENCE_COLUMNS
DISPLAY_STOCK_COLUMNS = STOCK_REFERENCE_COLUMNS

OUTPUT_COLUMNS = [
    "날짜",
    "플랫폼",
    "브랜드",
    "코드",
    "품번",
    "사이즈",
    "수량",
    "매장명",
    "최초가",
    "할인가",
    "공급율",
    "판매가",
    "공급가",
    "마진",
] + DISPLAY_STOCK_COLUMNS
SOURCE_COLUMN_GROUPS = [
    ("A", "C", "D"),
    ("G", "I", "J"),
    ("M", "O", "P"),
    ("S", "U", "V"),
    ("Y", "AA", "AB"),
]

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 120)
RETRIES = 3
RETRY_SLEEP_SEC = 2


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[BRAND] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        try:
            number = float(text)
            if number.is_integer():
                return str(int(number))
        except Exception:
            pass
    return text


def clean_item_no(value) -> str:
    return clean_text(value).replace("-", "")


def compact_code_text(value) -> str:
    return clean_text(value).replace("-", "")


def to_number(value, default=0):
    text = clean_text(value).replace(",", "").replace("%", "")
    if text == "":
        return default
    number = pd.to_numeric(text, errors="coerce")
    if pd.isna(number):
        return default
    return float(number)


def to_display_number(value) -> str:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return ""
    number = float(number)
    if number == 0:
        return ""
    if number.is_integer():
        return str(int(number))
    return f"{number:.4f}".rstrip("0").rstrip(".")


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
            log(f"CSV 다운로드 실패 ({attempt}/{RETRIES}) -> {e}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def fetch_csv_job(name: str, url: str) -> tuple[str, str]:
    return name, fetch_csv_text(make_session(), url)


def download_csv_texts(url_map: dict[str, str]) -> dict[str, str]:
    results = {}
    with ThreadPoolExecutor(max_workers=min(len(url_map), 5)) as executor:
        future_map = {
            executor.submit(fetch_csv_job, name, url): name
            for name, url in url_map.items()
        }
        for future in as_completed(future_map):
            name, text = future.result()
            results[name] = text
    return results


def excel_col_to_index(col: str) -> int:
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def load_selected_excel_cols(df_raw: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    use_cols = [excel_col_to_index(col) for col in cols]
    if df_raw.shape[1] <= max(use_cols):
        raise DataValidationError(
            f"참조 시트 열 개수가 부족합니다. 필요한 최대 열 인덱스는 {max(use_cols) + 1}인데 실제는 {df_raw.shape[1]}열입니다."
        )
    df = df_raw.iloc[:, use_cols].copy().fillna("")
    df.columns = cols
    return df


def format_discount_percent(value) -> str:
    text = clean_text(value)
    if text == "":
        return ""
    number = pd.to_numeric(text.replace("%", ""), errors="coerce")
    if pd.isna(number):
        return text
    number = float(number)
    if number <= 1:
        number *= 100
    if number.is_integer():
        return f"{int(number)}%"
    return f"{number:.2f}".rstrip("0").rstrip(".") + "%"


def normalize_brand(value) -> str:
    text = clean_text(value).upper()
    if "NORTH FACE" in text or "NORTHFACE" in text:
        if "DC" in text:
            return "THE NORTH FACE (DC)"
        return "THE NORTH FACE"
    if "NIKE" in text or "나이키" in text:
        return "NIKE"
    if "ASICS" in text or "아식스" in text:
        return "ASICS"
    if "CONVERSE" in text or "컨버스" in text:
        return "CONVERSE"
    if "ADIDAS" in text:
        return "ADIDAS"
    if "FILA" in text or "휠라" in text:
        return "FILA"
    if "PUMA" in text or "푸마" in text:
        return "PUMA"
    return clean_text(value)


def load_supply_rate_lookup(csv_text: str) -> dict[str, float]:
    df = pd.read_csv(StringIO(csv_text), dtype=str, keep_default_na=False).fillna("")
    if "품번" not in df.columns or "공급율" not in df.columns:
        raise DataValidationError("공급율 시트에 '품번' 또는 '공급율' 열이 없습니다.")

    lookup = {}
    for item_no, rate_value in df[["품번", "공급율"]].itertuples(index=False, name=None):
        key = clean_text(item_no)
        compact_key = compact_code_text(item_no)
        rate_num = to_number(rate_value, default=None)
        if rate_num is not None:
            rate_num = float(rate_num)
            if rate_num > 1:
                rate_num = rate_num / 100
        if key and rate_num is not None and key not in lookup:
            lookup[key] = float(rate_num)
        if compact_key and rate_num is not None and compact_key not in lookup:
            lookup[compact_key] = float(rate_num)
    return lookup


def format_discount_value(value) -> str:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return ""
    number = float(number)
    if number == 0:
        return ""
    percent_value = number * 100 if number <= 1 else number
    if float(percent_value).is_integer():
        return f"{int(percent_value)}%"
    return f"{percent_value:.2f}".rstrip("0").rstrip(".") + "%"


def make_transform_data(csv_text: str) -> tuple[dict[str, dict[str, str]], dict[str, str], dict[str, dict[str, int]]]:
    df_raw = pd.read_csv(StringIO(csv_text), dtype=str, keep_default_na=False).fillna("")
    selected = load_selected_excel_cols(df_raw, ["A", "D", "E", "H", "M", "N", "O"])

    named_df = df_raw.copy()
    for col, legacy_col in LEGACY_STORE_ALIASES.items():
        if col not in named_df.columns and legacy_col in named_df.columns:
            named_df[col] = named_df[legacy_col]
    for col in ["코드", "품번", "최초가", "할인율", "할인가", "공급가"]:
        if col not in named_df.columns:
            named_df[col] = ""
    for col in LOOKUP_STOCK_COLUMNS:
        if col not in named_df.columns:
            named_df[col] = ""

    item_detail_lookup = {}
    variant_code_to_size = {}
    stock_lookup = {}

    for row_idx, (brand, final_item_no, converted_item_no, final_size, code1, code2, code3) in enumerate(
        zip(selected["A"], selected["D"], selected["E"], selected["H"], selected["M"], selected["N"], selected["O"])
    ):
        key = clean_item_no(converted_item_no)
        final_item = clean_text(final_item_no)
        raw_brand = normalize_brand(brand)

        named_row = named_df.iloc[row_idx] if row_idx < len(named_df) else None
        first_price = clean_text(named_row.get("최초가", "")) if named_row is not None else ""
        discount_rate = format_discount_percent(named_row.get("할인율", "")) if named_row is not None else ""
        sale_price = clean_text(named_row.get("할인가", "")) if named_row is not None else ""
        supply_price = clean_text(named_row.get("공급가", "")) if named_row is not None else ""

        detail = {
            "브랜드": raw_brand,
            "최종 품번": final_item,
            "최초가": first_price,
            "할인율": discount_rate,
            "할인가": sale_price,
            "공급가": supply_price,
        }
        for lookup_key in {key, clean_item_no(final_item)}:
            if lookup_key and lookup_key not in item_detail_lookup:
                item_detail_lookup[lookup_key] = detail

        normalized_final_size = clean_text(final_size)
        code_value = clean_text(named_row.get("코드", "")) if named_row is not None else ""
        stock_entry = {}
        if named_row is not None:
            for col in LOOKUP_STOCK_COLUMNS:
                stock_entry[col] = int(to_number(named_row.get(col, ""), default=0))

        if code_value and code_value not in stock_lookup:
            stock_lookup[code_value] = stock_entry.copy()
        compact_code_value = compact_code_text(code_value)
        if compact_code_value and compact_code_value not in stock_lookup:
            stock_lookup[compact_code_value] = stock_entry.copy()

        for code in [code1, code2, code3]:
            code_key = clean_text(code)
            if code_key and code_key not in variant_code_to_size:
                variant_code_to_size[code_key] = normalized_final_size
            compact_key = compact_code_text(code)
            if compact_key and compact_key not in variant_code_to_size:
                variant_code_to_size[compact_key] = normalized_final_size

    return item_detail_lookup, variant_code_to_size, stock_lookup


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


def auto_allocate_rows(df: pd.DataFrame, stock_lookup: dict[str, dict[str, int]]) -> pd.DataFrame:
    alloc_rows = []

    for row_dict in df.to_dict("records"):
        final_code = clean_text(row_dict["최종 코드"])
        lookup_key = final_code if final_code in stock_lookup else compact_code_text(final_code)
        display_stock_entry = stock_lookup.get(lookup_key, {})
        stock_work_entry = display_stock_entry.copy()
        qty = int(to_number(row_dict["수량"], default=0))

        if qty <= 0:
            base_row = dict(row_dict)
            base_row["매장명"] = ""
            for col in DISPLAY_STOCK_COLUMNS:
                base_row[col] = int(display_stock_entry.get(col, 0))
            alloc_rows.append(base_row)
            continue

        allocations = []
        allocations.extend(allocate_by_priority(stock_work_entry, OFFICE_ALLOC_PRIORITY, qty))
        allocated_qty = sum(use_qty for _, use_qty in allocations)
        remain = qty - allocated_qty

        if remain > 0 and row_dict["브랜드"] == "FILA":
            allocations.extend(allocate_by_priority(stock_work_entry, [FILA_FIXED_STORE], remain))
            allocated_qty = sum(use_qty for _, use_qty in allocations)
            remain = qty - allocated_qty

        if remain > 0:
            allocations.extend(allocate_by_priority(stock_work_entry, SECONDARY_DYNAMIC_STORES, remain))
            allocated_qty = sum(use_qty for _, use_qty in allocations)
            remain = qty - allocated_qty

        if remain > 0:
            allocations.extend(allocate_by_priority(stock_work_entry, SECONDARY_FALLBACK_STORES, remain))
            allocated_qty = sum(use_qty for _, use_qty in allocations)
            remain = qty - allocated_qty

        if remain > 0:
            allocations.extend(allocate_by_priority(stock_work_entry, [RETURN_FALLBACK_STORE], remain))
            allocated_qty = sum(use_qty for _, use_qty in allocations)
            remain = qty - allocated_qty

        if remain > 0:
            allocations.extend(allocate_by_priority(stock_work_entry, [FINAL_FALLBACK_STORE], remain))
            allocated_qty = sum(use_qty for _, use_qty in allocations)
            remain = qty - allocated_qty

        if not allocations and remain > 0:
            base_row = dict(row_dict)
            base_row["매장명"] = "재고없음"
            for col in DISPLAY_STOCK_COLUMNS:
                base_row[col] = int(display_stock_entry.get(col, 0))
            alloc_rows.append(base_row)
            continue

        for store_name, use_qty in allocations:
            base_row = dict(row_dict)
            base_row["매장명"] = store_name
            base_row["수량"] = use_qty
            for col in DISPLAY_STOCK_COLUMNS:
                base_row[col] = int(display_stock_entry.get(col, 0))
            alloc_rows.append(base_row)

        if remain > 0:
            base_row = dict(row_dict)
            base_row["매장명"] = "재고없음"
            base_row["수량"] = remain
            for col in DISPLAY_STOCK_COLUMNS:
                base_row[col] = int(display_stock_entry.get(col, 0))
            alloc_rows.append(base_row)

    if not alloc_rows:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    return pd.DataFrame(alloc_rows)


def blank_zero_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    result = df.copy()
    for col in columns:
        if col not in result.columns:
            continue
        numeric = pd.to_numeric(result[col], errors="coerce").fillna(0)
        result[col] = numeric.round(0).astype(int).astype(str)
        result[col] = result[col].replace("0", "")
    return result


def sort_by_store(df: pd.DataFrame) -> pd.DataFrame:
    temp = df.copy()
    store_sort_lookup = {store: idx for idx, store in enumerate(STORE_SORT_ORDER)}
    temp["__store_sort__"] = temp["매장명"].map(store_sort_lookup).fillna(999).astype(int)
    temp = temp.sort_values(
        by=["__store_sort__", "매장명", "브랜드", "품번", "사이즈", "코드"],
        ascending=[True, True, True, True, True, True],
        kind="stable",
    ).drop(columns=["__store_sort__"]).reset_index(drop=True)
    return temp


def read_source_sheet(
    csv_text: str,
    item_detail_lookup: dict[str, dict[str, str]],
    variant_code_to_size: dict[str, str],
    stock_lookup: dict[str, dict[str, int]],
    supply_rate_lookup: dict[str, float],
) -> pd.DataFrame:
    df_raw = pd.read_csv(StringIO(csv_text), dtype=str, keep_default_na=False).fillna("")
    source_group_indexes = [
        tuple(excel_col_to_index(col) for col in group)
        for group in SOURCE_COLUMN_GROUPS
    ]
    source_frames = []
    for group_cols, group_indexes in zip(SOURCE_COLUMN_GROUPS, source_group_indexes):
        if df_raw.shape[1] <= max(group_indexes):
            continue
        group_df = df_raw.iloc[:, list(group_indexes)].copy()
        group_df.columns = ["원본 품번", "사이즈", "수량"]
        group_df["__source_group__"] = "/".join(group_cols)
        source_frames.append(group_df)

    if not source_frames:
        raise DataValidationError(
            f"원본 시트에서 읽을 수 있는 품번/사이즈/수량 열 그룹이 없습니다. 실제 열 수는 {df_raw.shape[1]}개입니다."
        )

    df = pd.concat(source_frames, ignore_index=True)
    for col in ["원본 품번", "사이즈", "수량"]:
        df[col] = df[col].map(clean_text)

    blank_mask = (df["원본 품번"] == "") & (df["사이즈"] == "") & (df["수량"] == "")
    df = df.loc[~blank_mask].drop(columns=["__source_group__"], errors="ignore").reset_index(drop=True)
    if df.empty:
        raise DataValidationError("업로드할 데이터가 없습니다.")

    df["날짜"] = datetime.now().strftime("%Y-%m-%d")
    df["품번 변환"] = df["원본 품번"].map(clean_item_no)
    matched_rows = [item_detail_lookup.get(key, {}) for key in df["품번 변환"]]
    matched_df = pd.DataFrame.from_records(matched_rows, index=df.index)

    for col in ["브랜드", "최종 품번", "최초가", "할인율", "할인가", "공급가"]:
        if col in matched_df.columns:
            df[col] = matched_df[col].fillna("").map(clean_text)
        else:
            df[col] = ""
    df["최종 품번"] = df["최종 품번"].where(
        df["최종 품번"].astype(str).str.strip() != "",
        df["품번 변환"],
    )

    df["코드"] = (df["최종 품번"] + df["사이즈"]).where(
        (df["최종 품번"] != "") & (df["사이즈"] != ""),
        "",
    )
    compact_code_series = df["코드"].map(compact_code_text)
    df["사이즈 변환"] = (
        df["코드"].map(variant_code_to_size)
        .fillna(compact_code_series.map(variant_code_to_size))
        .replace("", pd.NA)
        .fillna(df["사이즈"])
        .map(clean_text)
    )
    df["최종 코드"] = (df["최종 품번"] + df["사이즈 변환"]).where(
        (df["최종 품번"] != "") & (df["사이즈 변환"] != ""),
        "",
    )
    compact_item_series = df["최종 품번"].map(compact_code_text)
    df["공급율"] = df["최종 품번"].map(supply_rate_lookup)
    df["공급율"] = df["공급율"].fillna(compact_item_series.map(supply_rate_lookup))
    df["판매가"] = (
        pd.to_numeric(df["최초가"].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
        * pd.to_numeric(df["공급율"], errors="coerce").fillna(0)
    ).round(0)
    df["공급율"] = df["공급율"].apply(format_discount_value)
    df["공급가"] = df["공급가"].apply(to_display_number)
    df["판매가"] = df["판매가"].apply(to_display_number)

    allocated_df = auto_allocate_rows(df, stock_lookup)
    if allocated_df.empty:
        return allocated_df

    allocated_df["플랫폼"] = "브랜더"
    allocated_df["품번"] = allocated_df["최종 품번"]
    allocated_df["사이즈"] = allocated_df["사이즈 변환"]
    allocated_df["코드"] = allocated_df["최종 코드"]
    allocated_df["판매가 계산"] = (
        pd.to_numeric(allocated_df["판매가"], errors="coerce").fillna(0)
        * pd.to_numeric(allocated_df["수량"], errors="coerce").fillna(0)
    ).round(0).apply(to_display_number)
    allocated_df["공급가 계산"] = (
        pd.to_numeric(allocated_df["공급가"].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
        * pd.to_numeric(allocated_df["수량"], errors="coerce").fillna(0)
    ).round(0).apply(to_display_number)
    allocated_df["마진"] = (
        pd.to_numeric(allocated_df["판매가 계산"], errors="coerce").fillna(0)
        - pd.to_numeric(allocated_df["공급가 계산"], errors="coerce").fillna(0)
    ).round(0).apply(to_display_number)
    allocated_df["판매가"] = allocated_df["판매가 계산"]
    allocated_df["공급가"] = allocated_df["공급가 계산"]
    allocated_df = allocated_df.drop(columns=["판매가 계산", "공급가 계산", "할인율"], errors="ignore")

    numeric_blank_cols = ["수량"] + DISPLAY_STOCK_COLUMNS
    allocated_df = blank_zero_columns(allocated_df, numeric_blank_cols)
    allocated_df = sort_by_store(allocated_df)

    for col in OUTPUT_COLUMNS:
        if col not in allocated_df.columns:
            allocated_df[col] = ""

    return allocated_df[OUTPUT_COLUMNS]


def upload_to_google_sheet(df: pd.DataFrame, sheet_name: str, columns: list[str] | None = None):
    header = columns if columns is not None else df.columns.tolist()
    upload_df = df.copy()
    if columns is not None:
        for col in columns:
            if col not in upload_df.columns:
                upload_df[col] = ""
        upload_df = upload_df[columns]

    values = [header] + upload_df.fillna("").astype(str).values.tolist()
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": sheet_name,
        "values": values,
        "append": False,
        "clear": True,
    }

    last_error = None
    log(f"업로드 시작 -> {sheet_name} ({len(upload_df)}행)")
    for attempt in range(1, RETRIES + 1):
        try:
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            if not response.ok:
                body = (response.text or "").strip()
                raise RuntimeError(f"HTTP {response.status_code} / {body}")
            log(f"업로드 완료 -> {sheet_name}")
            log(f"응답: {(response.text or '').strip()}")
            return
        except Exception as e:
            last_error = e
            log(f"업로드 실패 ({sheet_name}, {attempt}/{RETRIES}) -> {e}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def main():
    log("브랜드 리스트 업로드 시작")
    csv_texts = download_csv_texts(
        {
            "source": SOURCE_CSV_URL,
            "transform": STOCK_TRANSFORM_CSV_URL,
            "supply_rate": SUPPLY_RATE_CSV_URL,
            "squareone_location": SQUAREONE_LOCATION_CSV_URL,
            "fila_paju_location": FILA_PAJU_LOCATION_CSV_URL,
        }
    )
    source_csv_text = csv_texts["source"]
    transform_csv_text = csv_texts["transform"]
    supply_rate_csv_text = csv_texts["supply_rate"]
    squareone_location_csv_text = csv_texts["squareone_location"]
    fila_paju_location_csv_text = csv_texts["fila_paju_location"]
    squareone_location_df = pd.read_csv(StringIO(squareone_location_csv_text), dtype=str, keep_default_na=False).fillna("")
    fila_paju_location_raw = pd.read_csv(StringIO(fila_paju_location_csv_text), dtype=str, keep_default_na=False).fillna("")
    fila_paju_location_df = load_selected_excel_cols(fila_paju_location_raw, ["E", "F"])
    item_detail_lookup, variant_code_to_size, stock_lookup = make_transform_data(transform_csv_text)
    supply_rate_lookup = load_supply_rate_lookup(supply_rate_csv_text)
    result_df = read_source_sheet(source_csv_text, item_detail_lookup, variant_code_to_size, stock_lookup, supply_rate_lookup)

    upload_jobs = [
        (result_df, TARGET_SHEET_NAME, OUTPUT_COLUMNS),
        (squareone_location_df, SQUAREONE_LOCATION_SHEET_NAME, None),
        (fila_paju_location_df, FILA_PAJU_LOCATION_SHEET_NAME, None),
    ]
    with ThreadPoolExecutor(max_workers=min(3, len(upload_jobs))) as executor:
        futures = [
            executor.submit(upload_to_google_sheet, df, sheet_name, columns)
            for df, sheet_name, columns in upload_jobs
        ]
        for future in as_completed(futures):
            future.result()

    log("작업 완료")


if __name__ == "__main__":
    main()
