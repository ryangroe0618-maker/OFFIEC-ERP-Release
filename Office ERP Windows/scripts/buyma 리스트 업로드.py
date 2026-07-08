# -*- coding: utf-8 -*-

from io import StringIO
from pathlib import Path
import importlib.util
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests


BUYMA_SHIPMENT_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR3-K2ZY0WEoktfTZFCLZEmlPJr_Pub9CwLvZDQSaE2ySGCdXNORnu5Wn6-Y-fzHNRlnNIvw01e4tyx/pub?gid=1602297321&single=true&output=csv"
STOCK_TRANSFORM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"
EXCHANGE_RATE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=295228098&single=true&output=csv"
SQUAREONE_LOCATION_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSZ4Mgu9j6y27nLBYU8gAhDTfy4eMpvBgvs3oorR3BUCpcgoyf6Z1SllaqsFyos8LcH5DfxoUsN4NYG/pub?gid=289091756&single=true&output=csv"
FILA_PAJU_LOCATION_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSA47SgFq9QQPg0D3AlBnpJX6q7Yx_Dh66E1ID9MlXTahJjL0FmFVtPgyTEtj4iVj7PvRkCUoCgbjkd/pub?gid=1813802704&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzzO_bCM41A2KKtLM_ZA7ROIEJ_HJxU5yhk3bQ8DsSzF0XfvEbrenLQozE3MXiX1hVc/exec"
SPREADSHEET_ID = "1P2JmY6iwEf7PF4_yC8TK3xbUfRWEMMZt6ucuGtXZ3-g"
TARGET_SHEET_NAME = "BUYMA"
PLATFORM_INFO_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRWhzq2rt_O-YnOwzWnE8L_d8--NBu-EMmDkxnnVy-oATAfCjDIX976b973bY2MiO8gYPTE5WCk-tVv/pub?gid=733480714&single=true&output=csv"
SQUAREONE_LOCATION_SHEET_NAME = "스퀘어원 제품 위치"
FILA_PAJU_LOCATION_SHEET_NAME = "휠라 파주 제품 위치"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 300)
RETRIES = 3
RETRY_SLEEP_SEC = 2
EXCHANGE_RATE_RETRIES = 12
EXCHANGE_RATE_SLEEP_SEC = 5

BUYMA_USECOLS = ["A", "D", "F", "G", "I", "P"]
BUYMA_DATE_COLUMN = "날짜"
BUYMA_SOURCE_COLUMNS = [BUYMA_DATE_COLUMN, "품번", "사이즈", "수량", "총 판매가", "주문번호"]


class BuymaUploadError(ValueError):
    pass


def log(message: str):
    print(f"[BUYMA 리스트 업로드] {message}", flush=True)


def load_kashion_helpers():
    module_path = Path(__file__).with_name("KASHION 리스트 업로드.py")
    spec = importlib.util.spec_from_file_location("kashion_list_upload_helpers", module_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


KASHION = load_kashion_helpers()
OUTPUT_COLUMNS = [
    "날짜",
    "원거래 날짜",
    "플랫폼",
    "주문번호",
    "뒤 4자리",
    "브랜드",
    "코드",
    "품번",
    "사이즈",
    "수량",
    "매장명",
    "할인가",
    "총 판매가",
    "총 수수료",
    "총 수입",
    "환율",
    "총 KRW",
    "총 공급가",
    "마진",
    "내역",
] + KASHION.ALLOC_PRIORITY


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(url: str) -> str:
    last_error = None
    session = make_session()
    for attempt in range(1, RETRIES + 1):
        try:
            response = session.get(url, timeout=DOWNLOAD_TIMEOUT)
            response.raise_for_status()
            response.encoding = "utf-8"
            text = response.text
            if not text.strip():
                raise BuymaUploadError("CSV 응답이 비어 있습니다.")
            return text
        except Exception as exc:
            last_error = exc
            log(f"다운로드 실패 ({attempt}/{RETRIES}): {exc}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def clean_text(value) -> str:
    return KASHION.clean_text(value)


def to_number_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.fillna("")
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₩", "", regex=False)
        .str.strip(),
        errors="coerce",
    ).fillna(0)


def parse_number(value) -> float:
    text = clean_text(value).replace(",", "")
    try:
        return float(text)
    except Exception:
        raise BuymaUploadError(f"숫자로 읽을 수 없는 값입니다: {value}")


def excel_col_to_index(column: str) -> int:
    result = 0
    for char in column.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1


def read_buyma_shipment_df(csv_text: str) -> pd.DataFrame:
    usecols = [excel_col_to_index(column) for column in BUYMA_USECOLS]
    df = pd.read_csv(StringIO(csv_text), header=2, usecols=usecols, dtype=str).fillna("")
    if df.shape[1] != len(BUYMA_SOURCE_COLUMNS):
        raise BuymaUploadError(
            f"BUYMA 출고 시트 열 수가 맞지 않습니다. 기대 {len(BUYMA_SOURCE_COLUMNS)}개, 실제 {df.shape[1]}개"
        )
    df.columns = BUYMA_SOURCE_COLUMNS
    for column in BUYMA_SOURCE_COLUMNS:
        df[column] = df[column].apply(clean_text)
    today = time.strftime("%Y-%m-%d")
    sale_dates = pd.to_datetime(df[BUYMA_DATE_COLUMN], errors="coerce").dt.strftime("%Y-%m-%d")
    df = df[sale_dates.eq(today) | df[BUYMA_DATE_COLUMN].eq(today)].copy()
    df = df[df["품번"].ne("")].copy()
    return df.reset_index(drop=True)


def read_exchange_rate(csv_text: str) -> str:
    df = pd.read_csv(StringIO(csv_text), header=None, dtype=str).fillna("")
    if df.shape[0] < 4 or df.shape[1] < 2:
        raise BuymaUploadError("환율 시트 B4 셀을 읽을 수 없습니다.")
    rate = clean_text(df.iloc[3, 1])
    parse_number(rate)
    return rate


def load_selected_excel_cols(df: pd.DataFrame, excel_cols: list[str]) -> pd.DataFrame:
    indexes = [excel_col_to_index(col) for col in excel_cols]
    available_indexes = [idx for idx in indexes if idx < df.shape[1]]
    if len(available_indexes) != len(indexes):
        raise BuymaUploadError(f"선택할 열이 부족합니다: {excel_cols}, 실제 {df.shape[1]}열")
    return df.iloc[:, available_indexes].copy()


def fetch_csv_job(name: str, url: str) -> tuple[str, str]:
    return name, fetch_csv_text(url)


def fetch_all_csv_texts(url_map: dict[str, str]) -> dict[str, str]:
    results = {}
    with ThreadPoolExecutor(max_workers=min(len(url_map), 6)) as executor:
        future_map = {
            executor.submit(fetch_csv_job, name, url): name
            for name, url in url_map.items()
        }
        for future in as_completed(future_map):
            name, text = future.result()
            results[name] = text
    return results


def fetch_exchange_rate() -> str:
    last_error = None
    for attempt in range(1, EXCHANGE_RATE_RETRIES + 1):
        try:
            return read_exchange_rate(fetch_csv_text(EXCHANGE_RATE_URL))
        except Exception as exc:
            last_error = exc
            log(f"환율 읽기 실패 ({attempt}/{EXCHANGE_RATE_RETRIES}): {exc}")
            if attempt < EXCHANGE_RATE_RETRIES:
                time.sleep(EXCHANGE_RATE_SLEEP_SEC)
    raise last_error


def get_order_suffix(value, length: int = 4) -> str:
    text = clean_text(value)
    if not text:
        return ""
    suffix = text[-length:]
    if suffix.startswith("0"):
        return f"'{suffix}"
    return suffix


def get_first_column_series(df: pd.DataFrame, column_names: list[str]) -> pd.Series:
    for column_name in column_names:
        if column_name in df.columns:
            return df[column_name].fillna("").astype(str)
    return pd.Series([""] * len(df), index=df.index, dtype=str)


def fetch_platform_reference_df() -> pd.DataFrame:
    try:
        csv_text = fetch_csv_text(PLATFORM_INFO_CSV_URL)
        df = pd.read_csv(
            StringIO(csv_text),
            dtype=str,
            header=0,
            keep_default_na=False,
            skip_blank_lines=False,
        ).fillna("")
    except Exception as exc:
        log(f"플랫폼 시트를 읽지 못해 구하는 중 이어붙기를 건너뜁니다: {exc}")
        return pd.DataFrame()

    df.columns = [clean_text(column) for column in df.columns]
    return df


def make_delivery_lookup(platform_df: pd.DataFrame) -> dict[str, str]:
    lookup = {}
    if platform_df.empty:
        return lookup

    temp = platform_df.fillna("").copy()
    temp.columns = [clean_text(column) for column in temp.columns]

    if "주문번호" in temp.columns:
        order_series = get_first_column_series(temp, ["주문번호"]).apply(clean_text)
    elif temp.shape[1] > 2:
        order_series = temp.iloc[:, 2].apply(clean_text)
    else:
        return lookup

    if "내역" in temp.columns or "택배사" in temp.columns:
        status_series = get_first_column_series(temp, ["내역", "택배사"]).apply(clean_text)
    elif temp.shape[1] > 18:
        status_series = temp.iloc[:, 18].apply(clean_text)
    else:
        return lookup

    platform_series = get_first_column_series(temp, ["플랫폼"]).apply(clean_text).str.upper()

    for order_no, status, platform in zip(order_series, status_series, platform_series):
        if platform and platform != "BUYMA":
            continue
        normalized_order_no = KASHION.normalize_key(order_no)
        if normalized_order_no and normalized_order_no not in lookup:
            lookup[normalized_order_no] = status

    return lookup


def build_searching_df(existing_df: pd.DataFrame) -> pd.DataFrame:
    if existing_df.empty or "내역" not in existing_df.columns:
        return pd.DataFrame()

    temp = existing_df.copy()
    platform_series = get_first_column_series(temp, ["플랫폼"]).apply(clean_text).str.upper()
    if not platform_series.empty:
        temp = temp[platform_series.eq("BUYMA")].copy()
    if temp.empty:
        return pd.DataFrame()

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
        if "날짜" in temp.columns
        else {}
    )

    newest_first = temp.iloc[::-1].copy()
    newest_first = newest_first[~newest_first["__latest_key"].duplicated()].copy()
    newest_first["__original_date"] = (
        newest_first["__latest_key"]
        .map(original_date_lookup)
        .fillna(get_first_column_series(newest_first, ["원거래 날짜", "날짜"]))
    )

    status_series = get_first_column_series(newest_first, ["내역"]).apply(clean_text)
    searching_values = {"구하는 중", "구하는중", "调货中"}
    return newest_first[status_series.isin(searching_values)].copy().reset_index(drop=True)


def build_searching_out_df(searching_df: pd.DataFrame) -> pd.DataFrame:
    out_df = pd.DataFrame("", index=searching_df.index, columns=OUTPUT_COLUMNS)
    if searching_df.empty:
        return out_df

    for output_column in OUTPUT_COLUMNS:
        out_df[output_column] = get_first_column_series(searching_df, [output_column])

    out_df["원거래 날짜"] = get_first_column_series(
        searching_df,
        ["__original_date", "원거래 날짜", "날짜"],
    ).reset_index(drop=True)
    out_df["날짜"] = time.strftime("%Y-%m-%d")
    out_df["플랫폼"] = out_df["플랫폼"].replace("", "BUYMA")
    out_df["내역"] = "구하는 중"
    out_df["매장명"] = out_df["매장명"].replace("", "재고없음")
    return out_df.fillna("").astype(str).reset_index(drop=True)


def proportional_number(total: float, alloc_qty: int, total_qty: int) -> int:
    if total_qty <= 0:
        return 0
    return int(round(float(total) * alloc_qty / total_qty))


def allocated_row(order_row: dict, store: str, alloc_qty: int) -> dict:
    row = order_row["base_output"].copy()
    total_qty = max(int(order_row["수량_int"]), 1)
    row["매장명"] = store
    row["수량"] = str(alloc_qty)
    for column in ["총 판매가", "총 수수료", "총 수입", "총 KRW", "총 공급가", "마진"]:
        row[column] = str(proportional_number(order_row[f"{column}_num"], alloc_qty, total_qty))
    return row


def build_prepared_rows(
    buyma_df: pd.DataFrame,
    stock_df: pd.DataFrame,
    exchange_rate: str,
    delivery_lookup: dict[str, str],
) -> list[dict]:
    transform_lookup, stock_lookup, supply_lookup, discount_lookup = KASHION.make_stock_lookups(stock_df)
    variant_code_to_size = KASHION.build_variant_code_to_size_map(stock_df)
    exchange_rate_num = parse_number(exchange_rate)
    prepared_rows = []

    for _, source_row in buyma_df.iterrows():
        item_no = clean_text(source_row.get("품번", ""))
        raw_size = clean_text(source_row.get("사이즈", ""))
        stock_info = transform_lookup.get(KASHION.normalize_key(item_no), {})
        converted_item_no = stock_info.get("품번_변환", "") or item_no
        brand = stock_info.get("브랜드", "")
        variant_code = KASHION.make_code(converted_item_no, raw_size)
        output_size = ""
        for key in KASHION.code_lookup_keys(variant_code):
            output_size = variant_code_to_size.get(key, "")
            if output_size:
                break
        output_size = output_size or raw_size
        code = KASHION.make_code(converted_item_no, output_size)
        qty = int(to_number_series(pd.Series([source_row.get("수량", "")])).iloc[0])
        total_sale_price = float(to_number_series(pd.Series([source_row.get("총 판매가", "")])).iloc[0])
        fee = round(total_sale_price * 0.077 + total_sale_price * 0.05 + (6000 / exchange_rate_num))
        income = round(total_sale_price - fee)
        krw = round(income * exchange_rate_num)
        supply_unit = KASHION.to_float(supply_lookup.get(converted_item_no, 0.0)) / 1.1
        supply_price = round(supply_unit * qty)
        margin = round(krw - supply_price)
        discount_price = KASHION.to_float(discount_lookup.get(converted_item_no, 0.0))
        order_no = clean_text(source_row.get("주문번호", ""))
        normalized_order_no = KASHION.normalize_key(order_no)

        base_output = {
            "날짜": clean_text(source_row.get(BUYMA_DATE_COLUMN, "")) or time.strftime("%Y-%m-%d"),
            "원거래 날짜": "",
            "플랫폼": "BUYMA",
            "주문번호": order_no,
            "뒤 4자리": get_order_suffix(order_no),
            "브랜드": brand,
            "코드": code,
            "품번": converted_item_no,
            "사이즈": output_size,
            "수량": str(qty),
            "매장명": "",
            "할인가": KASHION.format_number(discount_price),
            "총 판매가": str(round(total_sale_price)),
            "총 수수료": str(fee),
            "총 수입": str(income),
            "환율": exchange_rate,
            "총 KRW": str(krw),
            "총 공급가": str(supply_price),
            "마진": str(margin),
            "내역": delivery_lookup.get(normalized_order_no, ""),
        }
        stock_entry = KASHION.get_stock_entry(stock_lookup, code)
        stock_snapshot = {
            col: int(stock_entry.get(col, 0))
            for col in KASHION.ALLOC_PRIORITY
        }
        base_output.update(
            {
                col: ("" if stock_snapshot.get(col, 0) == 0 else stock_snapshot.get(col, 0))
                for col in KASHION.ALLOC_PRIORITY
            }
        )
        prepared_rows.append(
            {
                "주문번호": order_no,
                "코드": code,
                "수량_int": qty,
                "base_output": base_output,
                "총 판매가_num": total_sale_price,
                "총 수수료_num": fee,
                "총 수입_num": income,
                "총 KRW_num": krw,
                "총 공급가_num": supply_price,
                "마진_num": margin,
            }
        )

    return prepared_rows, stock_lookup


def apply_auto_allocation(prepared_rows: list[dict], stock_lookup: dict) -> pd.DataFrame:
    transformed_rows = []
    if not prepared_rows:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    prepared_df = pd.DataFrame(prepared_rows)
    for _, group_df in prepared_df.groupby("주문번호", sort=False):
        order_rows = group_df.to_dict("records")
        assigned = False

        for store in KASHION.get_order_candidate_stores(order_rows, stock_lookup):
            if KASHION.can_allocate_order_to_store(order_rows, stock_lookup, store):
                for row in order_rows:
                    stock_entry = KASHION.get_stock_entry(stock_lookup, row["코드"])
                    stock_entry[store] = int(stock_entry.get(store, 0)) - row["수량_int"]
                    transformed_rows.append(allocated_row(row, store, row["수량_int"]))
                assigned = True
                break

        if assigned:
            continue

        for row in order_rows:
            stock_entry = KASHION.get_stock_entry(stock_lookup, row["코드"])
            for store, alloc_qty in KASHION.allocate_qty(stock_entry, row["수량_int"]):
                transformed_rows.append(allocated_row(row, store, alloc_qty))

    output_df = pd.DataFrame(transformed_rows, columns=OUTPUT_COLUMNS)
    if output_df.empty:
        return output_df
    output_df["매장명_순서"] = output_df["매장명"].map(KASHION.STORE_ORDER_INDEX).fillna(999).astype(int)
    output_df = output_df.sort_values(
        by=["매장명_순서", "매장명", "주문번호", "품번", "사이즈"],
        ascending=[True, True, True, True, True],
    ).drop(columns=["매장명_순서"]).reset_index(drop=True)
    return output_df


def build_output_df(include_existing_searching: bool = True) -> pd.DataFrame:
    log("BUYMA 출고 리스트 다운로드")
    buyma_df = read_buyma_shipment_df(fetch_csv_text(BUYMA_SHIPMENT_URL))

    log("현재고 변환 시트 다운로드")
    stock_df = pd.read_csv(StringIO(fetch_csv_text(STOCK_TRANSFORM_CSV_URL)), dtype=str).fillna("")

    log("환율 시트 다운로드")
    exchange_rate = fetch_exchange_rate()

    platform_reference_df = fetch_platform_reference_df()
    delivery_lookup = make_delivery_lookup(platform_reference_df)

    prepared_rows, stock_lookup = build_prepared_rows(
        buyma_df,
        stock_df,
        exchange_rate,
        delivery_lookup,
    )
    output_df = apply_auto_allocation(prepared_rows, stock_lookup)

    if not include_existing_searching:
        return output_df

    searching_df = build_searching_df(platform_reference_df)
    searching_out_df = build_searching_out_df(searching_df)
    if not searching_out_df.empty:
        log(f"기존 구하는 중 내역 {len(searching_out_df)}행을 이어붙입니다.")
    return pd.concat([output_df, searching_out_df], ignore_index=True)


def upload_to_google_sheet(df: pd.DataFrame, sheet_name: str = TARGET_SHEET_NAME):
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": sheet_name,
        "values": [df.columns.tolist()] + df.fillna("").astype(str).values.tolist(),
    }

    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            log(f"업로드 시작: {sheet_name} / {len(df)}행")
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            log(f"업로드 완료 ({sheet_name}): {response.text}")
            return
        except requests.exceptions.RequestException as exc:
            last_error = exc
            log(f"업로드 실패 ({sheet_name}, {attempt}/{RETRIES}): {exc}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def upload_multiple_sheets(upload_jobs: list[tuple[pd.DataFrame, str]]):
    with ThreadPoolExecutor(max_workers=min(3, len(upload_jobs))) as executor:
        futures = [
            executor.submit(upload_to_google_sheet, df, sheet_name)
            for df, sheet_name in upload_jobs
        ]
        for future in as_completed(futures):
            future.result()


def main():
    start_time = time.perf_counter()
    output_df = build_output_df()
    log("제품 위치 시트 다운로드")
    csv_texts = fetch_all_csv_texts({
        "squareone_location": SQUAREONE_LOCATION_CSV_URL,
        "fila_paju_location": FILA_PAJU_LOCATION_CSV_URL,
    })
    squareone_location_df = pd.read_csv(StringIO(csv_texts["squareone_location"]), dtype=str).fillna("")
    fila_paju_location_raw = pd.read_csv(StringIO(csv_texts["fila_paju_location"]), dtype=str).fillna("")
    fila_paju_location_df = load_selected_excel_cols(fila_paju_location_raw, ["E", "F"])
    upload_multiple_sheets([
        (output_df, TARGET_SHEET_NAME),
        (squareone_location_df, SQUAREONE_LOCATION_SHEET_NAME),
        (fila_paju_location_df, FILA_PAJU_LOCATION_SHEET_NAME),
    ])
    log(f"완료 ({time.perf_counter() - start_time:.2f}초)")


if __name__ == "__main__":
    main()
