# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor, as_completed
from io import StringIO
from pathlib import Path
import importlib.util
import sys
import time

import pandas as pd
import requests


SOURCE_URL = "https://docs.google.com/spreadsheets/d/1P2JmY6iwEf7PF4_yC8TK3xbUfRWEMMZt6ucuGtXZ3-g/gviz/tq?tqx=out:csv&sheet=BUYMA"
STORE_URLS = {
    "스퀘어원": [
        "https://docs.google.com/spreadsheets/d/e/2PACX-1vTPSUW1W1iSIvGGrLkp1WHj6Dy_k4NQHv5xOZR4xviYMsZWUb6ZBQ4PqeI31RM_keSDaXeQsYyNLAav/pub?gid=1199416354&single=true&output=csv",
        "https://docs.google.com/spreadsheets/d/e/2PACX-1vTPSUW1W1iSIvGGrLkp1WHj6Dy_k4NQHv5xOZR4xviYMsZWUb6ZBQ4PqeI31RM_keSDaXeQsYyNLAav/pub?gid=1405262882&single=true&output=csv",
    ],
    "구월": [
        "https://docs.google.com/spreadsheets/d/e/2PACX-1vSHTZiYHkTrDlZ_pi1qxBsikvBAaMxtdEzwSYsWzk6sV1zk04SIYjflfnxMYRsmwevPovu4Mtnlx69M/pub?gid=1199416354&single=true&output=csv",
        "https://docs.google.com/spreadsheets/d/e/2PACX-1vSHTZiYHkTrDlZ_pi1qxBsikvBAaMxtdEzwSYsWzk6sV1zk04SIYjflfnxMYRsmwevPovu4Mtnlx69M/pub?gid=1405262882&single=true&output=csv",
    ],
    "부천": [
        "https://docs.google.com/spreadsheets/d/e/2PACX-1vQdYpr-dKLe-tguI2uOaYL9pjalY0jehboc1zb-B5XKbV8vAPQvtw1S4nu-TaxJULDsoKOTz8gz7A5y/pub?gid=1199416354&single=true&output=csv",
        "https://docs.google.com/spreadsheets/d/e/2PACX-1vQdYpr-dKLe-tguI2uOaYL9pjalY0jehboc1zb-B5XKbV8vAPQvtw1S4nu-TaxJULDsoKOTz8gz7A5y/pub?gid=1405262882&single=true&output=csv",
    ],
    "휠라 파주": [
        "https://docs.google.com/spreadsheets/d/e/2PACX-1vSA47SgFq9QQPg0D3AlBnpJX6q7Yx_Dh66E1ID9MlXTahJjL0FmFVtPgyTEtj4iVj7PvRkCUoCgbjkd/pub?gid=1226300508&single=true&output=csv",
    ],
    "푸마 여주": [
        "https://docs.google.com/spreadsheets/d/e/2PACX-1vSzkIBQ7UfqnUboNBWaQj6esNZzi_NSk0crAVPCljFog-YAnl1vSY6gqqTxH2CYosDoRL4q2PgMUhqL/pub?gid=1405262882&single=true&output=csv",
    ],
}
OFFICE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQhP9cP1QdWll3UPE-P-tUAmxFHCEXgQU_IKIDsftokFeyn5Y67OW2Zho5xYN4pwQKvcclbDS98bQum/pub?gid=202616820&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbytRA2OXU2jNt0llG8w2y0voM2q6HV8Xm1nCstfBQ4b3qMMSjBWLnDobr8aalYErcXdOw/exec"
SPREADSHEET_ID = "1cmY9wL9zP86mdP_BiwIqbjigHFdPemPq0sV8ulL_XMs"
TARGET_SHEET_NAME = "플랫폼"
SUMMARY_SHEET_NAME = "플랫폼 마감"
BRAND_SUMMARY_SHEET_NAME = "브랜드 마감"
APPEND_MODE = True
SUMMARY_COLUMNS = ["날짜", "플랫폼", "주문건수", "주문 수량", "판매가", "수수료", "총 KRW", "총 공급가", "마진", "마진율"]
BRAND_SUMMARY_COLUMNS = ["날짜", "플랫폼", "브랜드", "판매 수량", "총 판매가", "총 KRW", "총 공급가", "마진", "마진율"]

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 300)
RETRIES = 3
RETRY_SLEEP_SEC = 2
SOURCE_SLICE_START = 0
SOURCE_SLICE_END = 20
SOURCE_EXCLUDE_COLUMN_INDICES = {1, 19}
SEARCHING_STATUS = "구하는 중"
STANDARD_COLUMN_RENAMES = {
    "총 판매가": "판매가",
    "총 수수료": "수수료",
    "총 수입": "수입",
    "총 KRW": "KRW",
    "총 공급가": "공급가",
}


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[BUYMA 정리] {message}", flush=True)


def make_session():
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace('"', "").replace("\t", "").replace("\r", "").replace("\n", "")
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("\\", "/")
    text = " ".join(text.split())
    if text.startswith("'"):
        text = text[1:]
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def normalize_key(value) -> str:
    return clean_text(value).upper()


def to_number_series(series: pd.Series) -> pd.Series:
    cleaned = (
        series.fillna("")
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace('"', "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def get_first_column_series(df: pd.DataFrame, column_names: list[str]) -> pd.Series:
    for column_name in column_names:
        if column_name in df.columns:
            return df[column_name].fillna("").astype(str)
    return pd.Series([""] * len(df), index=df.index, dtype=str)


def format_percent(value) -> str:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return "0%"
    return f"{float(number) * 100:.2f}%"


def derive_status(store_name: str) -> str:
    store_name = clean_text(store_name)
    if store_name in {"취소", "취소건"}:
        return "취소건"
    if store_name in {"재고없음", "재고 없음"}:
        return "재고 없음"
    return "출고"


def summarize_platform_name(order_no: str) -> str:
    return "BUYMA"


def is_summary_included(carrier_name: str) -> bool:
    return clean_text(carrier_name) not in {"取消件", "调货中", SEARCHING_STATUS}


def resolve_store_group(store_name: str) -> str:
    store_name = clean_text(store_name)
    if store_name.startswith("사무실"):
        return "사무실"
    if "스퀘어원" in store_name:
        return "스퀘어원"
    if "구월" in store_name:
        return "구월"
    if "부천" in store_name:
        return "부천"
    if "휠라 파주" in store_name:
        return "휠라 파주"
    if "푸마 여주" in store_name:
        return "푸마 여주"
    return ""


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


def parse_csv_text(text: str) -> pd.DataFrame:
    return pd.read_csv(
        StringIO(text),
        dtype=str,
        header=0,
        keep_default_na=False,
        skip_blank_lines=False,
    ).fillna("")


def fetch_csv_job(name: str, url: str) -> tuple[str, str]:
    return name, fetch_csv_text(make_session(), url)


def load_buyma_upload_module():
    module_path = Path(__file__).with_name("buyma 리스트 업로드.py")
    spec = importlib.util.spec_from_file_location("buyma_list_upload_for_cleanup", module_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def download_all_csv_texts() -> dict[str, str]:
    url_map = {"source": SOURCE_URL, "office": OFFICE_URL}
    for store, urls in STORE_URLS.items():
        for idx, url in enumerate(urls, start=1):
            url_map[f"store::{store}::{idx}"] = url

    results = {}
    with ThreadPoolExecutor(max_workers=min(len(url_map), 8)) as executor:
        future_map = {
            executor.submit(fetch_csv_job, name, url): name
            for name, url in url_map.items()
        }
        for future in as_completed(future_map):
            name = future_map[future]
            try:
                name, text = future.result()
            except Exception as exc:
                if name == "source":
                    log(f"BUYMA 원본 시트를 직접 읽지 못해 리스트 업로드 결과를 재생성합니다: {exc}")
                    results[name] = ""
                    continue
                raise
            results[name] = text

    return results

def load_source_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    if df_raw.shape[1] < SOURCE_SLICE_END:
        raise DataValidationError(
            f"원본 시트 열 수가 부족합니다. 최소 S열까지 필요, 실제 {df_raw.shape[1]}열"
        )

    df = df_raw.iloc[:, SOURCE_SLICE_START:SOURCE_SLICE_END].copy()
    df = df.drop(df.columns[list(SOURCE_EXCLUDE_COLUMN_INDICES)], axis=1)
    df.columns = [clean_text(col) for col in df.columns]
    df = df.rename(columns=STANDARD_COLUMN_RENAMES)
    df = df.loc[~(df.astype(str).apply(lambda row: "".join(row).strip() == "", axis=1))].reset_index(drop=True)

    if df.empty:
        raise DataValidationError("원본 시트에 업로드할 데이터가 없습니다.")

    if "주문번호" not in df.columns or "매장명" not in df.columns:
        raise DataValidationError("원본 A:R 범위에 '주문번호' 또는 '매장명' 머리글이 없습니다.")

    return df


def load_source_df_from_upload_script() -> pd.DataFrame:
    buyma_upload = load_buyma_upload_module()
    output_df = buyma_upload.build_output_df(include_existing_searching=False)
    if output_df.empty:
        raise DataValidationError("BUYMA 리스트 업로드 결과에 정리할 데이터가 없습니다.")
    return load_source_df(output_df)


def build_store_lookup(dfs: list[pd.DataFrame]) -> dict[str, str]:
    lookup = {}

    for df in dfs:
        if df.shape[1] < 11:
            continue

        order_series = df.iloc[:, 9].apply(clean_text)
        tracking_series = df.iloc[:, 10].apply(clean_text)

        for order_no, tracking in zip(order_series, tracking_series):
            lookup_key = normalize_key(order_no)
            if not lookup_key or lookup_key in lookup:
                continue
            if not tracking:
                continue
            lookup[lookup_key] = tracking

    return lookup


def build_office_lookup(df: pd.DataFrame) -> dict[str, str]:
    lookup = {}

    if df.shape[1] < 12:
        return lookup

    order_series = df.iloc[:, 11].apply(clean_text)

    for order_no in order_series:
        lookup_key = normalize_key(order_no)
        if not lookup_key or lookup_key in lookup:
            continue
        lookup[lookup_key] = "自送到仓"

    return lookup


def build_output_df(source_df: pd.DataFrame, store_lookups: dict[str, dict[str, str]], office_lookup: dict[str, str]) -> pd.DataFrame:
    output_df = source_df.copy()

    carriers = []
    trackings = []
    for row in output_df.itertuples(index=False):
        order_no = normalize_key(getattr(row, "주문번호", ""))
        store_name = clean_text(getattr(row, "매장명", ""))
        store_group = resolve_store_group(store_name)
        carrier = ""
        tracking = ""

        if store_name in {"취소", "취소건"}:
            carrier = "取消件"
        elif store_name in {"재고없음", "재고 없음"}:
            carrier = SEARCHING_STATUS
        elif store_group == "사무실":
            carrier = office_lookup.get(order_no, "")
            if not carrier:
                carrier = SEARCHING_STATUS
        elif store_group:
            tracking = store_lookups.get(store_group, {}).get(order_no, "")
            if tracking:
                carrier = "CJ"
            else:
                carrier = SEARCHING_STATUS

        carriers.append(carrier)
        trackings.append(tracking)

    output_df["택배사"] = carriers
    output_df["운송장"] = trackings
    return output_df


def build_platform_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = ["날짜", "주문번호", "수량", "환율", "마진", "택배사"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise DataValidationError(f"플랫폼 마감 집계용 필수 열이 없습니다: {missing_columns}")
    metric_aliases = {
        "판매가": ["판매가", "총 판매가"],
        "수수료": ["수수료", "총 수수료"],
        "수입": ["수입", "총 수입"],
        "KRW": ["KRW", "총 KRW"],
        "공급가": ["공급가", "총 공급가"],
    }
    metric_missing = [
        name
        for name, aliases in metric_aliases.items()
        if not any(alias in df.columns for alias in aliases)
    ]
    if metric_missing:
        raise DataValidationError(f"플랫폼 마감 집계용 금액 열이 없습니다: {metric_missing}")

    summary_base_df = df.copy()
    summary_base_df["날짜"] = summary_base_df["날짜"].apply(clean_text)
    summary_base_df["주문번호"] = summary_base_df["주문번호"].apply(clean_text)
    summary_base_df["택배사"] = summary_base_df["택배사"].apply(clean_text)
    summary_base_df["집계플랫폼"] = summary_base_df["주문번호"].apply(summarize_platform_name)
    summary_base_df = summary_base_df[summary_base_df["택배사"].apply(is_summary_included)].copy()

    if summary_base_df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    summary_base_df["주문번호정규화"] = summary_base_df["주문번호"].apply(normalize_key)
    summary_base_df["수량_num"] = to_number_series(summary_base_df["수량"])
    summary_base_df["판매가_num"] = to_number_series(get_first_column_series(summary_base_df, metric_aliases["판매가"]))
    summary_base_df["수수료_num"] = to_number_series(get_first_column_series(summary_base_df, metric_aliases["수수료"]))
    summary_base_df["수입_num"] = to_number_series(get_first_column_series(summary_base_df, metric_aliases["수입"]))
    summary_base_df["환율_num"] = to_number_series(summary_base_df["환율"])
    summary_base_df["KRW_num"] = to_number_series(get_first_column_series(summary_base_df, metric_aliases["KRW"]))
    summary_base_df["공급가_num"] = to_number_series(get_first_column_series(summary_base_df, metric_aliases["공급가"]))
    summary_base_df["마진_num"] = to_number_series(summary_base_df["마진"])
    summary_base_df["판매가"] = summary_base_df["판매가_num"] * summary_base_df["환율_num"]
    summary_base_df["수수료"] = summary_base_df["수수료_num"] * summary_base_df["환율_num"]

    summary_df = (
        summary_base_df.groupby(["날짜", "집계플랫폼"], dropna=False, sort=False)
        .agg(
            주문건수=("주문번호정규화", "nunique"),
            주문_수량=("수량_num", "sum"),
            판매가=("판매가", "sum"),
            수수료=("수수료", "sum"),
            총_KRW=("KRW_num", "sum"),
            총_공급가=("공급가_num", "sum"),
            마진=("마진_num", "sum"),
        )
        .reset_index()
    )

    summary_df = summary_df.rename(
        columns={
            "집계플랫폼": "플랫폼",
            "주문_수량": "주문 수량",
            "총_KRW": "총 KRW",
            "총_공급가": "총 공급가",
        }
    )
    summary_df["마진율"] = summary_df["마진"].div(summary_df["총 KRW"].replace(0, pd.NA)).fillna(0)

    for col in ["주문건수", "주문 수량", "판매가", "수수료", "총 KRW", "총 공급가", "마진"]:
        summary_df[col] = summary_df[col].round(0).astype(int)
    summary_df["마진율"] = summary_df["마진율"].apply(format_percent)

    return summary_df[SUMMARY_COLUMNS]


def build_brand_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = ["날짜", "주문번호", "브랜드", "수량", "환율", "마진", "택배사"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise DataValidationError(f"브랜드 마감 집계용 필수 열이 없습니다: {missing_columns}")
    metric_aliases = {
        "판매가": ["판매가", "총 판매가"],
        "KRW": ["KRW", "총 KRW"],
        "공급가": ["공급가", "총 공급가"],
    }
    metric_missing = [
        name
        for name, aliases in metric_aliases.items()
        if not any(alias in df.columns for alias in aliases)
    ]
    if metric_missing:
        raise DataValidationError(f"브랜드 마감 집계용 금액 열이 없습니다: {metric_missing}")

    summary_base_df = df.copy()
    summary_base_df["날짜"] = summary_base_df["날짜"].apply(clean_text)
    summary_base_df["주문번호"] = summary_base_df["주문번호"].apply(clean_text)
    summary_base_df["택배사"] = summary_base_df["택배사"].apply(clean_text)
    summary_base_df["플랫폼"] = summary_base_df["주문번호"].apply(summarize_platform_name)
    summary_base_df = summary_base_df[summary_base_df["택배사"].apply(is_summary_included)].copy()

    if summary_base_df.empty:
        return pd.DataFrame(columns=BRAND_SUMMARY_COLUMNS)

    summary_base_df["브랜드"] = summary_base_df["브랜드"].apply(clean_text)
    summary_base_df["수량_num"] = to_number_series(summary_base_df["수량"])
    summary_base_df["판매가_num"] = to_number_series(get_first_column_series(summary_base_df, metric_aliases["판매가"]))
    summary_base_df["환율_num"] = to_number_series(summary_base_df["환율"])
    summary_base_df["KRW_num"] = to_number_series(get_first_column_series(summary_base_df, metric_aliases["KRW"]))
    summary_base_df["공급가_num"] = to_number_series(get_first_column_series(summary_base_df, metric_aliases["공급가"]))
    summary_base_df["마진_num"] = to_number_series(summary_base_df["마진"])
    summary_base_df["판매가_원화"] = summary_base_df["판매가_num"] * summary_base_df["환율_num"]

    summary_df = (
        summary_base_df.groupby(["날짜", "플랫폼", "브랜드"], dropna=False, sort=False)
        .agg(
            {
                "수량_num": "sum",
                "판매가_원화": "sum",
                "KRW_num": "sum",
                "공급가_num": "sum",
                "마진_num": "sum",
            }
        )
        .reset_index()
        .rename(
            columns={
                "수량_num": "판매 수량",
                "판매가_원화": "총 판매가",
                "KRW_num": "총 KRW",
                "공급가_num": "총 공급가",
                "마진_num": "마진",
            }
        )
    )

    for col in ["판매 수량", "총 판매가", "총 KRW", "총 공급가", "마진"]:
        summary_df[col] = summary_df[col].round(0).astype(int)
    summary_df["마진율"] = summary_df["마진"].div(summary_df["총 KRW"].replace(0, pd.NA)).fillna(0)
    summary_df["마진율"] = summary_df["마진율"].apply(format_percent)

    return summary_df[BRAND_SUMMARY_COLUMNS]


def upload_to_google_sheet(df: pd.DataFrame, sheet_name: str):
    if df.empty:
        raise DataValidationError("업로드할 데이터가 없습니다.")

    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": sheet_name,
        "values": values,
        "append": APPEND_MODE,
        "clear": not APPEND_MODE,
    }

    mode_label = "누적" if APPEND_MODE else "덮어쓰기"
    log(f"업로드 시작 -> {sheet_name} ({len(df)}행, {df.shape[1]}열, {mode_label})")
    last_error = None

    for attempt in range(1, RETRIES + 1):
        try:
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            log(f"업로드 완료 -> {sheet_name}")
            return
        except Exception as e:
            last_error = e
            log(f"업로드 실패 ({attempt}/{RETRIES}) -> {e}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)

    raise last_error


def main():
    start_time = time.perf_counter()
    log("CSV 병렬 다운로드 시작")
    csv_texts = download_all_csv_texts()

    log("원본 시트 로드 시작")
    if csv_texts.get("source", ""):
        source_df = load_source_df(parse_csv_text(csv_texts["source"]))
    else:
        source_df = load_source_df_from_upload_script()
    log(f"원본 시트 읽기 완료 -> {source_df.shape[0]}행, {source_df.shape[1]}열")

    store_lookups = {}
    for store, urls in STORE_URLS.items():
        dfs = [
            parse_csv_text(csv_texts[f"store::{store}::{idx}"])
            for idx, _ in enumerate(urls, start=1)
        ]
        store_lookups[store] = build_store_lookup(dfs)

    office_lookup = build_office_lookup(parse_csv_text(csv_texts["office"]))
    output_df = build_output_df(source_df, store_lookups, office_lookup)
    summary_df = build_platform_summary_df(output_df)
    brand_summary_df = build_brand_summary_df(output_df)

    upload_to_google_sheet(output_df, TARGET_SHEET_NAME)

    if summary_df.empty:
        log("플랫폼 마감 업로드 대상 없음 -> 출고 데이터가 없습니다.")
    else:
        log(f"플랫폼 마감 집계 완료 -> {summary_df.shape[0]}행")
        upload_to_google_sheet(summary_df, SUMMARY_SHEET_NAME)

    if brand_summary_df.empty:
        log(f"브랜드 마감 업로드 대상 없음 -> 取消件, {SEARCHING_STATUS} 제외 후 데이터가 없습니다.")
    else:
        log(f"브랜드 마감 집계 완료 -> {brand_summary_df.shape[0]}행")
        upload_to_google_sheet(brand_summary_df, BRAND_SUMMARY_SHEET_NAME)

    elapsed = time.perf_counter() - start_time
    log(f"완료 (총 소요 시간: {elapsed:.2f}초)")


if __name__ == "__main__":
    main()
