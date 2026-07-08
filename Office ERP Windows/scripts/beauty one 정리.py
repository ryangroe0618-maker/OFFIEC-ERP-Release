# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
import time

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
SOURCE_SCRIPT_PATH = BASE_DIR / "beauty one.py"
TARGET_SHEET_NAME = "취합"
REFERENCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSm0aI5pid38q7OLU_E1DNM9j8iULBFUPlbIPetHbvbDtgEZmmjPVo6UyQVgcSb0KDRxNCN0fdvWCpI/pub?gid=1267724548&single=true&output=csv"
PRODUCT_INFO_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=0&single=true&output=csv"
NORTHFACE_DISCOUNT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=1991382423&single=true&output=csv"

STORE_ORDER = [
    "사무실 - 사무실",
    "사무실 - 구월",
    "사무실 - 부천",
    "사무실 - 스퀘어원",
    "사무실 - S마켓",
    "스퀘어원",
    "부천",
    "구월",
    "푸마 여주",
    "휠라 파주",
]


def log(message: str):
    print(message, flush=True)


def load_beauty_one_module():
    spec = spec_from_file_location("beauty_one_module", SOURCE_SCRIPT_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"스크립트를 불러올 수 없습니다: {SOURCE_SCRIPT_PATH}")
    module = module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def ordered_store_columns(store_names: list[str]) -> list[str]:
    preferred = [store for store in STORE_ORDER if store in store_names]
    remaining = sorted(store for store in store_names if store not in STORE_ORDER)
    return preferred + remaining


def clean_text_series(series: pd.Series, beauty_one) -> pd.Series:
    return series.fillna("").astype(str).map(beauty_one.clean_text)


def load_reference_sheet(beauty_one) -> pd.DataFrame:
    session = beauty_one.make_session()
    ref_raw = beauty_one.read_google_sheet_csv_from_text(
        beauty_one.fetch_csv_text(session, REFERENCE_CSV_URL)
    )

    if ref_raw.shape[1] < 4:
        raise ValueError(f"참고 시트 열 개수가 부족합니다. 실제 열 수: {ref_raw.shape[1]}")

    ref_df = ref_raw.iloc[:, :4].copy()
    ref_df.columns = ["매장명", "품번", "사이즈", "수량"]

    for column in ["매장명", "품번", "사이즈"]:
        ref_df[column] = clean_text_series(ref_df[column], beauty_one)

    ref_df["수량"] = pd.to_numeric(ref_df["수량"], errors="coerce").fillna(0).astype(int)
    ref_df["임시코드"] = ref_df["품번"].astype(str) + ref_df["사이즈"].astype(str)
    ref_df = ref_df[
        (ref_df["매장명"] != "")
        & (ref_df["품번"] != "")
        & (ref_df["사이즈"] != "")
    ].reset_index(drop=True)
    return ref_df


def discount_to_ratio(value) -> float:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return 0.0
    number = float(number)
    if number > 1:
        return number / 100.0
    return number


def supply_rate_from_discount(value) -> float:
    ratio = discount_to_ratio(value)
    if ratio >= 0.30:
        return 0.55
    if ratio >= 0.20:
        return 0.62
    return 0.72


def load_price_sheet(beauty_one) -> pd.DataFrame:
    session = beauty_one.make_session()

    def fetch_sheet(url: str) -> pd.DataFrame:
        return beauty_one.read_google_sheet_csv_from_text(
            beauty_one.fetch_csv_text(session, url)
        )

    with ThreadPoolExecutor(max_workers=3) as executor:
        product_future = executor.submit(fetch_sheet, PRODUCT_INFO_CSV_URL)
        discount_future = executor.submit(fetch_sheet, NORTHFACE_DISCOUNT_CSV_URL)
        stock_future = executor.submit(fetch_sheet, beauty_one.STOCK_PREP_URL)
        product_raw = product_future.result()
        discount_raw = discount_future.result()
        stock_raw = stock_future.result()

    if product_raw.shape[1] < 7:
        raise ValueError(f"제품 정보 시트 열 개수가 부족합니다. 실제 열 수: {product_raw.shape[1]}")
    if discount_raw.shape[1] < 2:
        raise ValueError(f"노스페이스 할인율 시트 열 개수가 부족합니다. 실제 열 수: {discount_raw.shape[1]}")
    stock_required_columns = ["브랜드", "품번", "공급가"]
    stock_missing = [column for column in stock_required_columns if column not in stock_raw.columns]
    if stock_missing:
        raise ValueError(f"현재고 변환 시트 필수 열이 없습니다: {', '.join(stock_missing)}")

    product_df = product_raw.iloc[:, [2, 6]].copy()
    product_df.columns = ["품번", "최초가"]
    product_df["품번"] = clean_text_series(product_df["품번"], beauty_one)
    product_df["최초가"] = product_df["최초가"].apply(beauty_one.to_number).astype(int)
    product_df = product_df[product_df["품번"] != ""].drop_duplicates(subset=["품번"], keep="first")

    discount_df = discount_raw.iloc[:, [0, 1]].copy()
    discount_df.columns = ["품번", "할인율"]
    discount_df["품번"] = clean_text_series(discount_df["품번"], beauty_one)
    discount_df["할인율"] = pd.to_numeric(discount_df["할인율"], errors="coerce") / 100.0
    discount_df = discount_df[discount_df["품번"] != ""].drop_duplicates(subset=["품번"], keep="first")

    stock_df = stock_raw[stock_required_columns].copy()
    stock_df["브랜드"] = clean_text_series(stock_df["브랜드"], beauty_one)
    stock_df["품번"] = clean_text_series(stock_df["품번"], beauty_one)
    stock_df["공급가"] = stock_df["공급가"].apply(beauty_one.to_number).astype(int)
    stock_df = stock_df[stock_df["품번"] != ""].drop_duplicates(subset=["품번"], keep="first")

    price_df = product_df.merge(discount_df, on="품번", how="left")
    price_df = price_df.merge(stock_df, on="품번", how="left")
    price_df["할인율_ratio"] = pd.to_numeric(price_df["할인율"], errors="coerce").fillna(0.0)
    price_df["공급가"] = pd.to_numeric(price_df["공급가"], errors="coerce").fillna(0).astype(int)
    return price_df[["브랜드", "품번", "최초가", "할인율", "할인율_ratio", "공급가"]].copy()


def build_out_df(beauty_one) -> pd.DataFrame:
    out_df = beauty_one.build_allocation_result().copy()
    for column in ["코드", "품번", "사이즈", "매장명"]:
        out_df[column] = clean_text_series(out_df[column], beauty_one)
    out_df["수량"] = pd.to_numeric(out_df["수량"], errors="coerce").fillna(0).astype(int)
    out_df["임시코드"] = out_df["품번"].astype(str) + out_df["사이즈"].astype(str)
    return out_df


def build_summary_df(out_df: pd.DataFrame, ref_df: pd.DataFrame, price_df: pd.DataFrame, report_date: str) -> pd.DataFrame:
    out_total_df = (
        out_df.groupby(["품번", "사이즈", "임시코드"], as_index=False)["수량"]
        .sum()
        .rename(columns={"수량": "주문 수량"})
    )

    ref_total_df = (
        ref_df.groupby(["품번", "사이즈", "임시코드"], as_index=False)["수량"]
        .sum()
        .rename(columns={"수량": "출고 수량"})
    )

    store_ship_df = (
        ref_df.pivot_table(
            index=["품번", "사이즈", "임시코드"],
            columns="매장명",
            values="수량",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    summary_df = out_total_df.merge(
        ref_total_df,
        on=["품번", "사이즈", "임시코드"],
        how="outer",
    )
    summary_df = summary_df.merge(
        store_ship_df,
        on=["품번", "사이즈", "임시코드"],
        how="left",
    )
    summary_df = summary_df.merge(
        price_df,
        on=["품번"],
        how="left",
    )

    summary_df["주문 수량"] = pd.to_numeric(summary_df["주문 수량"], errors="coerce").fillna(0).astype(int)
    summary_df["출고 수량"] = pd.to_numeric(summary_df["출고 수량"], errors="coerce").fillna(0).astype(int)
    summary_df["부족 수량"] = (summary_df["출고 수량"] - summary_df["주문 수량"]).astype(int)
    summary_df["최초가"] = pd.to_numeric(summary_df["최초가"], errors="coerce").fillna(0).astype(int)
    summary_df["할인율"] = pd.to_numeric(summary_df["할인율"], errors="coerce")
    summary_df["할인율_ratio"] = pd.to_numeric(summary_df["할인율_ratio"], errors="coerce").fillna(0.0)
    summary_df["할인가"] = (summary_df["최초가"] * (1 - summary_df["할인율_ratio"])).round(0).astype(int)
    summary_df["총 판매가"] = (summary_df["최초가"] * summary_df["출고 수량"]).round(0).astype(int)
    summary_df["총 할인가"] = (summary_df["할인가"] * summary_df["출고 수량"]).round(0).astype(int)
    summary_df["공급가율"] = summary_df["할인율"].apply(supply_rate_from_discount)
    summary_df["총 공급가"] = (summary_df["출고 수량"] * summary_df["최초가"] * summary_df["공급가율"]).round(0).astype(int)
    summary_df["공급가"] = pd.to_numeric(summary_df["공급가"], errors="coerce").fillna(0).astype(int)
    summary_df["공급가"] = (summary_df["공급가"] * summary_df["출고 수량"]).round(0).astype(int)
    summary_df["마진"] = summary_df["총 공급가"] - summary_df["공급가"]

    store_columns = [
        column for column in summary_df.columns
        if column not in [
            "브랜드", "품번", "사이즈", "임시코드",
            "주문 수량", "출고 수량", "부족 수량",
            "최초가", "할인율", "할인율_ratio", "할인가", "총 판매가", "총 할인가", "공급가율", "총 공급가", "공급가", "마진",
        ]
    ]
    store_columns = ordered_store_columns(store_columns)

    for column in store_columns:
        summary_df[column] = pd.to_numeric(summary_df[column], errors="coerce").fillna(0).astype(int)

    summary_df = summary_df.sort_values(
        by=["품번", "사이즈"],
        ascending=[True, True],
    ).reset_index(drop=True)
    summary_df["날짜"] = report_date
    summary_df["플랫폼"] = "LIVE"
    summary_df["브랜드"] = summary_df["브랜드"].fillna("").astype(str)

    fixed_columns = [
        "날짜",
        "플랫폼",
        "브랜드",
        "품번",
        "사이즈",
        "주문 수량",
        "출고 수량",
        "부족 수량",
        "할인율",
        "최초가",
        "할인가",
        "총 판매가",
        "총 할인가",
        "총 공급가",
        "공급가",
        "마진",
    ]
    final_columns = fixed_columns + store_columns
    final_df = summary_df[final_columns].copy()
    final_df["할인율"] = final_df["할인율"].mask(final_df["할인율"].isna(), "")

    numeric_columns = ["주문 수량", "출고 수량", "부족 수량", "최초가", "할인가", "총 판매가", "총 할인가", "총 공급가", "공급가", "마진"] + store_columns
    for column in numeric_columns:
        final_df[column] = final_df[column].replace(0, "")

    return final_df


def main():
    start_time = time.perf_counter()
    beauty_one = load_beauty_one_module()
    report_date = datetime.now().strftime("%Y-%m-%d")

    log("OUT 기준 데이터 재생성 중...")
    out_df = build_out_df(beauty_one)

    log("참고 시트/제품 정보/할인율 시트 다운로드 중...")
    with ThreadPoolExecutor(max_workers=2) as executor:
        ref_future = executor.submit(load_reference_sheet, beauty_one)
        price_future = executor.submit(load_price_sheet, beauty_one)
        ref_df = ref_future.result()
        price_df = price_future.result()

    log(f"취합 생성 중... OUT {len(out_df)}행 / 참고 시트 {len(ref_df)}행")
    summary_df = build_summary_df(out_df, ref_df, price_df, report_date)

    beauty_one.upload_to_google_sheet(
        summary_df,
        beauty_one.WEB_APP_URL,
        beauty_one.SPREADSHEET_ID,
        TARGET_SHEET_NAME,
    )

    elapsed = time.perf_counter() - start_time
    log(f"완료 (총 소요 시간: {elapsed:.2f}초)")


if __name__ == "__main__":
    main()
