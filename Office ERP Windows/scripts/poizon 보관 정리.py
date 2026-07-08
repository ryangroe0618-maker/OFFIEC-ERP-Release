# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from io import StringIO
import re
import time

import pandas as pd
import requests


SOURCE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS64g4a29JXVwLOYSLZXRgXJk2NYUl7U0rFnYm3HWlh50KzVC21jttJypukcOVJqhTbDEx_2CSuCRDr/pub?output=csv"
ONLINE_SIZE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=2089368433&single=true&output=csv"
STOCK_PREP_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"
EXCHANGE_RATE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=295228098&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbz22A8bGUEUumvDbAa5hy556yLq_X1DF87-o08Z4hZEOj3GqKyx8vEQCfaaGH7X2dhmUA/exec"
SPREADSHEET_ID = "1-ArRdQ8VmpSsZ02gFF0zAQsUmFLdov5GCyg_D_NOmZY"
SHEET_NAME = "POIZON - OUT"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 300)
RETRIES = 3
RETRY_SLEEP_SEC = 2

OUTPUT_COLUMNS = [
    "날짜",
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
]


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[POIZON 보관] {message}", flush=True)


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


def download_all_csvs() -> dict[str, str]:
    url_map = {
        "source": SOURCE_URL,
        "online_size": ONLINE_SIZE_URL,
        "stock_prep": STOCK_PREP_URL,
        "exchange": EXCHANGE_RATE_URL,
    }
    results = {}

    with ThreadPoolExecutor(max_workers=len(url_map)) as executor:
        future_map = {
            executor.submit(fetch_csv_text, make_session(), url): key
            for key, url in url_map.items()
        }
        for future in as_completed(future_map):
            results[future_map[future]] = future.result()

    return results


def read_csv_text(csv_text: str) -> pd.DataFrame:
    return pd.read_csv(StringIO(csv_text), dtype=str).fillna("")


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("\\", "/")
    return re.sub(r"\s+", " ", text)


def to_number(value, default=0.0) -> float:
    text = clean_text(value)
    text = text.replace("¥", "").replace(",", "").replace("%", "")
    if text == "":
        return default
    try:
        return float(text)
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


def extract_size(value) -> str:
    text = clean_text(value)
    if text == "":
        return ""
    parts = text.split(" ")
    return clean_text(parts[-1])


def normalize_size(value) -> str:
    text = clean_text(value)
    text = re.sub(r"\.0+$", "", text)
    return re.sub(r"^0+(?=\d)", "", text)


def size_step1_series(series: pd.Series) -> pd.Series:
    s = (
        series.fillna("").astype(str).str.strip()
        .str.replace("（", "(", regex=False)
        .str.replace("）", ")", regex=False)
        .str.replace("\\", "/", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )
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


def get_last4(value) -> str:
    text = clean_text(value)
    if text == "":
        return ""
    suffix = text[-4:]
    if suffix.startswith("0"):
        return f"'{suffix}"
    return suffix


def load_exchange_rate(df_raw: pd.DataFrame) -> float:
    for _, row in df_raw.iterrows():
        for value in row.tolist():
            number = to_number(value, default=None)
            if number not in (None, 0):
                return number
    return 0.0


def build_supply_price_map(stock_df: pd.DataFrame) -> dict[str, float]:
    if "품번" not in stock_df.columns or "공급가" not in stock_df.columns:
        return {}

    temp = stock_df[["품번", "공급가"]].copy()
    temp["품번"] = temp["품번"].apply(clean_text)
    temp["공급가"] = temp["공급가"].apply(to_number)
    temp = temp[temp["품번"] != ""].drop_duplicates(subset=["품번"], keep="first")
    return dict(zip(temp["품번"], temp["공급가"]))


def build_sale_price_map(stock_df: pd.DataFrame) -> dict[str, float]:
    if "품번" not in stock_df.columns or "할인가" not in stock_df.columns:
        return {}

    temp = stock_df[["품번", "할인가"]].copy()
    temp["품번"] = temp["품번"].apply(clean_text)
    temp["할인가"] = temp["할인가"].apply(to_number)
    temp = temp[temp["품번"] != ""].drop_duplicates(subset=["품번"], keep="first")
    return dict(zip(temp["품번"], temp["할인가"]))


def load_online_size_map(df_raw: pd.DataFrame) -> pd.DataFrame:
    if df_raw.shape[1] < 3:
        raise DataValidationError("온라인 사이즈표 열 수가 부족합니다.")

    df = df_raw.iloc[:, :3].copy()
    df.columns = ["브랜드", "사이즈", "사이즈변환"]
    df["브랜드"] = normalize_brand_series(df["브랜드"])
    df["사이즈"] = df["사이즈"].apply(clean_text)
    df["사이즈변환"] = df["사이즈변환"].apply(clean_text)
    return df


def build_online_size_lookup(df: pd.DataFrame) -> dict[tuple[str, str], str]:
    temp = df[["브랜드", "사이즈", "사이즈변환"]].copy()
    temp["key"] = list(zip(temp["브랜드"], temp["사이즈"]))
    temp = temp.drop_duplicates(subset=["key"], keep="first")
    return dict(zip(temp["key"], temp["사이즈변환"]))


def build_output_df(source_df: pd.DataFrame, size_map_df: pd.DataFrame, stock_df: pd.DataFrame, exchange_df: pd.DataFrame) -> pd.DataFrame:
    required_columns = [
        "订单号",
        "POIZON商品名称",
        "POIZON商品货号",
        "SKU ID",
        "POIZON商品规格",
        "件数",
        "订单状态",
        "价格",
        "技术服务费",
        "交易处理手续费",
        "操作服务费",
        "其他费用",
        "违约金",
        "预计总收入",
    ]
    missing_columns = [col for col in required_columns if col not in source_df.columns]
    if missing_columns:
        raise DataValidationError(f"보관 시트 필수 열이 없습니다: {missing_columns}")

    output_df = source_df[required_columns].copy()
    output_df["订单状态"] = output_df["订单状态"].apply(clean_text)
    output_df = output_df[output_df["订单状态"].eq("交易成功")].copy()

    if output_df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    output_df["주문번호"] = output_df["订单号"].apply(clean_text)
    output_df["뒤 4자리"] = output_df["주문번호"].apply(get_last4)
    output_df["브랜드"] = normalize_brand_series(output_df["POIZON商品名称"])
    output_df["품번"] = transform_item_no_series(output_df["POIZON商品货号"], output_df["브랜드"])
    output_df["사이즈원본"] = output_df["POIZON商品规格"].apply(extract_size)
    output_df["사이즈1"] = size_step1_series(output_df["사이즈원본"])
    output_df["사이즈2"] = size_step2_series(output_df["사이즈1"])
    size_lookup = build_online_size_lookup(load_online_size_map(size_map_df))
    output_df["사이즈"] = list(zip(output_df["브랜드"], output_df["사이즈2"]))
    output_df["사이즈"] = output_df["사이즈"].map(size_lookup).fillna("")
    output_df["사이즈"] = output_df["사이즈"].where(
        output_df["사이즈"].astype(str).str.strip() != "",
        output_df["사이즈2"],
    )
    output_df["코드"] = output_df["품번"].apply(clean_text) + output_df["사이즈"].apply(normalize_size)
    output_df["수량"] = output_df["件数"].apply(lambda value: int(round(to_number(value, 0))))
    output_df["총 판매가"] = output_df["价格"].apply(to_number)

    fee_columns = ["技术服务费", "交易处理手续费", "操作服务费", "其他费用", "违约金"]
    output_df["총 수수료"] = output_df[fee_columns].apply(
        lambda row: abs(sum(to_number(value, 0) for value in row.tolist())),
        axis=1,
    )
    output_df["총 수입"] = output_df["预计总收入"].apply(to_number)

    exchange_rate = load_exchange_rate(exchange_df)
    sale_price_map = build_sale_price_map(stock_df)
    supply_price_map = build_supply_price_map(stock_df)

    output_df["할인가"] = output_df["품번"].map(sale_price_map).fillna(0)
    output_df["환율"] = exchange_rate
    output_df["총 KRW"] = (output_df["총 수입"] * exchange_rate).round(0)
    output_df["공급가"] = output_df["품번"].map(supply_price_map).fillna(0)
    output_df["공급가"] = pd.to_numeric(output_df["공급가"], errors="coerce").fillna(0).div(1.1)
    output_df["총 공급가"] = (output_df["공급가"] * output_df["수량"]).round(0)
    output_df["마진"] = (output_df["총 KRW"] - output_df["총 공급가"]).round(0)

    grouped_df = (
        output_df.groupby(["브랜드", "코드", "품번", "사이즈"], dropna=False, sort=False)
        .agg(
            수량=("수량", "sum"),
            총_판매가=("총 판매가", "sum"),
            총_수수료=("총 수수료", "sum"),
            총_수입=("총 수입", "sum"),
            총_KRW=("총 KRW", "sum"),
            총_공급가=("총 공급가", "sum"),
            마진=("마진", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "총_판매가": "총 판매가",
                "총_수수료": "총 수수료",
                "총_수입": "총 수입",
                "총_KRW": "총 KRW",
                "총_공급가": "총 공급가",
            }
        )
    )

    grouped_df["날짜"] = datetime.now().strftime("%Y-%m-%d")
    grouped_df["플랫폼"] = "POIZON 보관"
    grouped_df["주문번호"] = ""
    grouped_df["뒤 4자리"] = ""
    grouped_df["사이즈"] = grouped_df["사이즈"].apply(normalize_size)
    grouped_df["코드"] = grouped_df["품번"].apply(clean_text) + grouped_df["사이즈"].apply(normalize_size)
    grouped_df["매장명"] = "보관"
    grouped_df["할인가"] = grouped_df["품번"].map(sale_price_map).fillna(0)
    grouped_df["환율"] = exchange_rate

    for col in ["수량", "할인가", "총 판매가", "총 수수료", "총 수입", "총 KRW", "총 공급가", "마진"]:
        grouped_df[col] = pd.to_numeric(grouped_df[col], errors="coerce").fillna(0).round(0).astype(int)
    grouped_df["환율"] = pd.to_numeric(grouped_df["환율"], errors="coerce").fillna(0)

    return grouped_df[OUTPUT_COLUMNS].copy()


def upload_to_google_sheet(df: pd.DataFrame) -> bool:
    if df.empty:
        raise DataValidationError("업로드할 데이터가 없습니다.")

    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": SHEET_NAME,
        "values": [df.columns.tolist()] + df.fillna("").astype(str).values.tolist(),
    }

    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            log(f"업로드 완료 -> {SHEET_NAME} / {len(df)}행")
            return True
        except Exception as e:
            last_error = e
            response_text = ""
            if hasattr(e, "response") and getattr(e, "response") is not None:
                response_text = (e.response.text or "").strip()
            log(f"업로드 실패 ({attempt}/{RETRIES}) -> {e}")
            if response_text:
                log(f"응답 본문 -> {response_text}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)

    log("업로드를 중단합니다. 백업 CSV를 사용해주세요.")
    return False


def main():
    start_time = time.perf_counter()
    log("CSV 다운로드 시작")
    csv_texts = download_all_csvs()

    source_df = read_csv_text(csv_texts["source"])
    size_map_df = read_csv_text(csv_texts["online_size"])
    stock_df = read_csv_text(csv_texts["stock_prep"])
    exchange_df = read_csv_text(csv_texts["exchange"])

    output_df = build_output_df(source_df, size_map_df, stock_df, exchange_df)
    if output_df.empty:
        raise DataValidationError("交易成功 상태 데이터가 없습니다.")

    uploaded = upload_to_google_sheet(output_df)
    if not uploaded:
        log("웹앱 URL 또는 배포 권한을 확인해주세요.")

    elapsed = time.perf_counter() - start_time
    log(f"완료 ({elapsed:.2f}초)")


if __name__ == "__main__":
    main()
