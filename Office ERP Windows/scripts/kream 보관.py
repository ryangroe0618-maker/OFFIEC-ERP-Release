# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from io import StringIO
import re
import time

import pandas as pd
import requests


SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTK9BjPuP782_4akOc_w3CayC7LRrb_lp4DtrQYrelclSyJjIgS0gVvORws5uSGu0su3KDgYMGN92ny/pub?gid=0&single=true&output=csv"
ONLINE_SIZE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=2089368433&single=true&output=csv"
STOCK_TRANSFORM_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyD7gGTb-jtE6l4-3vM5po6Yxpy52I8f3mY_SV4w2BA6z-gVw0HZeyL6gMk43y720WO/exec"
SPREADSHEET_ID = "1omENfalGRqmgSOa3QOejuXf99wIhy-gCCacaveBA0IY"
TARGET_SHEET_NAME = "OUT"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 300)
RETRIES = 3
RETRY_SLEEP_SEC = 2

OUTPUT_COLUMNS = ["날짜", "플랫폼", "주문번호", "운송장번호", "브랜드", "코드", "품번", "사이즈", "수량", "매장명", "할인가", "총 판매가", "총 수수료", "총 수입", "환율", "총 KRW", "총 공급가", "마진"]
SOURCE_USECOLS = ["주문/보관번호", "판매유형", "모델품번", "상품명(영문 상품명)", "옵션", "거래금액", "정산예정금액"]
STOCK_REQUIRED_COLUMNS = ["코드", "품번", "사이즈", "변환코드1", "변환코드2", "변환코드3", "할인가", "공급가"]


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[KREAM 보관] {message}", flush=True)


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
        except Exception as exc:
            last_error = exc
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


def clean_text_series(series: pd.Series) -> pd.Series:
    return (
        series.fillna("").astype(str).str.strip()
        .str.replace("（", "(", regex=False)
        .str.replace("）", ")", regex=False)
        .str.replace("\\", "/", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )


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
            .str.replace("鞋", "", regex=False),
        ),
    )

    tnf_fila_values = result.where(~tnf_fila_mask, result)
    tnf_fila_values = tnf_fila_values.str.replace(r"^1100", "", regex=True)
    tnf_fila_values = tnf_fila_values.str.replace("_", "", regex=False).str.replace("-", "", regex=False)
    result = result.mask(tnf_fila_mask, tnf_fila_values)

    asics_mask = brand_series.eq("ASICS")
    asics_values = result.where(~asics_mask, result)
    asics_values = asics_values.str.split("_", n=1).str[0]
    result = result.mask(asics_mask, asics_values)
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
    missing = [col for col in STOCK_REQUIRED_COLUMNS if col not in df_raw.columns]
    if missing:
        raise DataValidationError(f"분배준비 시트 필수 열이 없습니다: {missing}")

    df = df_raw[STOCK_REQUIRED_COLUMNS].copy()
    text_cols = ["코드", "품번", "사이즈", "변환코드1", "변환코드2", "변환코드3"]
    for col in text_cols:
        df[col] = clean_text_series(df[col])
    for col in ["할인가", "공급가"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
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


def load_all_sources() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    csvs = fetch_all_csv_texts(
        {
            "source": SOURCE_CSV_URL,
            "online_size": ONLINE_SIZE_CSV_URL,
            "stock": STOCK_TRANSFORM_CSV_URL,
        }
    )
    source_df = pd.read_csv(
        StringIO(csvs["source"]),
        dtype=str,
        keep_default_na=False,
        usecols=lambda col: col in SOURCE_USECOLS,
    ).fillna("")
    size_df = pd.read_csv(
        StringIO(csvs["online_size"]),
        dtype=str,
        keep_default_na=False,
        usecols=[0, 1, 2],
    ).fillna("")
    stock_df = pd.read_csv(
        StringIO(csvs["stock"]),
        dtype=str,
        keep_default_na=False,
        usecols=lambda col: col in STOCK_REQUIRED_COLUMNS,
    ).fillna("")
    return source_df, size_df, stock_df


def prepare_order_source(source_df: pd.DataFrame, size_map_df: pd.DataFrame, stock_df: pd.DataFrame) -> pd.DataFrame:
    required_columns = SOURCE_USECOLS
    missing = [col for col in required_columns if col not in source_df.columns]
    if missing:
        raise DataValidationError(f"KREAM 보관 원본 시트 필수 열이 없습니다: {missing}")

    src = pd.DataFrame(
        {
            "주문번호": clean_text_series(source_df["주문/보관번호"]),
            "품번원본": clean_text_series(source_df["모델품번"]),
            "브랜드원본": clean_text_series(source_df["상품명(영문 상품명)"]),
            "사이즈원본": clean_text_series(source_df["옵션"]),
            "판매가": source_df["거래금액"].apply(to_number_text),
            "수입": source_df["정산예정금액"].apply(to_number_text),
            "운송장번호": "",
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

    cleaned_stock_codes = clean_text_series(stock_df["코드"])
    stock_code_set = set(cleaned_stock_codes.tolist()) | set(cleaned_stock_codes.apply(compact_code_text).tolist())
    src["매칭코드"] = src["매칭코드"].where(src["매칭코드"].isin(stock_code_set), src["매칭코드"].apply(compact_code_text))
    src = src[["주문번호", "브랜드", "품번", "출력사이즈", "최종코드", "매칭코드", "판매가", "수입", "운송장번호", "수량"]].copy()
    src["수량"] = pd.to_numeric(src["수량"], errors="coerce").fillna(0).astype(int)
    src["판매가"] = pd.to_numeric(src["판매가"], errors="coerce").fillna(0).round(0).astype(int).astype(str)
    src["수입"] = pd.to_numeric(src["수입"], errors="coerce").fillna(0).round(0).astype(int).astype(str)
    return src.reset_index(drop=True)


def build_output_df(order_df: pd.DataFrame, stock_df: pd.DataFrame, today_text: str) -> pd.DataFrame:
    discount_lookup = build_price_lookup(stock_df, "할인가")
    supply_lookup = build_price_lookup(stock_df, "공급가")
    if order_df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    result_df = order_df.copy()
    result_df["날짜"] = today_text
    result_df["플랫폼"] = "KREAM 보관"
    result_df["매장명"] = "보관"
    result_df["환율"] = ""
    result_df["코드"] = clean_text_series(result_df["최종코드"])
    result_df["사이즈"] = clean_text_series(result_df["출력사이즈"])
    result_df["주문번호"] = clean_text_series(result_df["주문번호"])
    result_df["운송장번호"] = clean_text_series(result_df["운송장번호"])
    result_df["브랜드"] = clean_text_series(result_df["브랜드"])
    result_df["품번"] = clean_text_series(result_df["품번"])

    result_df["수량_num"] = pd.to_numeric(result_df["수량"], errors="coerce").fillna(0).astype(int)
    result_df["판매가_num"] = pd.to_numeric(result_df["판매가"], errors="coerce").fillna(0)
    result_df["수입_num"] = pd.to_numeric(result_df["수입"], errors="coerce").fillna(0)
    result_df["품번_압축"] = result_df["품번"].apply(compact_code_text)

    result_df["할인가_num"] = (
        result_df["품번"].map(discount_lookup)
        .fillna(result_df["품번_압축"].map(discount_lookup))
        .fillna(0)
    )
    result_df["공급가_num"] = (
        result_df["품번"].map(supply_lookup)
        .fillna(result_df["품번_압축"].map(supply_lookup))
        .fillna(0)
    )

    result_df["총 판매가_num"] = (result_df["판매가_num"] * result_df["수량_num"]).round(0).astype(int)
    result_df["총 수입_num"] = (result_df["수입_num"] * result_df["수량_num"]).round(0).astype(int)
    result_df["총 KRW_num"] = result_df["총 수입_num"]
    result_df["총 공급가_num"] = (result_df["공급가_num"] * result_df["수량_num"]).round(0).astype(int)
    result_df["총 수수료_num"] = result_df["총 판매가_num"] - result_df["총 수입_num"]
    result_df["마진_num"] = result_df["총 수입_num"] - result_df["총 공급가_num"]

    result_df["수량"] = result_df["수량_num"].mask(result_df["수량_num"].eq(0), "").astype(str)
    result_df["할인가"] = result_df["할인가_num"].round(0).astype(int).mask(result_df["할인가_num"].round(0).astype(int).eq(0), "").astype(str)
    result_df["총 판매가"] = result_df["총 판매가_num"].astype(str)
    result_df["총 수수료"] = result_df["총 수수료_num"].astype(str)
    result_df["총 수입"] = result_df["총 수입_num"].astype(str)
    result_df["총 KRW"] = result_df["총 KRW_num"].astype(str)
    result_df["총 공급가"] = result_df["총 공급가_num"].astype(str)
    result_df["마진"] = result_df["마진_num"].astype(str)

    result_df = result_df[OUTPUT_COLUMNS]
    if result_df.empty:
        result_df = pd.DataFrame(columns=OUTPUT_COLUMNS)
    return result_df.sort_values(by=["주문번호", "코드"], ascending=[True, True]).reset_index(drop=True)


def upload_to_google_sheet(df: pd.DataFrame):
    if df.empty:
        raise DataValidationError("업로드할 데이터가 없습니다.")
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": [df.columns.tolist()] + df.fillna("").astype(str).values.tolist(),
        "clear": True,
    }
    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            log(f"업로드 완료 -> {TARGET_SHEET_NAME} / {len(df)}행")
            return
        except Exception as exc:
            last_error = exc
            log(f"업로드 실패 ({attempt}/{RETRIES}) -> {exc}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def main():
    start_time = time.perf_counter()
    today_text = datetime.now().strftime("%Y-%m-%d")
    log("KREAM 보관 다운로드 시작")
    source_df, size_map_df, stock_raw_df = load_all_sources()
    stock_df = load_stock_prepare(stock_raw_df)
    order_df = prepare_order_source(source_df, size_map_df, stock_df)
    output_df = build_output_df(order_df, stock_df, today_text)
    log(f"출력 데이터 생성 완료 -> {output_df.shape[0]}행")
    upload_to_google_sheet(output_df)
    elapsed = time.perf_counter() - start_time
    log(f"완료 ({elapsed:.2f}초)")


if __name__ == "__main__":
    main()
