# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor, as_completed
from io import StringIO
import re
import time

import pandas as pd
import requests


RETURN_INPUT_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQFuLbNobNHN9skvSn0CJemLbJmplfqcUiLYr_KTlChLSijczOy_INmeYmaMtimT3LyY8YT8FfU2ws8/pub?gid=0&single=true&output=csv"
PLATFORM_CLOSE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRWhzq2rt_O-YnOwzWnE8L_d8--NBu-EMmDkxnnVy-oATAfCjDIX976b973bY2MiO8gYPTE5WCk-tVv/pub?output=csv"
SIZE_CHART_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=1121158649&single=true&output=csv"
STOCK_PREP_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=1491233836&single=true&output=csv"

WEB_APP_URL = "https://script.google.com/macros/s/AKfycby0Hi2WRMlCYC6zP-nBhgjj7YURqsIPjw7DfNfSmFF0x48bOzsNzq8IzZ7tUXea-mx9/exec"
SPREADSHEET_ID = "1DCZPfapugO6HNFYO0jzpwQQjE8hthq-YI5rFwp-KdGM"
SHEET_NAME = "OUT"

DOWNLOAD_RETRIES = 3
DOWNLOAD_TIMEOUT = (10, 60)
DOWNLOAD_SLEEP_SEC = 2
UPLOAD_RETRIES = 3
UPLOAD_TIMEOUT = (10, 300)
UPLOAD_SLEEP_SEC = 3

RETURN_USECOLS = ["E", "F", "G", "I", "J", "M"]
RETURN_COLUMNS = ["운송장", "품번원본", "브랜드원본", "사이즈원본", "주문번호", "수량"]
PLATFORM_CLOSE_COLUMNS = [
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
]
PLATFORM_SOURCE_COLUMNS = ["날짜", "플랫폼", "주문번호", "운송장"] + PLATFORM_CLOSE_COLUMNS
OUTPUT_COLUMNS = ["날짜", "원거래 날짜", "플랫폼", "주문번호", "운송장"] + PLATFORM_CLOSE_COLUMNS


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[poizon 반품] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(url: str) -> str:
    last_error = None
    session = make_session()
    for attempt in range(1, DOWNLOAD_RETRIES + 1):
        try:
            response = session.get(url, timeout=DOWNLOAD_TIMEOUT)
            response.raise_for_status()
            response.encoding = "utf-8"
            text = response.text
            if not text.strip():
                raise DataValidationError("CSV 응답이 비어 있습니다.")
            return text
        except Exception as exc:
            last_error = exc
            log(f"다운로드 실패 ({attempt}/{DOWNLOAD_RETRIES}): {exc}")
            if attempt < DOWNLOAD_RETRIES:
                time.sleep(DOWNLOAD_SLEEP_SEC)
    raise last_error


def download_all_csvs() -> dict[str, str]:
    urls = {
        "return": RETURN_INPUT_URL,
        "platform_close": PLATFORM_CLOSE_URL,
        "size_chart": SIZE_CHART_URL,
        "stock_prep": STOCK_PREP_URL,
    }
    results = {}

    with ThreadPoolExecutor(max_workers=len(urls)) as executor:
        future_map = {
            executor.submit(fetch_csv_text, url): key
            for key, url in urls.items()
        }
        for future in as_completed(future_map):
            results[future_map[future]] = future.result()

    return results


def read_google_sheet_csv_from_text(csv_text: str) -> pd.DataFrame:
    return pd.read_csv(StringIO(csv_text), dtype=str).fillna("")


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("\\", "/")
    return re.sub(r"\s+", " ", text)


def to_number(value, default=0):
    text = clean_text(value).replace(",", "").replace("%", "")
    if text == "":
        return default
    try:
        return float(text)
    except Exception:
        return default


def excel_col_to_index(column: str) -> int:
    result = 0
    for char in column.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1


def ensure_min_columns(df_raw: pd.DataFrame, required_indices: list[int], label: str):
    if df_raw.shape[1] <= max(required_indices):
        raise DataValidationError(
            f"{label} 열 개수가 부족합니다. 필요한 최대 열 인덱스: {max(required_indices) + 1}, 실제 열 수: {df_raw.shape[1]}"
        )


def normalize_lookup_key(value) -> str:
    return clean_text(value).upper()


def normalize_code_lookup_key(value) -> str:
    return re.sub(r"\s+", "", clean_text(value)).upper()


def normalize_brand_series(series: pd.Series) -> pd.Series:
    normalized = (
        series.fillna("").astype(str).str.strip().str.upper()
        .str.replace("（", "(", regex=False)
        .str.replace("）", ")", regex=False)
        .str.replace("\\", "/", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )

    result = pd.Series("", index=series.index, dtype=str)
    result = result.mask(normalized.str.contains("NORTH FACE.*\\(DC\\)|NORTH FACE.*DC|\\(DC\\).*NORTH FACE", na=False), "THE NORTH FACE (DC)")
    result = result.mask(normalized.str.contains("NORTH FACE", na=False) & result.eq(""), "THE NORTH FACE")
    result = result.mask(normalized.str.contains("NIKE|나이키", na=False), "NIKE")
    result = result.mask(normalized.str.contains("ASICS|아식스", na=False), "ASICS")
    result = result.mask(normalized.str.contains("CONVERSE|CONBERSE|컨버스", na=False), "CONVERSE")
    result = result.mask(normalized.str.contains("ADIDAS", na=False), "ADIDAS")
    result = result.mask(normalized.str.contains("FILA", na=False), "FILA")
    result = result.mask(normalized.str.contains("PUMA", na=False), "PUMA")
    return result


def strip_leading_1100(value: str) -> str:
    return value[4:] if value.startswith("1100") else value


def normalize_product_code(value, brand: str) -> str:
    text = clean_text(value)
    brand_text = clean_text(brand).upper()

    if brand_text == "PUMA":
        return text.replace("(黑色标)", "").replace("鞋", "")

    if brand_text == "FILA":
        text = strip_leading_1100(text)
        return text.replace("服", "").replace("鞋", "").replace("_", "").replace("-", "")

    if brand_text.startswith("THE NORTH FACE"):
        text = strip_leading_1100(text)
        return text.replace("包", "").replace("_", "").replace("-", "")

    return text


APPAREL_SIZE_RE = r"(?:W(?:XS|S|M|L|XL|XXL|XXXL|[2-9]XL)|[2-9]XL|XXXXL|XXXL|XXL|XL|XS|S|M|L|FREE|ONE|OS)"
SIZE_TOKEN_RE = rf"(?<![A-Z0-9])(?:{APPAREL_SIZE_RE}|F|[A-Z]{{1,2}}\d{{2,3}}|\d{{1,3}}(?:\.\d+)?)(?![A-Z0-9])"


def remove_chinese(value: str) -> str:
    return clean_text(re.sub(r"[\u4e00-\u9fff]+", "", value))


def normalize_fraction_marks(value: str) -> str:
    return clean_text(str(value).replace("⅔", ".5").replace("½", ".5").replace("⅓", ".5"))


def remove_size_descriptions(value: str) -> str:
    text = clean_text(value)
    text = re.sub(r"^(?:size\s+)?specification\s+", "", text, flags=re.IGNORECASE)
    return clean_text(text)


def is_apparel_size(value: str) -> bool:
    return bool(re.fullmatch(APPAREL_SIZE_RE, clean_text(value), flags=re.IGNORECASE))


def is_size_token(value: str) -> bool:
    return bool(re.fullmatch(SIZE_TOKEN_RE, clean_text(value), flags=re.IGNORECASE))


def pick_size_token(value: str) -> str:
    text = normalize_fraction_marks(remove_size_descriptions(remove_chinese(value)))
    if "/" in text:
        text = text.rsplit("/", 1)[-1]

    text = clean_text(text)
    if is_size_token(text):
        return text.upper() if re.fullmatch(r"[A-Za-z]+", text) else text

    for token in re.findall(SIZE_TOKEN_RE, text, flags=re.IGNORECASE):
        token = clean_text(token)
        if is_size_token(token):
            return token.upper() if re.fullmatch(r"[A-Za-z]+", token) else token

    return text


def extract_size_detail(value: str) -> str:
    text = clean_text(value)
    bracket_match = re.search(r"\(([^()]*)\)", text)
    if bracket_match:
        before_bracket = pick_size_token(text[:bracket_match.start()])
        bracket_text = pick_size_token(bracket_match.group(1) or "")

        if is_apparel_size(before_bracket):
            text = before_bracket
        elif bracket_text and is_size_token(bracket_text):
            text = bracket_text
        else:
            text = re.sub(r"\([^()]*\)", "", text)

    return pick_size_token(text)


def finalize_size_value(size_value: str, brand: str) -> str:
    size_text = clean_text(size_value)
    brand_text = clean_text(brand).upper()

    if brand_text == "ADIDAS" and size_text == "37":
        return "37.5"
    if brand_text.startswith("THE NORTH FACE") and size_text.upper() == "F":
        return "ONE"
    if brand_text.startswith("THE NORTH FACE") and size_text.upper() in {"2XL", "3XL", "4XL"}:
        return {"2XL": "XXL", "3XL": "XXXL", "4XL": "XXXXL"}[size_text.upper()]
    return size_text


def normalize_size_value(value, brand: str = "") -> str:
    text = clean_text(value)
    if re.search(r"\d+\s*寸|\d+\s*[lL]\b", text):
        return "ONE"

    match = re.search(r"\b(EU|EH|JP|KR|SIZE|KOREA|CHN)\b\s*[:：]?\s*(.+)$", text, flags=re.IGNORECASE)
    if match:
        label = match.group(1).upper()
        size_text = extract_size_detail(match.group(2))
        if size_text == "":
            return "ONE"
        if label == "JP" and re.fullmatch(r"\d{2}(?:\.\d+)?", size_text):
            return str(int(float(size_text) * 10))
        return finalize_size_value(size_text, brand)

    fallback_size = pick_size_token(text)
    if fallback_size and is_size_token(fallback_size):
        return finalize_size_value(fallback_size, brand)
    return "ONE"


def make_output_code(product_code: str, converted_size: str) -> str:
    size_text = re.sub(r"^0+(?=\d)", "", clean_text(converted_size))
    return f"{clean_text(product_code)}{size_text}"


def make_converted_output_code(product_code: str, converted_size2: str, fallback_code: str) -> str:
    if clean_text(converted_size2):
        return f"{clean_text(product_code)}{clean_text(converted_size2)}"
    return clean_text(fallback_code)


def build_size_chart_lookup(df_raw: pd.DataFrame) -> dict:
    ensure_min_columns(df_raw, [0, 1, 2, 3], "사이즈표")
    size_df = df_raw.iloc[:, :4].copy()
    size_df.columns = ["브랜드", "품번", "사이즈", "EU"]
    size_df = size_df.fillna("").astype(str).apply(lambda col: col.map(clean_text))

    lookup = {}
    for brand, product_code, size_value, eu_size in size_df.itertuples(index=False, name=None):
        brand_key = normalize_lookup_key(brand)
        product_code_key = normalize_lookup_key(product_code)
        eu_key = normalize_lookup_key(eu_size)
        if brand_key and eu_key and size_value:
            lookup.setdefault((brand_key, product_code_key, eu_key), clean_text(size_value))
    return lookup


def apply_size_chart_lookup(brand: str, product_code: str, converted_size: str, size_lookup: dict) -> str:
    brand_key = normalize_lookup_key(brand)
    product_code_key = normalize_lookup_key(product_code)
    if brand_key.startswith("THE NORTH FACE"):
        product_code_key = product_code_key[:7]

    key = (brand_key, product_code_key, normalize_lookup_key(converted_size))
    if brand_key.startswith("THE NORTH FACE"):
        base_brand_key = ("THE NORTH FACE", product_code_key, normalize_lookup_key(converted_size))
        return size_lookup.get(key, size_lookup.get(base_brand_key, converted_size))

    fallback_key = (brand_key, "", normalize_lookup_key(converted_size))
    return size_lookup.get(key, size_lookup.get(fallback_key, converted_size))


def build_size_chart_brand_lookup(df_raw: pd.DataFrame) -> dict:
    ensure_min_columns(df_raw, [0, 1], "사이즈표")
    size_df = df_raw.iloc[:, :2].copy()
    size_df.columns = ["브랜드", "품번"]
    size_df = size_df.fillna("").astype(str).apply(lambda col: col.map(clean_text))
    size_df = size_df[size_df["품번"].ne("") & size_df["브랜드"].ne("")]
    size_df["브랜드"] = normalize_brand_series(size_df["브랜드"])
    size_df = size_df[size_df["브랜드"].ne("")]
    size_df = size_df.drop_duplicates(subset=["품번"], keep="first")
    return dict(zip(size_df["품번"], size_df["브랜드"]))


def load_stock_prepare(df_raw: pd.DataFrame) -> pd.DataFrame:
    required_columns = ["코드", "품번", "사이즈"]
    missing = [col for col in required_columns if col not in df_raw.columns]
    if missing:
        raise DataValidationError(f"분배준비 시트 필수 열이 없습니다: {missing}")

    df = df_raw.copy()
    for col in ["변환코드1", "변환코드2", "변환코드3"]:
        if col not in df.columns:
            df[col] = ""

    keep_cols = required_columns + ["변환코드1", "변환코드2", "변환코드3"]
    if "브랜드" in df.columns:
        keep_cols.append("브랜드")
    if "품번_변환" in df.columns:
        keep_cols.append("품번_변환")

    df = df[keep_cols].copy()
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).map(clean_text)

    return df


def build_stock_brand_lookup(stock_df: pd.DataFrame) -> dict:
    if "브랜드" not in stock_df.columns:
        return {}

    temp = stock_df.copy()
    temp["브랜드"] = normalize_brand_series(temp["브랜드"])
    temp = temp[temp["브랜드"].ne("")]

    lookup = {}
    for _, row in temp.iterrows():
        brand = clean_text(row.get("브랜드", ""))
        raw_product_code = clean_text(row.get("품번", ""))
        converted_product_code = clean_text(row.get("품번_변환", ""))
        for product_code in [raw_product_code, converted_product_code, normalize_product_code(raw_product_code, brand)]:
            product_code_key = clean_text(product_code)
            if product_code_key:
                lookup.setdefault(product_code_key, brand)
    return lookup


def build_stock_size_lookup(stock_df: pd.DataFrame) -> dict:
    lookup = {}
    code_columns = ["코드", "변환코드1", "변환코드2", "변환코드3"]
    for row in stock_df[["사이즈"] + code_columns].itertuples(index=False, name=None):
        stock_size = clean_text(row[0])
        if not stock_size:
            continue
        for code in row[1:]:
            code_key = normalize_code_lookup_key(code)
            if code_key:
                lookup.setdefault(code_key, stock_size)
    return lookup


def lookup_brand_from_product_code(product_code: str, stock_brand_lookup: dict) -> str:
    text = clean_text(product_code)
    if text in stock_brand_lookup:
        return stock_brand_lookup[text]
    stripped_text = strip_leading_1100(text)
    if stripped_text in stock_brand_lookup:
        return stock_brand_lookup[stripped_text]
    return ""


def build_platform_close_lookup(df_raw: pd.DataFrame) -> dict:
    required_columns = ["날짜", "플랫폼", "주문번호", "운송장"] + PLATFORM_CLOSE_COLUMNS
    missing = [col for col in required_columns if col not in df_raw.columns]
    if missing:
        raise DataValidationError(f"플랫폼 마감 시트 필수 열이 없습니다: {missing}")

    lookup_df = df_raw[PLATFORM_SOURCE_COLUMNS].copy().fillna("").astype(str)
    lookup_df = lookup_df.apply(lambda col: col.map(clean_text))
    lookup_df = lookup_df.iloc[::-1]
    lookup_df = lookup_df[lookup_df["주문번호"].ne("")]
    lookup_df = lookup_df.drop_duplicates(subset=["주문번호"], keep="first")
    lookup_df = lookup_df.rename(columns={"날짜": "원거래 날짜"})
    lookup_df.insert(0, "날짜", time.strftime("%Y-%m-%d"))
    lookup_df["플랫폼"] = "POIZON"
    return lookup_df.set_index("주문번호", drop=False)[OUTPUT_COLUMNS].to_dict("index")


def read_return_source(df_raw: pd.DataFrame) -> pd.DataFrame:
    use_cols = [excel_col_to_index(col) for col in RETURN_USECOLS]
    ensure_min_columns(df_raw, use_cols, "POIZON 반품 시트")
    output_df = df_raw.iloc[:, use_cols].copy()
    output_df.columns = RETURN_COLUMNS
    output_df = output_df.fillna("").astype(str).apply(lambda col: col.map(clean_text))
    output_df = output_df[output_df["품번원본"].ne("")].copy()
    output_df["수량"] = output_df["수량"].apply(lambda value: int(to_number(value)))
    output_df = output_df[output_df["수량"].gt(0)].copy()
    return output_df.reset_index(drop=True)


def transform_return_df(
    return_df: pd.DataFrame,
    size_lookup: dict,
    stock_size_lookup: dict,
    stock_brand_lookup: dict,
) -> pd.DataFrame:
    output_df = return_df.copy()
    stock_brands = [
        lookup_brand_from_product_code(product_code, stock_brand_lookup)
        for product_code in output_df["품번원본"]
    ]
    source_brands = normalize_brand_series(output_df["브랜드원본"])
    output_df["브랜드"] = [
        stock_brand if clean_text(stock_brand) else source_brand
        for stock_brand, source_brand in zip(stock_brands, source_brands)
    ]
    output_df["품번"] = [
        normalize_product_code(product_code, brand)
        for product_code, brand in zip(output_df["품번원본"], output_df["브랜드"])
    ]
    converted_sizes = [
        normalize_size_value(size_value, brand)
        for size_value, brand in zip(output_df["사이즈원본"], output_df["브랜드"])
    ]
    output_df["사이즈변환"] = [
        apply_size_chart_lookup(brand, product_code, converted_size, size_lookup)
        for brand, product_code, converted_size in zip(output_df["브랜드"], output_df["품번"], converted_sizes)
    ]
    base_codes = [
        make_output_code(product_code, converted_size)
        for product_code, converted_size in zip(output_df["품번"], output_df["사이즈변환"])
    ]
    stock_size_matches = [
        stock_size_lookup.get(normalize_code_lookup_key(code), "")
        for code in base_codes
    ]
    output_df["사이즈"] = [
        stock_size if clean_text(stock_size) else converted_size
        for stock_size, converted_size in zip(stock_size_matches, output_df["사이즈변환"])
    ]
    output_df["코드"] = [
        make_converted_output_code(product_code, stock_size, base_code)
        for product_code, stock_size, base_code in zip(output_df["품번"], stock_size_matches, base_codes)
    ]
    output_df["날짜"] = time.strftime("%Y-%m-%d")
    output_df["원거래 날짜"] = time.strftime("%Y-%m-%d")
    output_df["플랫폼"] = "POIZON"
    for column in ["매장명", "할인가", "총 판매가", "총 수수료", "총 수입", "환율", "총 KRW", "총 공급가", "마진", "내역"]:
        output_df[column] = ""
    return output_df[OUTPUT_COLUMNS].copy()


def apply_platform_close_lookup(return_df: pd.DataFrame, fallback_df: pd.DataFrame, platform_close_lookup: dict) -> pd.DataFrame:
    rows = []
    matched_count = 0

    for idx, return_row in return_df.reset_index(drop=True).iterrows():
        order_no = clean_text(return_row.get("주문번호", ""))
        matched_row = platform_close_lookup.get(order_no) if order_no else None
        if matched_row is not None:
            output_row = dict(matched_row)
            output_row["날짜"] = time.strftime("%Y-%m-%d")
            output_row["플랫폼"] = "POIZON"
            if not clean_text(output_row.get("운송장", "")):
                output_row["운송장"] = clean_text(return_row.get("운송장", ""))
            rows.append(output_row)
            matched_count += 1
        else:
            rows.append(fallback_df.iloc[idx].to_dict())

    log(f"플랫폼 마감 주문번호 매칭 {matched_count}행 / 기존 변환 사용 {len(rows) - matched_count}행")
    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS).fillna("")


def apply_stock_brand_lookup(output_df: pd.DataFrame, stock_brand_lookup: dict) -> pd.DataFrame:
    df = output_df.copy()
    stock_brands = [
        lookup_brand_from_product_code(product_code, stock_brand_lookup)
        for product_code in df["품번"]
    ]
    df["브랜드"] = [
        stock_brand if clean_text(stock_brand) else current_brand
        for stock_brand, current_brand in zip(stock_brands, df["브랜드"])
    ]
    return df


def build_output_df() -> pd.DataFrame:
    log("CSV 병렬 다운로드")
    csvs = download_all_csvs()

    return_raw = read_google_sheet_csv_from_text(csvs["return"])
    platform_close_raw = read_google_sheet_csv_from_text(csvs["platform_close"])
    size_chart_raw = read_google_sheet_csv_from_text(csvs["size_chart"])
    stock_prep_raw = read_google_sheet_csv_from_text(csvs["stock_prep"])

    return_df = read_return_source(return_raw)
    size_lookup = build_size_chart_lookup(size_chart_raw)
    stock_df = load_stock_prepare(stock_prep_raw)
    stock_brand_lookup = build_stock_brand_lookup(stock_df)
    stock_size_lookup = build_stock_size_lookup(stock_df)
    platform_close_lookup = build_platform_close_lookup(platform_close_raw)

    fallback_df = transform_return_df(
        return_df,
        size_lookup=size_lookup,
        stock_size_lookup=stock_size_lookup,
        stock_brand_lookup=stock_brand_lookup,
    )
    output_df = apply_platform_close_lookup(return_df, fallback_df, platform_close_lookup)
    output_df = apply_stock_brand_lookup(output_df, stock_brand_lookup)
    if output_df.empty:
        raise DataValidationError("변환할 POIZON 반품 데이터가 없습니다.")
    missing_brand_df = output_df[output_df["브랜드"].eq("")].copy()
    if not missing_brand_df.empty:
        missing_product_codes = ", ".join(sorted(missing_brand_df["품번"].dropna().astype(str).unique()))
        log(f"브랜드 미매칭 {len(missing_brand_df)}행 -> 현재고 변환 시트에 품번 없음: {missing_product_codes}")
    return output_df


def upload_to_google_sheet(df: pd.DataFrame):
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": SHEET_NAME,
        "values": [df.columns.tolist()] + df.fillna("").astype(str).values.tolist(),
    }

    last_error = None
    for attempt in range(1, UPLOAD_RETRIES + 1):
        try:
            log(f"업로드 시작 -> {SHEET_NAME} / {len(df)}행")
            response = requests.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            log(f"업로드 완료 -> {response.text}")
            return
        except requests.exceptions.RequestException as exc:
            last_error = exc
            log(f"업로드 실패 ({attempt}/{UPLOAD_RETRIES}): {exc}")
            if attempt < UPLOAD_RETRIES:
                time.sleep(UPLOAD_SLEEP_SEC)
    raise last_error


def main():
    start_time = time.perf_counter()
    output_df = build_output_df()
    upload_to_google_sheet(output_df)
    log(f"완료 ({time.perf_counter() - start_time:.2f}초)")


if __name__ == "__main__":
    main()
