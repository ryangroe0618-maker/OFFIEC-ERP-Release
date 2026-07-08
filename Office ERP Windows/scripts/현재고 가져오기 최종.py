# -*- coding: utf-8 -*-
import time
import re
import json
from io import StringIO
from concurrent.futures import ThreadPoolExecutor, as_completed

import numpy as np
import pandas as pd
import requests
from pandas.errors import EmptyDataError


# =========================
# 업로드 설정
# =========================
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyKiv4r8GjUPIsX5Tz7nVlpeTw7WIqPHsh8nvVw9WqKPihQSIKk4Vg4QMaO8VfkS0aN/exec"
SPREADSHEET_ID = "1aK2IZzdfsEx8YBd0G4oSUg3GPUmDJ0nas9AUAVTNopE"
TARGET_SHEET_NAME = "통합재고_TEST"

MIRROR_WEB_APP_URL = "https://script.google.com/macros/s/AKfycby2c66vK1s9KYY4hSggVRm1hb6sPWSK7PStHgH1424meynewV5GGPF2-3iLBj1AohNSIA/exec"
MIRROR_SPREADSHEET_ID = "1vcEzHur60Zp4clbABWrbKajxBlFeUOvYcyyQ_0vlPCw"
MIRROR_SHEET_NAME = "현재고"

# =========================
# URL
# =========================
URL_OFFICE = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSwUJyMJP8Qdjx3P65fzeHw61SNPJ6fkkms-TEqWvAKZV9A8jrzmjWLt16i5u8Xbv2xG2AGCbwf855o/pub?gid=783524871&single=true&output=csv"
URL_OFFICE_SELF = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSwUJyMJP8Qdjx3P65fzeHw61SNPJ6fkkms-TEqWvAKZV9A8jrzmjWLt16i5u8Xbv2xG2AGCbwf855o/pub?gid=0&single=true&output=csv"
URL_OFFICE_SQ = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSwUJyMJP8Qdjx3P65fzeHw61SNPJ6fkkms-TEqWvAKZV9A8jrzmjWLt16i5u8Xbv2xG2AGCbwf855o/pub?gid=1910705471&single=true&output=csv"
URL_OFFICE_GUWOL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSwUJyMJP8Qdjx3P65fzeHw61SNPJ6fkkms-TEqWvAKZV9A8jrzmjWLt16i5u8Xbv2xG2AGCbwf855o/pub?gid=1873164003&single=true&output=csv"
URL_OFFICE_BUCHEON = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSwUJyMJP8Qdjx3P65fzeHw61SNPJ6fkkms-TEqWvAKZV9A8jrzmjWLt16i5u8Xbv2xG2AGCbwf855o/pub?gid=987924910&single=true&output=csv"
URL_OFFICE_ADIDAS_KIDS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSwUJyMJP8Qdjx3P65fzeHw61SNPJ6fkkms-TEqWvAKZV9A8jrzmjWLt16i5u8Xbv2xG2AGCbwf855o/pub?gid=987327973&single=true&output=csv"
URL_OFFICE_PUMA_YEOJU = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSwUJyMJP8Qdjx3P65fzeHw61SNPJ6fkkms-TEqWvAKZV9A8jrzmjWLt16i5u8Xbv2xG2AGCbwf855o/pub?gid=1819712722&single=true&output=csv"
URL_OFFICE_RETURN = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSwUJyMJP8Qdjx3P65fzeHw61SNPJ6fkkms-TEqWvAKZV9A8jrzmjWLt16i5u8Xbv2xG2AGCbwf855o/pub?gid=1443331117&single=true&output=csv"
URL_GUWOL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSHTZiYHkTrDlZ_pi1qxBsikvBAaMxtdEzwSYsWzk6sV1zk04SIYjflfnxMYRsmwevPovu4Mtnlx69M/pub?gid=1240644793&single=true&output=csv"
URL_BUCHEON = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQdYpr-dKLe-tguI2uOaYL9pjalY0jehboc1zb-B5XKbV8vAPQvtw1S4nu-TaxJULDsoKOTz8gz7A5y/pub?gid=1240644793&single=true&output=csv"
URL_SQ = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTPSUW1W1iSIvGGrLkp1WHj6Dy_k4NQHv5xOZR4xviYMsZWUb6ZBQ4PqeI31RM_keSDaXeQsYyNLAav/pub?gid=1240644793&single=true&output=csv"
URL_FILA = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSA47SgFq9QQPg0D3AlBnpJX6q7Yx_Dh66E1ID9MlXTahJjL0FmFVtPgyTEtj4iVj7PvRkCUoCgbjkd/pub?gid=0&single=true&output=csv"
URL_PUMA = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSzkIBQ7UfqnUboNBWaQj6esNZzi_NSk0crAVPCljFog-YAnl1vSY6gqqTxH2CYosDoRL4q2PgMUhqL/pub?gid=1929205311&single=true&output=csv"

URL_MASTER = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=0&single=true&output=csv"
URL_NF_DISCOUNT = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=1991382423&single=true&output=csv"
URL_PUMA_DISCOUNT = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=504358853&single=true&output=csv"
URL_SUPPLY = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=1616917746&single=true&output=csv"
URL_SIZE_MAP = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=1121158649&single=true&output=csv"
URL_NF_DC_ITEMS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/pub?gid=599962083&single=true&output=csv"

# =========================
# 타임아웃 / 재시도
# =========================
MASTER_TIMEOUT = (5, 40)
MASTER_RETRIES = 3

DISCOUNT_TIMEOUT = (5, 15)
DISCOUNT_RETRIES = 2

SUPPLY_TIMEOUT = (5, 15)
SUPPLY_RETRIES = 2

SIZE_MAP_TIMEOUT = (5, 15)
SIZE_MAP_RETRIES = 2

OFFICE_TIMEOUT = (10, 60)
OFFICE_RETRIES = 4

NORMAL_TIMEOUT = (5, 20)
NORMAL_RETRIES = 3

UPLOAD_TIMEOUT = (10, 300)
UPLOAD_RETRIES = 3
UPLOAD_SLEEP_SEC = 3

START_TIME = time.time()


class DataValidationError(ValueError):
    pass


def log(msg: str):
    sec = time.time() - START_TIME
    print(f"[{sec:6.1f}s] {msg}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(session, url, label="", timeout=(5, 20), retries=3, sleep_sec=2):
    last_error = None

    for attempt in range(1, retries + 1):
        try:
            if label:
                log(f"{label} 다운로드 시작 ({attempt}/{retries})")
            r = session.get(url, timeout=timeout)
            r.raise_for_status()
            r.encoding = "utf-8"
            if label:
                log(f"{label} 다운로드 완료")
            return r.text

        except Exception as e:
            last_error = e
            if label:
                log(f"{label} 다운로드 실패 ({attempt}/{retries}): {e}")
            if attempt < retries:
                time.sleep(sleep_sec)

    raise last_error


def normalize_brand(value):
    s = clean_text(value).upper()
    if s == "":
        return ""
    if "ASICS" in s or "아식스" in s:
        return "ASICS"
    if "CONVERSE" in s or "컨버스" in s:
        return "CONVERSE"
    if "NIKE" in s or "나이키" in s:
        return "NIKE"
    if "NORTHFACE" in s or "NORTH FACE" in s or "노스페이스" in s:
        return "THE NORTH FACE"
    if "PUMA" in s or "푸마" in s:
        return "PUMA"
    if "FILA" in s or "휠라" in s:
        return "FILA"
    if "ADIDAS" in s or "아디다스" in s:
        return "ADIDAS"
    return s


def clean_text(value):
    if pd.isna(value):
        return ""
    s = str(value).strip()
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("\\", "/")
    s = re.sub(r"\s+", "", s)
    return s


def normalize_size(value):
    if pd.isna(value):
        return ""

    s = str(value).strip()
    if s == "":
        return ""

    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("\\", "/")
    s = s.replace("　", " ")
    s = re.sub(r"\s+", " ", s).strip().upper()

    compact = s.replace(" ", "")
    size_aliases = {
        "ONE SIZE": "ONE SIZE",
        "ONESIZE": "ONE SIZE",
        "ONE-SIZE": "ONE SIZE",
        "FREE": "FREE",
        "F": "FREE",
    }
    if compact in size_aliases:
        return size_aliases[compact]

    mm_match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*MM", s)
    if mm_match:
        num = float(mm_match.group(1))
        normalized = str(int(num)) if num.is_integer() else str(num).rstrip("0").rstrip(".")
        return f"{normalized}MM"

    num_match = re.fullmatch(r"\d+(?:\.\d+)?", s)
    if num_match:
        num = float(s)
        return str(int(num)) if num.is_integer() else str(num).rstrip("0").rstrip(".")

    return s


def normalize_brand_series(series: pd.Series) -> pd.Series:
    s = series.fillna("").astype(str).str.strip().str.upper()

    result = s.copy()
    result = result.mask(s.str.contains("ASICS|아식스", na=False), "ASICS")
    result = result.mask(s.str.contains("CONVERSE|컨버스", na=False), "CONVERSE")
    result = result.mask(s.str.contains("NIKE|나이키", na=False), "NIKE")
    result = result.mask(s.str.contains("NORTH FACE|노스페이스", na=False), "THE NORTH FACE")
    result = result.mask(s.str.contains("PUMA|푸마", na=False), "PUMA")
    result = result.mask(s.str.contains("FILA|휠라", na=False), "FILA")
    result = result.mask(s.str.contains("ADIDAS|아디다스", na=False), "ADIDAS")
    result = result.mask(s == "", "")
    return result


def build_aux_item_no_series(brand_series: pd.Series, item_no_series: pd.Series) -> pd.Series:
    brand = brand_series.fillna("").astype(str).str.strip().str.upper()
    item_no = item_no_series.fillna("").astype(str).str.strip()

    aux = item_no.copy()

    nf_mask = brand == "THE NORTH FACE"
    aux.loc[nf_mask] = item_no.loc[nf_mask].str.slice(0, 7)

    puma_mask = brand == "PUMA"
    puma_left = item_no.loc[puma_mask].str.slice(0, 6)
    aux.loc[puma_mask] = puma_left.apply(
        lambda x: str(int(x)) if x.isdigit() else x
    )

    fila_mask = brand == "FILA"
    fila_item = item_no.loc[fila_mask]
    aux.loc[fila_mask] = np.where(
        fila_item.str.startswith("F"),
        fila_item.str.slice(0, 13),
        fila_item.str.slice(0, 9)
    )

    return aux.fillna("").astype(str)


def require_columns(df: pd.DataFrame, required_cols, label: str):
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise DataValidationError(f"{label} 필수 컬럼 누락: {missing}")


def validate_stock_frame(df: pd.DataFrame, label: str):
    if df.empty:
        raise DataValidationError(f"{label} 데이터가 비어 있습니다.")

    if df["품번"].eq("").all():
        raise DataValidationError(f"{label} 품번 데이터가 비어 있습니다.")

    numeric_qty = pd.to_numeric(df["수량"], errors="coerce")
    if numeric_qty.notna().sum() == 0:
        raise DataValidationError(f"{label} 수량 컬럼이 숫자로 인식되지 않습니다.")


def clean_stock_df(df):
    df = df.copy()
    df["브랜드"] = df["브랜드"].apply(normalize_brand)
    df["품번"] = df["품번"].apply(clean_text)
    df["사이즈"] = df["사이즈"].apply(normalize_size)
    header_like_brands = {"브랜드"}
    header_like_items = {"품번", "품번원본", "상품코드", "코드"}
    header_like_sizes = {"사이즈", "SIZE"}

    df = df[
        ~df["브랜드"].isin(header_like_brands)
        & ~df["품번"].isin(header_like_items)
        & ~df["사이즈"].isin(header_like_sizes)
    ].copy()

    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)
    df = df[df["품번"] != ""].copy()
    return df.reset_index(drop=True)


def empty_stock_df():
    return pd.DataFrame(columns=["브랜드", "품번", "사이즈", "수량"])


def load_stock_sheet(csv_text, start_row, cols, column_names, label, allow_empty=False):
    try:
        df = parse_read_csv_cols(csv_text, start_row, cols, label=label)
    except DataValidationError as e:
        if allow_empty and "데이터가 비어 있습니다." in str(e):
            log(f"{label} 데이터 없음 -> 빈 시트로 처리")
            return empty_stock_df()
        raise

    df.columns = column_names
    df = clean_stock_df(df)

    try:
        validate_stock_frame(df, label)
    except DataValidationError as e:
        if allow_empty and "데이터가 비어 있습니다." in str(e):
            log(f"{label} 유효 데이터 없음 -> 빈 시트로 처리")
            return empty_stock_df()
        raise

    return df


def add_source(df, source_name):
    temp = df.copy()
    temp["출처"] = source_name
    return temp


def format_discount_percent(value):
    s = str(value).strip().replace("%", "").replace(",", "")
    if s == "":
        return ""
    try:
        num = float(s)
    except Exception:
        return ""
    if 0 <= num <= 1:
        num *= 100
    if float(num).is_integer():
        return f"{int(num)}%"
    return f"{num:.2f}".rstrip("0").rstrip(".") + "%"


def display_stock_vectorized(series):
    s = pd.to_numeric(series, errors="coerce")
    s = s.mask(s == 0)

    def _fmt(x):
        if pd.isna(x):
            return ""
        if float(x).is_integer():
            return int(x)
        return x

    return s.apply(_fmt)


def build_sale_price(df):
    temp = df.copy()

    temp["최초가_num"] = pd.to_numeric(
        temp["최초가"].astype(str).str.replace(",", "", regex=False),
        errors="coerce"
    )

    sheet_sale_price = pd.to_numeric(
        temp.get(
            "할인가_시트",
            pd.Series("", index=temp.index),
        )
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip(),
        errors="coerce",
    )

    discount_raw = (
        temp["할인율"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    temp["할인율_num"] = pd.to_numeric(discount_raw, errors="coerce")

    temp["할인율_ratio"] = np.where(
        temp["할인율_num"] > 1,
        temp["할인율_num"] / 100,
        temp["할인율_num"]
    )

    sale_price = np.where(
        sheet_sale_price.notna(),
        sheet_sale_price,
        np.where(
            pd.isna(temp["할인율_ratio"]),
            temp["최초가_num"],
            temp["최초가_num"] * (1 - temp["할인율_ratio"])
        ),
    )
    sale_price = pd.Series(sale_price, index=temp.index).round()

    def _fmt(x):
        if pd.isna(x):
            return ""
        return int(x)

    return sale_price.apply(_fmt)


def build_supply_price_vectorized(df, puma_extra_ratio=1.0):
    temp = df.copy()

    temp["최초가_num"] = pd.to_numeric(
        temp["최초가"].astype(str).str.replace(",", "", regex=False),
        errors="coerce"
    )

    temp["공급가_파일_num"] = pd.to_numeric(
        temp["공급가_파일"].astype(str).str.replace(",", "", regex=False),
        errors="coerce"
    )

    discount_raw = (
        temp["할인율"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    temp["할인율_num"] = pd.to_numeric(discount_raw, errors="coerce")

    temp["할인율_ratio"] = np.where(
        temp["할인율_num"] > 1,
        temp["할인율_num"] / 100,
        temp["할인율_num"]
    )

    brand = temp["브랜드"].astype(str).str.upper().str.strip()
    first_price = temp["최초가_num"]
    discount_ratio = pd.Series(temp["할인율_ratio"], index=temp.index, dtype="float64")
    sale_price = pd.to_numeric(
        temp["할인가"].astype(str).str.replace(",", "", regex=False),
        errors="coerce",
    )
    file_supply = temp["공급가_파일_num"]

    result = pd.Series(np.nan, index=temp.index, dtype="float64")

    has_file_supply = file_supply.notna()
    result.loc[has_file_supply] = file_supply.loc[has_file_supply]

    need_calc = (~has_file_supply) & first_price.notna()

    nf_mask = need_calc & (brand == "THE NORTH FACE")

    nf_factor = pd.Series(
        np.select(
            [
                discount_ratio.isna(),
                discount_ratio <= 0.1,
                discount_ratio <= 0.2,
                discount_ratio <= 0.3,
                discount_ratio > 0.3,
            ],
            [0.67 * 0.95, 0.67, 0.71, 0.71, 0.74],
            default=np.nan
        ),
        index=temp.index,
        dtype="float64"
    )

    result.loc[nf_mask] = (
        sale_price.loc[nf_mask]
        * nf_factor.loc[nf_mask]
    )

    fila_mask = need_calc & (brand == "FILA")
    result.loc[fila_mask] = first_price.loc[fila_mask] * 0.567

    puma_mask = need_calc & (brand == "PUMA")
    puma_multiplier = pd.Series(
        np.where(discount_ratio.isna(), 1, 1 - discount_ratio.fillna(0)),
        index=temp.index,
        dtype="float64"
    )
    result.loc[puma_mask] = (
        first_price.loc[puma_mask]
        * puma_multiplier.loc[puma_mask]
        * 0.4
        * puma_extra_ratio
    )

    adidas_mask = need_calc & (brand == "ADIDAS")
    result.loc[adidas_mask] = first_price.loc[adidas_mask] * 0.5

    result = result.round()

    def _fmt(x):
        if pd.isna(x):
            return ""
        return int(x)

    return result.apply(_fmt)


def parse_read_csv_cols(csv_text, start_row, cols, label=""):
    if not str(csv_text).strip():
        raise DataValidationError(
            f"{label or '재고 시트'} 데이터가 비어 있습니다. 시작 행({start_row}) 또는 열({cols})을 확인해주세요."
        )

    try:
        df = pd.read_csv(
            StringIO(csv_text),
            skiprows=start_row - 1,
            dtype=str,
            header=None
        )
    except EmptyDataError as e:
        raise DataValidationError(
            f"{label or '재고 시트'} 데이터가 비어 있습니다. 시작 행({start_row}) 또는 열({cols})을 확인해주세요."
        ) from e

    df = df.fillna("")
    idx = [ord(c.upper()) - 65 for c in cols]
    max_idx = max(idx)
    if df.shape[1] <= max_idx:
        raise DataValidationError(
            f"{label or '재고 시트'} 열 위치 오류: 요청 열 {cols} 를 읽을 수 없습니다. "
            f"(실제 열 수: {df.shape[1]})"
        )
    df = df.iloc[:, idx]
    if df.shape[1] != len(cols):
        raise DataValidationError(
            f"{label or '재고 시트'} 열 수 오류: 기대 {len(cols)}개, 실제 {df.shape[1]}개"
        )
    non_empty_mask = df.astype(str).apply(lambda col: col.str.strip()).ne("").any(axis=1)
    df = df[non_empty_mask]
    if df.empty:
        raise DataValidationError(
            f"{label or '재고 시트'} 데이터가 비어 있습니다. 시작 행({start_row}) 또는 열({cols})을 확인해주세요."
        )
    return df.reset_index(drop=True)


def parse_master(csv_text):
    master = pd.read_csv(StringIO(csv_text), dtype=str).fillna("")
    needed_cols = ["코드", "바코드", "품번", "상품명", "컬러", "사이즈", "최초가"]
    require_columns(master, needed_cols, "마스터")
    master = master[needed_cols].copy()
    for col in ["바코드", "품번"]:
        master[col] = master[col].apply(clean_text)
    master["사이즈"] = master["사이즈"].apply(normalize_size)
    master["코드"] = master["품번"] + master["사이즈"]
    for col in ["상품명", "컬러", "최초가"]:
        master[col] = master[col].astype(str).str.strip()
    master = master[master["코드"] != ""].copy()
    if master.empty:
        raise DataValidationError("마스터 유효 데이터가 없습니다.")
    return master.drop_duplicates(subset=["코드"], keep="first").reset_index(drop=True)


def parse_discount(csv_text, header_row):
    df = pd.read_csv(
        StringIO(csv_text),
        dtype=str,
        skiprows=header_row - 1
    ).fillna("")
    df = df.iloc[:, :2].copy()
    df.columns = ["품번", "할인율"]
    df["품번"] = df["품번"].apply(clean_text)
    df["할인율"] = df["할인율"].astype(str).str.strip()
    df = df[df["품번"] != ""].copy()
    return df.drop_duplicates(subset=["품번"], keep="first").reset_index(drop=True)


def parse_nf_discount(csv_text):
    df = pd.read_csv(
        StringIO(csv_text),
        dtype=str,
        header=0,
        keep_default_na=False,
    ).fillna("")
    if df.shape[1] < 6:
        raise DataValidationError(
            "노스페이스 할인 시트의 A~F열을 확인해 주세요."
        )

    df = df.iloc[:, :6].copy()
    df.columns = [
        "품번",
        "최초가",
        "기준가",
        "할인가_시트",
        "할인율",
        "최초가대비 할인금액",
    ]
    df["품번"] = df["품번"].apply(clean_text)
    for col in ["할인가_시트", "할인율"]:
        df[col] = df[col].astype(str).str.strip()
    df = df[df["품번"] != ""].copy()
    return (
        df[["품번", "할인율", "할인가_시트"]]
        .drop_duplicates(subset=["품번"], keep="first")
        .reset_index(drop=True)
    )


def parse_puma_c1_ratio(csv_text):
    df = pd.read_csv(
        StringIO(csv_text),
        dtype=str,
        header=None,
        keep_default_na=False,
    ).fillna("")

    if df.shape[1] < 3 or df.shape[0] < 1:
        return 1.0

    raw_value = clean_text(df.iat[0, 2]).replace("%", "").replace(",", "")
    if raw_value == "":
        return 1.0

    try:
        return 1 - (float(raw_value) / 100)
    except Exception:
        return 1.0


def parse_supply(csv_text):
    df = pd.read_csv(StringIO(csv_text), dtype=str).fillna("")
    df = df.iloc[:, :2].copy()
    df.columns = ["품번", "공급가"]
    df["품번"] = df["품번"].apply(clean_text)
    df["공급가"] = df["공급가"].astype(str).str.strip()
    df = df[df["품번"] != ""].copy()
    return df.drop_duplicates(subset=["품번"], keep="first").reset_index(drop=True)


def parse_size_map(csv_text):
    df = pd.read_csv(StringIO(csv_text), dtype=str).fillna("")
    if df.shape[1] < 4:
        raise DataValidationError("사이즈표 컬럼 수가 부족합니다.")
    df = df.iloc[:, :4].copy()
    df.columns = ["브랜드", "품번", "사이즈", "EU"]
    df["브랜드"] = df["브랜드"].apply(normalize_brand)
    df["품번"] = df["품번"].apply(lambda x: clean_text(x)[:7])
    df["사이즈"] = df["사이즈"].apply(normalize_size)
    df["EU"] = df["EU"].apply(normalize_size)
    df = df[(df["브랜드"] != "") & (df["사이즈"] != "")].copy()
    if df.empty:
        raise DataValidationError("사이즈표 유효 데이터가 없습니다.")
    return df.drop_duplicates(subset=["브랜드", "품번", "사이즈"], keep="first").reset_index(drop=True)


def add_eu_column(df, size_map_df):
    result = df.copy()
    result["EU"] = ""

    size_lookup = size_map_df.drop_duplicates(
        subset=["브랜드", "사이즈"],
        keep="first"
    )[["브랜드", "사이즈", "EU"]]

    size_matched = result.merge(
        size_lookup.rename(columns={"EU": "EU_사이즈"}),
        on=["브랜드", "사이즈"],
        how="left"
    )

    size_brand_mask = result["브랜드"].isin(["PUMA", "ADIDAS", "FILA"])
    result.loc[size_brand_mask, "EU"] = (
        size_matched.loc[size_brand_mask, "EU_사이즈"].fillna("").astype(str).str.strip()
    )

    nf_lookup = size_map_df[size_map_df["브랜드"] == "THE NORTH FACE"].drop_duplicates(
        subset=["품번", "사이즈"],
        keep="first"
    )[["품번", "사이즈", "EU"]].rename(columns={"품번": "품번_7", "EU": "EU_품번"})

    nf_source = result[["품번", "사이즈"]].copy()
    nf_source["품번_7"] = nf_source["품번"].apply(lambda x: clean_text(x)[:7])

    nf_matched = nf_source.merge(
        nf_lookup,
        on=["품번_7", "사이즈"],
        how="left"
    )

    nf_brand_mask = result["브랜드"].isin(["THE NORTH FACE", "THE NORTH FACE (DC)"])
    result.loc[nf_brand_mask, "EU"] = (
        nf_matched.loc[nf_brand_mask, "EU_품번"].fillna("").astype(str).str.strip()
    )

    result["EU"] = result["EU"].fillna("").astype(str).str.strip()
    result["EU"] = result["EU"].mask(
        result["EU"] == "",
        result["사이즈"].fillna("").astype(str).str.strip()
    )
    return result


def parse_nf_dc_items(csv_text):
    df = pd.read_csv(StringIO(csv_text), dtype=str).fillna("")
    if df.shape[1] < 1:
        raise DataValidationError("노스페이스 다년차 품번 시트 컬럼 수가 부족합니다.")
    df = df.iloc[:, :1].copy()
    df.columns = ["품번"]
    df["품번"] = df["품번"].apply(clean_text)
    df = df[df["품번"] != ""].copy()
    if df.empty:
        raise DataValidationError("노스페이스 다년차 품번 시트 유효 데이터가 없습니다.")
    return df.drop_duplicates(subset=["품번"], keep="first").reset_index(drop=True)


def upload_to_google_sheet(df, web_app_url, spreadsheet_id, sheet_name):
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    payload = {
        "spreadsheetId": spreadsheet_id,
        "sheetName": sheet_name,
        "values": values,
    }
    log(
        f"구글 시트 업로드 시작 -> {sheet_name} ({len(values) - 1}행) / "
        f"spreadsheetId={spreadsheet_id}"
    )
    last_error = None

    for attempt in range(1, UPLOAD_RETRIES + 1):
        try:
            r = requests.post(web_app_url, json=payload, timeout=UPLOAD_TIMEOUT)
            r.raise_for_status()
            response_text = (r.text or "").strip()

            if "다음 스크립트 함수(doPost)를 찾을 수 없습니다" in response_text or "script function not found: doPost" in response_text:
                raise RuntimeError(
                    "대상 Apps Script 웹앱에 doPost(e)가 배포되어 있지 않습니다. "
                    f"웹앱 URL: {web_app_url}"
                )

            if "<title>오류</title>" in response_text and "Google Apps Script" in response_text:
                raise RuntimeError(
                    "대상 Apps Script 웹앱이 오류 페이지를 반환했습니다. "
                    f"응답 일부: {response_text[:200]}"
                )

            try:
                response_json = json.loads(response_text)
            except Exception:
                response_json = None

            if isinstance(response_json, dict) and response_json.get("ok") is False:
                raise RuntimeError(
                    f"Apps Script 처리 실패: {response_json.get('message', '알 수 없는 오류')} "
                    f"/ sheetName={sheet_name} / spreadsheetId={spreadsheet_id}"
                )

            log(f"구글 시트 업로드 완료")
            log(f"구글 시트 응답: {response_text}")
            return
        except requests.exceptions.Timeout as e:
            last_error = e
            log(f"구글 시트 업로드 타임아웃 ({attempt}/{UPLOAD_RETRIES}): {e}")
        except requests.exceptions.RequestException as e:
            last_error = e
            log(f"구글 시트 업로드 실패 ({attempt}/{UPLOAD_RETRIES}): {e}")

        if attempt < UPLOAD_RETRIES:
            time.sleep(UPLOAD_SLEEP_SEC)

    raise last_error


def upload_target_safe(df, web_app_url, spreadsheet_id, sheet_name, required=True):
    try:
        upload_to_google_sheet(df, web_app_url, spreadsheet_id, sheet_name)
        return
    except Exception as e:
        message = str(e)
        if "doPost" in message:
            log(
                f"업로드 경고 -> {sheet_name}: 대상 웹앱이 POST를 지원하지 않습니다. "
                f"Apps Script에 doPost(e) 배포가 필요합니다."
            )
        else:
            log(
                f"업로드 경고 -> {sheet_name}: {message} "
                f"/ webAppUrl={web_app_url} / spreadsheetId={spreadsheet_id}"
            )

        if required:
            raise


def upload_to_all_targets(main_df, mirror_df):
    targets = [
        (main_df, WEB_APP_URL, SPREADSHEET_ID, TARGET_SHEET_NAME, True),
        (mirror_df, MIRROR_WEB_APP_URL, MIRROR_SPREADSHEET_ID, MIRROR_SHEET_NAME, False),
    ]

    with ThreadPoolExecutor(max_workers=len(targets)) as executor:
        futures = [
            executor.submit(upload_target_safe, df, web_app_url, spreadsheet_id, sheet_name, required)
            for df, web_app_url, spreadsheet_id, sheet_name, required in targets
        ]
        for future in as_completed(futures):
            future.result()


def parse_office_stock_jobs(csvs):
    jobs = {
        "office": (csvs["office"], 2, ["A", "C", "D", "E"], ["브랜드", "품번", "사이즈", "수량"], "사무실", False),
        "office_self": (csvs["office_self"], 4, ["A", "C", "D", "F"], ["브랜드", "품번", "사이즈", "수량"], "사무실-사무실", True),
        "office_sq": (csvs["office_sq"], 4, ["A", "C", "D", "F"], ["브랜드", "품번", "사이즈", "수량"], "사무실-스퀘어원", True),
        "office_guwol": (csvs["office_guwol"], 4, ["A", "C", "D", "F"], ["브랜드", "품번", "사이즈", "수량"], "사무실-구월", True),
        "office_bucheon": (csvs["office_bucheon"], 4, ["A", "C", "D", "F"], ["브랜드", "품번", "사이즈", "수량"], "사무실-부천", True),
        "office_adidas_kids": (csvs["office_adidas_kids"], 4, ["A", "C", "D", "F"], ["브랜드", "품번", "사이즈", "수량"], "사무실-아디다스 키즈", True),
        "office_puma_yeoju": (csvs["office_puma_yeoju"], 4, ["A", "C", "D", "F"], ["브랜드", "품번", "사이즈", "수량"], "사무실-푸마 여주", True),
        "office_return": (csvs["office_return"], 3, ["A", "C", "D", "F"], ["브랜드", "품번", "사이즈", "수량"], "사무실-반품", True),
    }

    results = {}
    with ThreadPoolExecutor(max_workers=len(jobs)) as executor:
        future_map = {
            executor.submit(load_stock_sheet, *args): key
            for key, args in jobs.items()
        }
        for future in as_completed(future_map):
            results[future_map[future]] = future.result()
    return results


def parse_regular_stock_jobs(csvs):
    def parse_nf_store(csv_text, label):
        df = parse_read_csv_cols(csv_text, 4, ["A", "F", "G"], label=label)
        df.columns = ["품번", "사이즈", "수량"]
        df.insert(0, "브랜드", "THE NORTH FACE")
        df = clean_stock_df(df)
        validate_stock_frame(df, label)
        return df

    def parse_fila_store(csv_text):
        df = parse_read_csv_cols(csv_text, 8, ["C", "H", "R"], label="휠라 파주")
        df.columns = ["품번", "사이즈", "수량"]
        df.insert(0, "브랜드", "FILA")
        df = clean_stock_df(df)
        validate_stock_frame(df, "휠라 파주")
        return df

    def parse_puma_store(csv_text):
        df = parse_read_csv_cols(csv_text, 1, ["C", "E", "F"], label="푸마 여주")
        df.columns = ["품번", "사이즈", "수량"]
        df.insert(0, "브랜드", "PUMA")
        df = clean_stock_df(df)
        validate_stock_frame(df, "푸마 여주")
        return df

    jobs = {
        "guwol": (parse_nf_store, csvs["guwol"], "구월"),
        "bucheon": (parse_nf_store, csvs["bucheon"], "부천"),
        "square": (parse_nf_store, csvs["square"], "스퀘어원"),
        "fila": (parse_fila_store, csvs["fila"]),
        "puma": (parse_puma_store, csvs["puma"]),
    }

    results = {}
    with ThreadPoolExecutor(max_workers=len(jobs)) as executor:
        future_map = {}
        for key, job in jobs.items():
            fn, *args = job
            future_map[executor.submit(fn, *args)] = key
        for future in as_completed(future_map):
            results[future_map[future]] = future.result()
    return results


def download_all_csvs(session):
    jobs = {
        "master": (URL_MASTER, "마스터", MASTER_TIMEOUT, MASTER_RETRIES),
        "nf_discount": (URL_NF_DISCOUNT, "노스페이스 할인율", DISCOUNT_TIMEOUT, DISCOUNT_RETRIES),
        "puma_discount": (URL_PUMA_DISCOUNT, "푸마 할인율", DISCOUNT_TIMEOUT, DISCOUNT_RETRIES),
        "supply": (URL_SUPPLY, "공급가 파일", SUPPLY_TIMEOUT, SUPPLY_RETRIES),
        "size_map": (URL_SIZE_MAP, "사이즈표", SIZE_MAP_TIMEOUT, SIZE_MAP_RETRIES),
        "nf_dc_items": (URL_NF_DC_ITEMS, "노스페이스 다년차 품번", DISCOUNT_TIMEOUT, DISCOUNT_RETRIES),

        "office": (URL_OFFICE, "사무실", OFFICE_TIMEOUT, OFFICE_RETRIES),
        "office_self": (URL_OFFICE_SELF, "사무실-사무실", OFFICE_TIMEOUT, OFFICE_RETRIES),
        "office_sq": (URL_OFFICE_SQ, "사무실-스퀘어원", OFFICE_TIMEOUT, OFFICE_RETRIES),
        "office_guwol": (URL_OFFICE_GUWOL, "사무실-구월", OFFICE_TIMEOUT, OFFICE_RETRIES),
        "office_bucheon": (URL_OFFICE_BUCHEON, "사무실-부천", OFFICE_TIMEOUT, OFFICE_RETRIES),
        "office_adidas_kids": (URL_OFFICE_ADIDAS_KIDS, "사무실-아디다스 키즈", OFFICE_TIMEOUT, OFFICE_RETRIES),
        "office_puma_yeoju": (URL_OFFICE_PUMA_YEOJU, "사무실-푸마 여주", OFFICE_TIMEOUT, OFFICE_RETRIES),
        "office_return": (URL_OFFICE_RETURN, "사무실-반품", OFFICE_TIMEOUT, OFFICE_RETRIES),
        "guwol": (URL_GUWOL, "구월", NORMAL_TIMEOUT, NORMAL_RETRIES),
        "bucheon": (URL_BUCHEON, "부천", NORMAL_TIMEOUT, NORMAL_RETRIES),
        "square": (URL_SQ, "스퀘어원", NORMAL_TIMEOUT, NORMAL_RETRIES),
        "fila": (URL_FILA, "휠라 파주", NORMAL_TIMEOUT, NORMAL_RETRIES),
        "puma": (URL_PUMA, "푸마 여주", NORMAL_TIMEOUT, NORMAL_RETRIES),
    }

    results = {}

    with ThreadPoolExecutor(max_workers=min(len(jobs), 12)) as executor:
        future_map = {
            executor.submit(fetch_csv_text, make_session(), url, label, timeout, retries, 2): key
            for key, (url, label, timeout, retries) in jobs.items()
        }

        for future in as_completed(future_map):
            key = future_map[future]
            results[key] = future.result()

    return results


session = make_session()

log("=== 전체 CSV 병렬 다운로드 시작 ===")
csvs = download_all_csvs(session)
log("=== 전체 CSV 병렬 다운로드 완료 ===")

master = parse_master(csvs["master"])

nf_discount = parse_nf_discount(csvs["nf_discount"])
puma_discount = parse_discount(csvs["puma_discount"], header_row=2)
puma_discount["할인가_시트"] = ""
puma_c1_ratio = parse_puma_c1_ratio(csvs["puma_discount"])

discount_df = pd.concat([nf_discount, puma_discount], ignore_index=True)
discount_df["할인율"] = discount_df["할인율"].apply(format_discount_percent)
discount_df["할인가_시트"] = (
    discount_df["할인가_시트"].fillna("").astype(str).str.strip()
)
discount_df = discount_df.drop_duplicates(subset=["품번"], keep="first").reset_index(drop=True)

supply_df = parse_supply(csvs["supply"])
size_map_df = parse_size_map(csvs["size_map"])
nf_dc_items_df = parse_nf_dc_items(csvs["nf_dc_items"])

log("=== 재고 시트 파싱 시작 ===")
office_results = parse_office_stock_jobs(csvs)
regular_results = parse_regular_stock_jobs(csvs)
log("=== 재고 시트 파싱 완료 ===")

office = office_results["office"]
office_self = office_results["office_self"]
office_sq = office_results["office_sq"]
office_guwol = office_results["office_guwol"]
office_bucheon = office_results["office_bucheon"]
office_adidas_kids = office_results["office_adidas_kids"]
office_puma_yeoju = office_results["office_puma_yeoju"]
office_return = office_results["office_return"]

guwol = regular_results["guwol"]
bucheon = regular_results["bucheon"]
square = regular_results["square"]
fila = regular_results["fila"]
puma = regular_results["puma"]

all_data = pd.concat([
    add_source(office, "사무실"),
    add_source(office_self, "사무실 - 사무실"),
    add_source(square, "스퀘어원"),
    add_source(guwol, "구월"),
    add_source(bucheon, "부천"),
    add_source(fila, "휠라 파주"),
    add_source(puma, "푸마 여주"),
    add_source(office_sq, "사무실 - 스퀘어원"),
    add_source(office_guwol, "사무실 - 구월"),
    add_source(office_bucheon, "사무실 - 부천"),
    add_source(office_adidas_kids, "사무실 - 아디다스 키즈"),
    add_source(office_puma_yeoju, "사무실 - 푸마 여주"),
    add_source(office_return, "사무실 - 반품"),
], ignore_index=True)

pivot = (
    all_data
    .groupby(["브랜드", "품번", "사이즈", "출처"], as_index=False)["수량"]
    .sum()
    .pivot_table(
        index=["브랜드", "품번", "사이즈"],
        columns="출처",
        values="수량",
        fill_value=0,
        aggfunc="sum"
    )
    .reset_index()
)

store_cols = [
    "사무실",
    "사무실 - 사무실",
    "스퀘어원",
    "구월",
    "부천",
    "휠라 파주",
    "푸마 여주",
    "사무실 - 스퀘어원",
    "사무실 - 구월",
    "사무실 - 부천",
    "사무실 - 푸마 여주",
    "사무실 - 아디다스 키즈",
    "사무실 - 반품",
]

aggregate_cols = [
    "사무실",
    "스퀘어원",
    "구월",
    "부천",
    "휠라 파주",
    "푸마 여주",
    "사무실 - 아디다스 키즈",
]

for col in store_cols:
    if col not in pivot.columns:
        pivot[col] = 0

for store_col, office_sub_col in [
    ("스퀘어원", "사무실 - 스퀘어원"),
    ("구월", "사무실 - 구월"),
    ("부천", "사무실 - 부천"),
    ("푸마 여주", "사무실 - 푸마 여주"),
]:
    pivot[store_col] = (
        pd.to_numeric(pivot[store_col], errors="coerce").fillna(0)
        - pd.to_numeric(pivot[office_sub_col], errors="coerce").fillna(0)
    )

pivot["현재고"] = pivot[aggregate_cols].apply(
    lambda col: pd.to_numeric(col, errors="coerce").fillna(0)
).sum(axis=1)

pivot["코드"] = (
    pivot["품번"].astype(str).str.strip()
    + pivot["사이즈"].astype(str).str.strip()
)

master_lookup = master[["코드", "바코드", "상품명", "컬러", "최초가"]].copy()
pivot = pivot.merge(master_lookup, on="코드", how="left")

for col in ["바코드", "상품명", "컬러", "최초가"]:
    pivot[col] = pivot[col].fillna("").astype(str).str.strip()

pivot = pivot.merge(
    discount_df[["품번", "할인율", "할인가_시트"]],
    on="품번",
    how="left"
)
pivot["할인율"] = pivot["할인율"].fillna("").astype(str).str.strip()
pivot["할인가_시트"] = (
    pivot["할인가_시트"].fillna("").astype(str).str.strip()
)
pivot["할인가"] = build_sale_price(pivot)

pivot = pivot.merge(
    supply_df[["품번", "공급가"]].rename(columns={"공급가": "공급가_파일"}),
    on="품번",
    how="left"
)
pivot["공급가_파일"] = pivot["공급가_파일"].fillna("").astype(str).str.strip()

pivot["공급가"] = build_supply_price_vectorized(pivot, puma_extra_ratio=puma_c1_ratio)

nf_dc_item_set = set(nf_dc_items_df["품번"].tolist())
nf_dc_mask = pivot["품번"].astype(str).map(lambda x: clean_text(x) in nf_dc_item_set)
pivot["브랜드"] = np.where(nf_dc_mask, "THE NORTH FACE (DC)", pivot["브랜드"])
pivot = add_eu_column(pivot, size_map_df)

for col in store_cols + ["현재고"]:
    pivot[col] = display_stock_vectorized(pivot[col])

final_cols = [
    "브랜드",
    "코드",
    "바코드",
    "품번",
    "상품명",
    "컬러",
    "사이즈",
    "EU",
    "최초가",
    "할인율",
    "할인가",
    "공급가",
    "현재고",
    "사무실",
    "스퀘어원",
    "구월",
    "부천",
    "휠라 파주",
    "푸마 여주",
    "사무실 - 사무실",
    "사무실 - 스퀘어원",
    "사무실 - 구월",
    "사무실 - 부천",
    "사무실 - 푸마 여주",
    "사무실 - 아디다스 키즈",
    "사무실 - 반품",
]

pivot = pivot[final_cols].copy()
pivot = pivot.sort_values(
    by=["브랜드", "품번", "사이즈"],
    ascending=[True, True, True]
).reset_index(drop=True)

mirror_pivot = pivot.copy()
mirror_pivot.insert(0, "보조 품번", build_aux_item_no_series(mirror_pivot["브랜드"], mirror_pivot["품번"]))

upload_to_all_targets(pivot, mirror_pivot)

print("완료")
