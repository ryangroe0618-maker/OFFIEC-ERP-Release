# -*- coding: utf-8 -*-

from io import StringIO
import time

import pandas as pd
import requests


SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQhP9cP1QdWll3UPE-P-tUAmxFHCEXgQU_IKIDsftokFeyn5Y67OW2Zho5xYN4pwQKvcclbDS98bQum/pub?gid=202616820&single=true&output=csv"
CANCEL_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTsZm7LRlCqZRYE6BqWfqeTH_PqhxT-hTDK4ypOPPnwcTr7hpr98L6_nL19I2k_P-uk5WlOGREye13p/pub?gid=1041081062&single=true&output=csv"
BUYER_RETURN_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTsZm7LRlCqZRYE6BqWfqeTH_PqhxT-hTDK4ypOPPnwcTr7hpr98L6_nL19I2k_P-uk5WlOGREye13p/pub?gid=887766708&single=true&output=csv"
RETURN_PROGRESS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTsZm7LRlCqZRYE6BqWfqeTH_PqhxT-hTDK4ypOPPnwcTr7hpr98L6_nL19I2k_P-uk5WlOGREye13p/pub?gid=2094966301&single=true&output=csv"
PLATFORM_SOURCE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRWhzq2rt_O-YnOwzWnE8L_d8--NBu-EMmDkxnnVy-oATAfCjDIX976b973bY2MiO8gYPTE5WCk-tVv/pub?gid=733480714&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzpkEt8TIsHzW9VvNaXhDW63RoKmqOOA_YEQ-PEnYCH2E8FEuvvXemYot2PPvqyP8DEUw/exec"
SPREADSHEET_ID = "1fIQI2OYrInXnXfHvPfMQeERmTW1LE_MoLsyUf5emeKQ"
TARGET_SHEET_NAME = "KREAM - 반품"

DOWNLOAD_TIMEOUT = (10, 60)
UPLOAD_TIMEOUT = (10, 120)
RETRIES = 3
RETRY_SLEEP_SEC = 2

TRACKING_COL_INDEX = 11  # L열
COMPANY_COL_INDEX = 4  # E열
HEADER_ROW_INDEX = 2  # 3행
TRACKING_HEADER = "운송장"
COMPANY_HEADER = "업체명"
TARGET_COMPANY = "RETURNKREAM"
ORDER_HEADER = "주문번호"
RETURN_TRACKING_HEADER = "반송운송장번호"
ORDER_HEADER_ALIASES = {ORDER_HEADER, "주문/보관번호"}
RETURN_TRACKING_HEADER_ALIASES = {RETURN_TRACKING_HEADER}
ORDER_COL_INDEX = 0  # A열
RETURN_TRACKING_COL_INDEX = 9  # J열
OUTPUT_COLUMNS = [
    "날짜",
    "원거래 날짜",
    "플랫폼",
    "주문번호",
    "운송장",
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
PLATFORM_REQUIRED_COLUMNS = [
    "날짜",
    "플랫폼",
    "주문번호",
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


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[KREAM 반품] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def today_str() -> str:
    return time.strftime("%Y-%m-%d")


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
            log(f"CSV 다운로드 실패 ({attempt}/{RETRIES}) -> {exc}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def build_output_df(csv_text: str) -> pd.DataFrame:
    raw_df = pd.read_csv(
        StringIO(csv_text),
        dtype=str,
        header=None,
        keep_default_na=False,
        skip_blank_lines=False,
    ).fillna("")

    if raw_df.shape[0] <= HEADER_ROW_INDEX:
        raise DataValidationError("출고 시트에 3행 머리글이 없습니다.")
    required_col_index = max(COMPANY_COL_INDEX, TRACKING_COL_INDEX)
    if raw_df.shape[1] <= required_col_index:
        raise DataValidationError(f"출고 시트에 L열이 없습니다. 실제 열 수: {raw_df.shape[1]}")

    company_header = clean_text(raw_df.iat[HEADER_ROW_INDEX, COMPANY_COL_INDEX])
    header = clean_text(raw_df.iat[HEADER_ROW_INDEX, TRACKING_COL_INDEX])
    if company_header != COMPANY_HEADER:
        raise DataValidationError(f"E열 3행 머리글이 '{COMPANY_HEADER}'이 아닙니다: '{company_header}'")
    if header != TRACKING_HEADER:
        raise DataValidationError(f"L열 3행 머리글이 '{TRACKING_HEADER}'이 아닙니다: '{header}'")

    data_df = raw_df.iloc[HEADER_ROW_INDEX + 1 :, [COMPANY_COL_INDEX, TRACKING_COL_INDEX]].copy()
    data_df.columns = [COMPANY_HEADER, TRACKING_HEADER]
    company_series = data_df[COMPANY_HEADER].apply(lambda value: clean_text(value).upper())
    tracking_series = data_df[TRACKING_HEADER].apply(clean_text)
    tracking_series = tracking_series[company_series.eq(TARGET_COMPANY) & tracking_series.ne("")]
    return pd.DataFrame({TRACKING_HEADER: tracking_series.tolist()})


def build_order_lookup(csv_texts: list[str]) -> dict[str, str]:
    lookup = {}

    for csv_text in csv_texts:
        df = pd.read_csv(StringIO(csv_text), dtype=str, keep_default_na=False).fillna("")
        required_col_index = max(ORDER_COL_INDEX, RETURN_TRACKING_COL_INDEX)
        if df.shape[1] <= required_col_index:
            raise DataValidationError(f"반품 참고 시트에 J열이 없습니다. 실제 열 수: {df.shape[1]}")

        order_header = clean_text(df.columns[ORDER_COL_INDEX])
        tracking_header = clean_text(df.columns[RETURN_TRACKING_COL_INDEX])
        if order_header not in ORDER_HEADER_ALIASES:
            raise DataValidationError(
                f"반품 참고 시트 A열 머리글이 '{ORDER_HEADER}' 또는 '주문/보관번호'가 아닙니다: '{order_header}'"
            )
        if tracking_header not in RETURN_TRACKING_HEADER_ALIASES:
            raise DataValidationError(
                f"반품 참고 시트 J열 머리글이 '{RETURN_TRACKING_HEADER}'이 아닙니다: '{tracking_header}'"
            )

        order_series = df.iloc[:, ORDER_COL_INDEX].apply(clean_text)
        tracking_series = df.iloc[:, RETURN_TRACKING_COL_INDEX].apply(clean_text)
        for order_no, tracking_no in zip(order_series, tracking_series):
            if not order_no or not tracking_no:
                continue
            lookup.setdefault(tracking_no, order_no)

    return lookup


def add_order_numbers(output_df: pd.DataFrame, order_lookup: dict[str, str]) -> pd.DataFrame:
    result_df = output_df.copy()
    order_numbers = result_df[TRACKING_HEADER].apply(lambda tracking_no: order_lookup.get(clean_text(tracking_no), ""))
    result_df.insert(0, ORDER_HEADER, order_numbers)
    return result_df


def build_platform_detail_df(order_tracking_df: pd.DataFrame, platform_csv_text: str) -> pd.DataFrame:
    platform_df = pd.read_csv(StringIO(platform_csv_text), dtype=str, keep_default_na=False).fillna("")
    platform_df.columns = [clean_text(column) for column in platform_df.columns]

    missing_columns = [column for column in PLATFORM_REQUIRED_COLUMNS if column not in platform_df.columns]
    if missing_columns:
        raise DataValidationError(f"플랫폼 출고 내역 시트 필수 열이 없습니다: {missing_columns}")

    source_df = platform_df[PLATFORM_REQUIRED_COLUMNS].copy()
    source_df = source_df.apply(lambda col: col.map(clean_text))
    source_df = source_df[source_df["주문번호"].ne("")].copy()

    tracking_lookup = {}
    for order_no, tracking_no in zip(order_tracking_df[ORDER_HEADER], order_tracking_df[TRACKING_HEADER]):
        order_no = clean_text(order_no)
        tracking_no = clean_text(tracking_no)
        if not order_no or not tracking_no:
            continue
        tracking_lookup.setdefault(order_no, tracking_no)

    target_orders = set(tracking_lookup)
    matched_df = source_df[source_df["주문번호"].isin(target_orders)].copy()

    rows = []
    for _, row in matched_df.iterrows():
        output_row = {column: "" for column in OUTPUT_COLUMNS}
        output_row["날짜"] = today_str()
        output_row["원거래 날짜"] = row["날짜"]
        for column in [
            "플랫폼",
            "주문번호",
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
        ]:
            output_row[column] = row[column]
        output_row["운송장"] = tracking_lookup.get(clean_text(row["주문번호"]), "")
        output_row["내역"] = "반품"
        rows.append(output_row)

    matched_orders = {clean_text(order_no) for order_no in matched_df["주문번호"].tolist()}
    for _, row in order_tracking_df.iterrows():
        order_no = clean_text(row.get(ORDER_HEADER, ""))
        tracking_no = clean_text(row.get(TRACKING_HEADER, ""))
        if order_no in matched_orders or (not order_no and not tracking_no):
            continue
        output_row = {column: "" for column in OUTPUT_COLUMNS}
        output_row["날짜"] = today_str()
        output_row["주문번호"] = order_no
        output_row["운송장"] = tracking_no
        output_row["내역"] = "반품"
        rows.append(output_row)

    log(f"플랫폼 출고 내역 매칭 {len(matched_df)}행 / 상세 미매칭 {len(rows) - len(matched_df)}행")
    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS).fillna("").astype(str).reset_index(drop=True)


def upload_to_google_sheet(session: requests.Session, df: pd.DataFrame):
    payload = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": TARGET_SHEET_NAME,
        "values": [df.columns.tolist()] + df.fillna("").astype(str).values.tolist(),
        "clear": True,
        "append": False,
    }

    last_error = None
    for attempt in range(1, RETRIES + 1):
        try:
            response = session.post(WEB_APP_URL, json=payload, timeout=UPLOAD_TIMEOUT)
            response.raise_for_status()
            log(f"업로드 완료 -> {TARGET_SHEET_NAME} / {len(df)}행")
            log(f"응답: {response.text}")
            return
        except Exception as exc:
            last_error = exc
            log(f"업로드 실패 ({attempt}/{RETRIES}) -> {exc}")
            if attempt < RETRIES:
                time.sleep(RETRY_SLEEP_SEC)
    raise last_error


def main():
    session = make_session()
    log("출고 시트 다운로드 시작")
    csv_text = fetch_csv_text(session, SOURCE_CSV_URL)
    output_df = build_output_df(csv_text)
    log("반품 참고 시트 다운로드 시작")
    cancel_csv_text = fetch_csv_text(session, CANCEL_CSV_URL)
    buyer_return_csv_text = fetch_csv_text(session, BUYER_RETURN_CSV_URL)
    return_progress_csv_text = fetch_csv_text(session, RETURN_PROGRESS_CSV_URL)
    order_lookup = build_order_lookup([cancel_csv_text, buyer_return_csv_text, return_progress_csv_text])
    output_df = add_order_numbers(output_df, order_lookup)
    log("플랫폼 출고 내역 다운로드 시작")
    platform_csv_text = fetch_csv_text(session, PLATFORM_SOURCE_CSV_URL)
    output_df = build_platform_detail_df(output_df, platform_csv_text)
    log(f"업로드 데이터 생성 완료 -> {len(output_df)}행")
    upload_to_google_sheet(session, output_df)


if __name__ == "__main__":
    main()
