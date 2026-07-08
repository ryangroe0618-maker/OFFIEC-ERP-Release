# -*- coding: utf-8 -*-
import re
import time
from io import StringIO

import pandas as pd
import requests


# =========================
# 입력 구글 시트 CSV
# =========================
INPUT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC-ZAQixoxbta3FsLfG_f8qugvCGbjK2yrsKcLA6tJsi1ww5qDl-_21od9-55Lv0F9dmzlR5c1QIh/pub?gid=0&single=true&output=csv"

# =========================
# 구글 시트 업로드 설정
# =========================
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyKiv4r8GjUPIsX5Tz7nVlpeTw7WIqPHsh8nvVw9WqKPihQSIKk4Vg4QMaO8VfkS0aN/exec"
SPREADSHEET_ID = "1aK2IZzdfsEx8YBd0G4oSUg3GPUmDJ0nas9AUAVTNopE"
TARGET_SHEET_NAME = "현재고 변환"
START_TIME = time.time()


# =========================
# 공통 함수
# =========================
def log(message: str):
    elapsed = time.time() - START_TIME
    print(f"[{elapsed:6.1f}s] {message}", flush=True)


def read_google_sheet_csv(url: str) -> pd.DataFrame:
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, timeout=(10, 60), headers=headers)
    r.raise_for_status()
    r.encoding = "utf-8"
    return pd.read_csv(StringIO(r.text), dtype=str).fillna("")


def upload_to_google_sheet(df: pd.DataFrame, web_app_url: str, spreadsheet_id: str, sheet_name: str):
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()

    payload = {
        "spreadsheetId": spreadsheet_id,
        "sheetName": sheet_name,
        "values": values,
    }

    log(f"구글 시트 업로드 시작: {sheet_name} / {len(df)}행")
    r = requests.post(web_app_url, json=payload, timeout=(10, 120))
    r.raise_for_status()
    log("구글 시트 업로드 완료")
    log(f"응답: {r.text}")


def clean_text(value):
    if pd.isna(value):
        return ""
    s = str(value).strip()
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("\\", "/")
    s = re.sub(r"\s+", "", s)
    return s


def clean_item_no(value):
    if pd.isna(value):
        return ""
    s = str(value).strip()
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("\\", "/")
    s = s.replace("-", "")
    s = re.sub(r"\s+", "", s)
    return s


def to_pretty_value(value):
    if value is None or pd.isna(value):
        return ""

    s = str(value).strip()
    if s == "":
        return ""

    try:
        num = float(s)
        if num.is_integer():
            return str(int(num))
        return str(num)
    except Exception:
        return s


# =========================
# 할인가 계산
# 할인율 없으면 최초가 그대로
# =========================
def build_sale_price(df):
    temp = df.copy()

    temp["최초가_num"] = pd.to_numeric(
        temp["최초가"].astype(str).str.replace(",", "", regex=False),
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

    temp["할인율_ratio"] = temp["할인율_num"].where(
        temp["할인율_num"].isna() | (temp["할인율_num"] <= 1),
        temp["할인율_num"] / 100,
    )

    sale_price = temp["최초가_num"].where(
        temp["할인율_ratio"].isna(),
        temp["최초가_num"] * (1 - temp["할인율_ratio"]),
    ).round()

    def _fmt(x):
        if pd.isna(x):
            return ""
        return str(int(x))

    return sale_price.apply(_fmt)


# =========================
# 사이즈 변환
# =========================
def size_conv_1(value):
    s = clean_text(value)
    if s == "":
        return ""

    m = re.search(r"\(([^)]*)\)", s)
    if m:
        return to_pretty_value(m.group(1))

    return to_pretty_value(s)


def size_conv_2(value):
    s = clean_text(value)
    if s == "":
        return ""

    if "(" in s:
        left = s.split("(", 1)[0]
    else:
        left = s

    return to_pretty_value(left)


def size_conv_3(value):
    s = clean_text(value)
    if s == "":
        return ""

    digits = re.sub(r"[^\d]", "", s)
    if digits == "":
        return ""

    return to_pretty_value(digits)


def make_code(item_no, size_value):
    item_no = "" if pd.isna(item_no) else str(item_no).strip()
    size_value = "" if pd.isna(size_value) else str(size_value).strip()

    if item_no == "" or size_value == "":
        return ""

    return f"{item_no}{size_value}"


def make_code_series(item_no_series: pd.Series, size_series: pd.Series) -> pd.Series:
    item = item_no_series.fillna("").astype(str).str.strip()
    size = size_series.fillna("").astype(str).str.strip()
    return (item + size).where((item != "") & (size != ""), "")


# =========================
# 열 순서 정리
# =========================
def reorder_columns(df):
    ordered_cols = [
        "브랜드",
        "코드",
        "바코드",
        "품번",
        "품번_변환",
        "상품명",
        "컬러",
        "사이즈",
        "사이즈_원본정리",
        "변환사이즈1",
        "변환사이즈2",
        "변환사이즈3",
        "변환코드1",
        "변환코드2",
        "변환코드3",
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
        "사무실 - 스퀘어원",
        "사무실 - 반품",
    ]

    existing_cols = [c for c in ordered_cols if c in df.columns]
    other_cols = [c for c in df.columns if c not in existing_cols]

    return df[existing_cols + other_cols].copy()


# =========================
# 실행
# =========================
def main():
    log("구글 시트 파일 읽는 중...")
    df = read_google_sheet_csv(INPUT_CSV_URL)

    if "품번" not in df.columns:
        print("오류: '품번' 열이 없습니다.")
        return

    if "사이즈" not in df.columns:
        print("오류: '사이즈' 열이 없습니다.")
        return

    # 없을 수 있는 열 보정
    optional_cols = [
        "브랜드", "코드", "바코드", "상품명", "컬러", "EU",
        "최초가", "할인율", "공급가", "현재고",
        "사무실", "스퀘어원", "구월", "부천", "휠라 파주", "푸마 여주", "사무실 - 스퀘어원", "사무실 - 반품"
    ]
    for col in optional_cols:
        if col not in df.columns:
            df[col] = ""

    # 원본 정리
    log("원본 정리 시작")
    df["품번_변환"] = df["품번"].apply(clean_item_no)
    df["사이즈_원본정리"] = df["사이즈"].apply(clean_text)

    # 변환 사이즈
    df["변환사이즈1"] = df["사이즈"].apply(size_conv_1)
    df["변환사이즈2"] = df["사이즈"].apply(size_conv_2)
    df["변환사이즈3"] = df["사이즈"].apply(size_conv_3)

    # 변환 코드
    df["변환코드1"] = make_code_series(df["품번"], df["변환사이즈1"])
    df["변환코드2"] = make_code_series(df["품번"], df["변환사이즈2"])
    df["변환코드3"] = make_code_series(df["품번"], df["변환사이즈3"])

    # 할인가
    log("할인가 계산 시작")
    df["할인가"] = build_sale_price(df)

    # 열 정리
    log("열 정리 시작")
    df = reorder_columns(df)

    # 바로 업로드
    upload_to_google_sheet(
        df,
        web_app_url=WEB_APP_URL,
        spreadsheet_id=SPREADSHEET_ID,
        sheet_name=TARGET_SHEET_NAME,
    )

    log("완료")


if __name__ == "__main__":
    main()
