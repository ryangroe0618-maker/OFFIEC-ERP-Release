# -*- coding: utf-8 -*-

from collections import defaultdict
import csv
from io import StringIO
import time

import pandas as pd
import requests


ACTUAL_SHIPMENT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRzG4kGawePQdGxZB9CII2zfEJKV4Vgdp4Ux3MmiXgr9KFHSX00xdOPFQZ_YyxO46lc0Jq-lcA8AuS5/pub?gid=798336653&single=true&output=csv"
OUT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRzG4kGawePQdGxZB9CII2zfEJKV4Vgdp4Ux3MmiXgr9KFHSX00xdOPFQZ_YyxO46lc0Jq-lcA8AuS5/pub?gid=594145141&single=true&output=csv"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbytRA2OXU2jNt0llG8w2y0voM2q6HV8Xm1nCstfBQ4b3qMMSjBWLnDobr8aalYErcXdOw/exec"
SPREADSHEET_ID = "1cmY9wL9zP86mdP_BiwIqbjigHFdPemPq0sV8ulL_XMs"
TARGET_SHEET_NAME = "플랫폼"
FULLMAKE_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwkEjqAetHvOLxoZgWlxL0EjWPbuFL1wGCCd5Pv_Gi5_g5HV5Yjjz7hojXr61ghRRrg-Q/exec"
FULLMAKE_SPREADSHEET_ID = "1MCYrN0nsrXk5_VbXmp4XQMBPJujDwBQCAMpvALIRa6g"
FULLMAKE_SHEET_NAME = "풀메이커"

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
    "내역",
    "운송장",
]
FULLMAKE_OUTPUT_COLUMNS = [
    "날짜",
    "브랜드",
    "코드",
    "품번",
    "사이즈",
    "수량",
    "매장명",
    "할인가",
    "총 수입",
    "환율",
    "총 공급가",
    "마진",
    "내역",
]

DOWNLOAD_RETRIES = 3
DOWNLOAD_TIMEOUT = (10, 60)
DOWNLOAD_SLEEP_SEC = 2
UPLOAD_RETRIES = 3
UPLOAD_TIMEOUT = (10, 300)
UPLOAD_SLEEP_SEC = 3


class DataValidationError(ValueError):
    pass


def log(message: str):
    print(f"[FULLMAKE 마진 수정본] {message}", flush=True)


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    return session


def fetch_csv_text(session: requests.Session, url: str) -> str:
    last_error = None
    for attempt in range(1, DOWNLOAD_RETRIES + 1):
        try:
            response = session.get(url, timeout=DOWNLOAD_TIMEOUT)
            response.raise_for_status()
            response.encoding = "utf-8"
            return response.text
        except Exception as exc:
            last_error = exc
            log(f"CSV 다운로드 실패 ({attempt}/{DOWNLOAD_RETRIES}) -> {exc}")
            if attempt < DOWNLOAD_RETRIES:
                time.sleep(DOWNLOAD_SLEEP_SEC)
    raise last_error


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace('"', "").replace("\t", "").replace("\r", "").replace("\n", "")
    return " ".join(text.split())


def clean_key(value) -> str:
    return clean_text(value).replace("-", "").upper()


def to_number(value, default=0.0) -> float:
    text = clean_text(value).replace(",", "").replace("%", "")
    if text == "":
        return float(default)
    number = pd.to_numeric(text, errors="coerce")
    if pd.isna(number):
        return float(default)
    return float(number)


def display_number(value) -> str:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number) or float(number) == 0:
        return ""
    number = float(number)
    if number.is_integer():
        return str(int(number))
    return f"{number:.4f}".rstrip("0").rstrip(".")


def row_is_blank(row: pd.Series) -> bool:
    return "".join(row.astype(str).map(clean_text)).strip() == ""


def read_actual_shipment(csv_text: str) -> pd.DataFrame:
    required = ["날짜", "브랜드", "품번", "사이즈", "주문 수량", "출고 수량", "부족 수량"]
    try:
        legacy_df = pd.read_csv(
            StringIO(csv_text),
            dtype=str,
            header=1,
            keep_default_na=False,
            skip_blank_lines=False,
        ).fillna("")
        if all(col in legacy_df.columns for col in required):
            legacy_df = legacy_df.loc[~legacy_df.apply(row_is_blank, axis=1)].reset_index(drop=True)
            for col in legacy_df.columns:
                legacy_df[col] = legacy_df[col].map(clean_text)
            return legacy_df[required]
    except Exception:
        pass

    raw_rows = [row for row in csv.reader(StringIO(csv_text))]
    max_cols = max((len(row) for row in raw_rows), default=0)
    raw_rows = [row + [""] * (max_cols - len(row)) for row in raw_rows]
    raw_df = pd.DataFrame(raw_rows, dtype=str).fillna("")
    for col in raw_df.columns:
        raw_df[col] = raw_df[col].map(clean_text)

    header_positions = []
    for row_idx in range(raw_df.shape[0]):
        for col_idx in range(raw_df.shape[1] - 4):
            header_values = [clean_text(raw_df.iat[row_idx, col_idx + offset]) for offset in range(5)]
            if header_values == ["품번", "컬러", "사이즈", "수량", "매장명"]:
                header_positions.append((row_idx, col_idx))

    if not header_positions:
        raise DataValidationError("FULLMAKE 실출고 시트에서 가로 취합 머리글(품번/컬러/사이즈/수량/매장명)을 찾을 수 없습니다.")

    rows = []
    for header_row_idx, start_col_idx in header_positions:
        brand_name = ""
        if header_row_idx >= 2:
            brand_name = clean_text(raw_df.iat[header_row_idx - 2, start_col_idx])
        if not brand_name:
            brand_name = "브랜드 없음"

        for row_idx in range(header_row_idx + 1, raw_df.shape[0]):
            item_no = clean_text(raw_df.iat[row_idx, start_col_idx])
            color = clean_text(raw_df.iat[row_idx, start_col_idx + 1])
            size = clean_text(raw_df.iat[row_idx, start_col_idx + 2])
            qty = clean_text(raw_df.iat[row_idx, start_col_idx + 3])
            store_name = clean_text(raw_df.iat[row_idx, start_col_idx + 4])
            if item_no == "" and color == "" and size == "" and qty == "" and store_name == "":
                continue
            if item_no == "":
                continue

            is_no_stock = brand_name == "재고 없음" or store_name.replace(" ", "") == "재고없음"
            rows.append(
                {
                    "날짜": "",
                    "브랜드": "" if is_no_stock else brand_name,
                    "품번": item_no,
                    "사이즈": size,
                    "수량": qty,
                    "매장명": "재고없음" if is_no_stock else store_name,
                    "내역": "재고 없음" if is_no_stock else "출고",
                }
            )

    if not rows:
        raise DataValidationError("FULLMAKE 실출고 시트에 가로 취합 데이터가 없습니다.")
    return pd.DataFrame(rows)


def read_out_sheet(csv_text: str) -> pd.DataFrame:
    df = pd.read_csv(StringIO(csv_text), dtype=str, header=0, keep_default_na=False).fillna("")
    required = [
        "날짜",
        "플랫폼",
        "브랜드",
        "코드",
        "품번",
        "사이즈",
        "수량",
        "매장명",
        "최초가",
        "공급율",
        "판매가",
        "공급가",
    ]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise DataValidationError(f"출고 시트 필수 열이 없습니다: {missing}")

    df = df.loc[~df.apply(row_is_blank, axis=1)].reset_index(drop=True)
    for col in df.columns:
        df[col] = df[col].map(clean_text)
    return df[required]


def make_compare_key(date_value, item_no, size) -> tuple[str, str, str]:
    return (clean_text(date_value), clean_key(item_no), clean_text(size).upper())


def make_item_size_key(item_no, size) -> tuple[str, str]:
    return (clean_key(item_no), clean_text(size).upper())


def build_out_lookup(out_df: pd.DataFrame) -> dict[tuple[str, str, str], list[dict]]:
    lookup = defaultdict(list)
    for _, row in out_df.iterrows():
        qty = to_number(row["수량"])
        total_sale = to_number(row["판매가"])
        total_supply = to_number(row["공급가"])
        unit_sale = total_sale / qty if qty else total_sale
        unit_supply = total_supply / qty if qty else total_supply
        entry = row.to_dict()
        entry["등록 수량_num"] = qty
        entry["판매 단가_num"] = unit_sale
        entry["공급 단가_num"] = unit_supply
        lookup[make_item_size_key(row["품번"], row["사이즈"])].append(entry)
    return lookup


def compare_result(order_qty: float, actual_qty: float, registered_qty: float) -> str:
    if actual_qty == registered_qty:
        return "일치"
    if actual_qty == 0:
        return "재고 없음"
    if actual_qty < registered_qty:
        return "부분 출고"
    return "초과 출고"


def make_output_rows(actual_df: pd.DataFrame, out_df: pd.DataFrame) -> list[dict]:
    out_lookup = build_out_lookup(out_df)
    rows = []

    def make_standard_row(
        actual_row: pd.Series,
        entry: dict,
        qty,
        unit_sale,
        unit_supply,
        total_sale,
        total_supply,
        status: str,
    ) -> dict:
        return {
            "날짜": clean_text(actual_row.get("날짜", "")) or entry.get("날짜", ""),
            "플랫폼": entry.get("플랫폼", "FULLMAKE"),
            "주문번호": "",
            "뒤 4자리": "",
            "브랜드": entry.get("브랜드", "") or actual_row["브랜드"],
            "코드": entry.get("코드", ""),
            "품번": entry.get("품번", actual_row["품번"]),
            "사이즈": entry.get("사이즈", actual_row["사이즈"]),
            "수량": display_number(qty),
            "매장명": clean_text(actual_row.get("매장명", "")) or entry.get("매장명", ""),
            "할인가": display_number(unit_sale),
            "총 판매가": display_number(total_sale),
            "총 수수료": "",
            "총 수입": display_number(total_sale),
            "환율": entry.get("공급율", ""),
            "총 KRW": display_number(total_sale),
            "총 공급가": display_number(total_supply),
            "마진": display_number(total_sale - total_supply),
            "내역": status,
            "운송장": "",
        }

    for _, actual_row in actual_df.iterrows():
        key = make_item_size_key(actual_row["품번"], actual_row["사이즈"])
        out_entries = out_lookup.get(key, [])
        if "수량" in actual_df.columns:
            actual_qty = to_number(actual_row["수량"])
            status = clean_text(actual_row.get("내역", "")) or "출고"
            is_no_stock = status == "재고 없음" or clean_text(actual_row.get("매장명", "")).replace(" ", "") == "재고없음"

            if is_no_stock:
                base_entry = out_entries[0] if out_entries else {}
                unit_sale = base_entry.get("판매 단가_num", 0)
                unit_supply = base_entry.get("공급 단가_num", 0)
                total_sale = round(unit_sale * actual_qty, 0)
                total_supply = round(unit_supply * actual_qty, 0)
                no_stock_row = make_standard_row(
                    actual_row,
                    base_entry,
                    actual_qty,
                    unit_sale,
                    unit_supply,
                    total_sale,
                    total_supply,
                    "재고 없음",
                )
                no_stock_row["매장명"] = "재고없음"
                rows.append(no_stock_row)
                continue

            remaining_qty = actual_qty
            if out_entries and actual_qty > 0:
                for entry in out_entries:
                    if remaining_qty <= 0:
                        break
                    use_qty = min(remaining_qty, entry["등록 수량_num"])
                    if use_qty <= 0:
                        continue
                    total_sale = round(entry["판매 단가_num"] * use_qty, 0)
                    total_supply = round(entry["공급 단가_num"] * use_qty, 0)
                    rows.append(
                        make_standard_row(
                            actual_row,
                            entry,
                            use_qty,
                            entry["판매 단가_num"],
                            entry["공급 단가_num"],
                            total_sale,
                            total_supply,
                            "출고",
                        )
                    )
                    remaining_qty -= use_qty

            if remaining_qty > 0:
                base_entry = out_entries[0] if out_entries else {}
                unit_sale = base_entry.get("판매 단가_num", 0)
                unit_supply = base_entry.get("공급 단가_num", 0)
                total_sale = round(unit_sale * remaining_qty, 0)
                total_supply = round(unit_supply * remaining_qty, 0)
                rows.append(
                    make_standard_row(
                        actual_row,
                        base_entry,
                        remaining_qty,
                        unit_sale,
                        unit_supply,
                        total_sale,
                        total_supply,
                        "초과 출고" if out_entries else "출고 시트 없음",
                    )
                )
            continue

        order_qty = to_number(actual_row["주문 수량"])
        actual_qty = to_number(actual_row["출고 수량"])
        shortage_qty = to_number(actual_row["부족 수량"])
        registered_qty = sum(entry["등록 수량_num"] for entry in out_entries)
        result = compare_result(order_qty, actual_qty, registered_qty)
        shortage_abs = abs(shortage_qty) if shortage_qty < 0 else shortage_qty
        missing_qty = max(order_qty - actual_qty, shortage_abs, 0)

        remaining_qty = actual_qty
        if out_entries and actual_qty > 0:
            for entry in out_entries:
                if remaining_qty <= 0:
                    break
                use_qty = min(remaining_qty, entry["등록 수량_num"])
                if use_qty <= 0:
                    continue
                total_sale = round(entry["판매 단가_num"] * use_qty, 0)
                total_supply = round(entry["공급 단가_num"] * use_qty, 0)
                rows.append(
                    make_standard_row(
                        actual_row,
                        entry,
                        use_qty,
                        entry["판매 단가_num"],
                        entry["공급 단가_num"],
                        total_sale,
                        total_supply,
                        "출고",
                    )
                )
                remaining_qty -= use_qty

        if remaining_qty > 0:
            base_entry = out_entries[0] if out_entries else {}
            unit_sale = base_entry.get("판매 단가_num", 0)
            unit_supply = base_entry.get("공급 단가_num", 0)
            total_sale = round(unit_sale * remaining_qty, 0)
            total_supply = round(unit_supply * remaining_qty, 0)
            rows.append(
                make_standard_row(
                    actual_row,
                    base_entry,
                    remaining_qty,
                    unit_sale,
                    unit_supply,
                    total_sale,
                    total_supply,
                    "초과 출고" if out_entries else "출고 시트 없음",
                )
            )

        if missing_qty > 0:
            base_entry = out_entries[0] if out_entries else {}
            unit_sale = base_entry.get("판매 단가_num", 0)
            unit_supply = base_entry.get("공급 단가_num", 0)
            total_sale = round(unit_sale * missing_qty, 0)
            total_supply = round(unit_supply * missing_qty, 0)
            missing_row = make_standard_row(
                actual_row,
                base_entry,
                missing_qty,
                unit_sale,
                unit_supply,
                total_sale,
                total_supply,
                "재고 없음",
            )
            missing_row["매장명"] = "재고없음"
            rows.append(missing_row)

    return rows


def build_margin_df(actual_df: pd.DataFrame, out_df: pd.DataFrame) -> pd.DataFrame:
    rows = make_output_rows(actual_df, out_df)
    if not rows:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    result_df = pd.DataFrame(rows)
    for col in OUTPUT_COLUMNS:
        if col not in result_df.columns:
            result_df[col] = ""
    result_df["__store_sort__"] = result_df["매장명"].map(
        lambda value: (2, "") if clean_text(value) == "재고없음" else ((1, "") if clean_text(value) == "" else (0, clean_text(value)))
    )
    result_df = result_df.sort_values(
        by=["__store_sort__", "날짜", "브랜드", "품번", "사이즈"],
        ascending=[True, True, True, True, True],
        kind="stable",
    ).drop(columns=["__store_sort__"]).reset_index(drop=True)
    return result_df[OUTPUT_COLUMNS]


def build_fullmake_upload_df(df: pd.DataFrame) -> pd.DataFrame:
    filtered_df = df.loc[df["내역"].map(clean_text).eq("출고")].copy()
    for col in FULLMAKE_OUTPUT_COLUMNS:
        if col not in filtered_df.columns:
            filtered_df[col] = ""
    return filtered_df[FULLMAKE_OUTPUT_COLUMNS]


def upload_to_google_sheet(
    df: pd.DataFrame,
    web_app_url: str,
    spreadsheet_id: str,
    sheet_name: str,
    columns: list[str],
):
    if df.empty:
        log(f"업로드 대상 없음 -> {sheet_name}")
        return

    upload_df = df.copy()
    for col in columns:
        if col not in upload_df.columns:
            upload_df[col] = ""
    upload_df = upload_df[columns]

    values = [columns] + upload_df.fillna("").astype(str).values.tolist()
    payload = {
        "spreadsheetId": spreadsheet_id,
        "sheetName": sheet_name,
        "values": values,
        "append": True,
        "clear": False,
    }

    last_error = None
    log(f"업로드 시작 -> {sheet_name} ({len(upload_df)}행)")
    for attempt in range(1, UPLOAD_RETRIES + 1):
        try:
            response = requests.post(
                web_app_url,
                json=payload,
                headers={"User-Agent": "Mozilla/5.0"},
                timeout=UPLOAD_TIMEOUT,
            )
            response.raise_for_status()
            response_text = (response.text or "").strip()
            try:
                response_json = response.json()
            except Exception:
                response_json = None
            if response_text.startswith("<!DOCTYPE html") or "Google Apps Script" in response_text:
                raise RuntimeError(f"Apps Script 오류 페이지를 반환했습니다. 응답 일부: {response_text[:200]}")
            if not isinstance(response_json, dict):
                raise RuntimeError(f"Apps Script JSON 응답이 아닙니다. 응답 일부: {response_text[:200]}")
            if isinstance(response_json, dict) and response_json.get("ok") is False:
                raise RuntimeError(response_json.get("error") or response_json.get("message") or response_text)
            log(f"업로드 완료 -> {sheet_name}")
            log(f"응답: {response_text}")
            return
        except Exception as exc:
            last_error = exc
            log(f"업로드 실패 ({sheet_name}, {attempt}/{UPLOAD_RETRIES}) -> {exc}")
            if attempt < UPLOAD_RETRIES:
                time.sleep(UPLOAD_SLEEP_SEC)
    raise last_error


def main():
    start_time = time.perf_counter()
    session = make_session()
    log("FULLMAKE 실출고 시트 다운로드 시작")
    actual_df = read_actual_shipment(fetch_csv_text(session, ACTUAL_SHIPMENT_CSV_URL))
    log(f"FULLMAKE 실출고 로드 완료 -> {len(actual_df)}행")

    log("출고 시트 다운로드 시작")
    out_df = read_out_sheet(fetch_csv_text(session, OUT_CSV_URL))
    log(f"출고 시트 로드 완료 -> {len(out_df)}행")

    result_df = build_margin_df(actual_df, out_df)
    log(f"비교/마진 정리 완료 -> {len(result_df)}행")

    upload_to_google_sheet(
        result_df,
        web_app_url=WEB_APP_URL,
        spreadsheet_id=SPREADSHEET_ID,
        sheet_name=TARGET_SHEET_NAME,
        columns=OUTPUT_COLUMNS,
    )

    fullmake_df = build_fullmake_upload_df(result_df)
    upload_to_google_sheet(
        fullmake_df,
        web_app_url=FULLMAKE_WEB_APP_URL,
        spreadsheet_id=FULLMAKE_SPREADSHEET_ID,
        sheet_name=FULLMAKE_SHEET_NAME,
        columns=FULLMAKE_OUTPUT_COLUMNS,
    )

    elapsed = time.perf_counter() - start_time
    log(f"완료 ({elapsed:.2f}초)")


if __name__ == "__main__":
    main()
