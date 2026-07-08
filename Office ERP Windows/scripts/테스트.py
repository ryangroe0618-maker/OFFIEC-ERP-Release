# -*- coding: utf-8 -*-

from pathlib import Path
import json

import gspread
from google.auth.exceptions import GoogleAuthError


SOURCE_SPREADSHEET_ID = "1aK2IZzdfsEx8YBd0G4oSUg3GPUmDJ0nas9AUAVTNopE"
SOURCE_SHEET_NAME = "현재고 변환"

TARGET_SPREADSHEET_ID = "1k5Y-vlq4oEwBEX1q0C3WHYivyfaB99Cneo2r6588hBE"
TARGET_SHEET_NAME = "테스트"

BASE_DIR = Path(__file__).resolve().parent
SERVICE_ACCOUNT_FILE = BASE_DIR / "service_account.json"
OAUTH_CREDENTIALS_FILE = BASE_DIR / "credentials.json"
OAUTH_TOKEN_FILE = BASE_DIR / "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def log(message: str) -> None:
    print(message, flush=True)


def find_service_account_file() -> Path | None:
    if SERVICE_ACCOUNT_FILE.exists():
        return SERVICE_ACCOUNT_FILE

    for json_file in sorted(BASE_DIR.glob("*.json")):
        if json_file in {OAUTH_CREDENTIALS_FILE, OAUTH_TOKEN_FILE}:
            continue
        try:
            with json_file.open(encoding="utf-8") as file:
                data = json.load(file)
        except (OSError, json.JSONDecodeError):
            continue

        if data.get("type") == "service_account" and data.get("private_key"):
            return json_file

    return None


def create_google_client() -> gspread.Client:
    service_account_file = find_service_account_file()
    if service_account_file:
        log(
            "서비스 계정으로 Google Sheets API에 연결합니다: "
            f"{service_account_file.name}"
        )
        return gspread.service_account(
            filename=str(service_account_file),
            scopes=SCOPES,
        )

    if OAUTH_CREDENTIALS_FILE.exists():
        log("OAuth 사용자 계정으로 Google Sheets API에 연결합니다.")
        return gspread.oauth(
            scopes=SCOPES,
            credentials_filename=str(OAUTH_CREDENTIALS_FILE),
            authorized_user_filename=str(OAUTH_TOKEN_FILE),
        )

    raise FileNotFoundError(
        "Google API 인증 파일이 없습니다.\n"
        f"- 서비스 계정 방식: {SERVICE_ACCOUNT_FILE.name}\n"
        f"- OAuth 방식: {OAUTH_CREDENTIALS_FILE.name}\n"
        "둘 중 하나를 테스트.py와 같은 폴더에 넣어 주세요."
    )


def read_source_values(client: gspread.Client) -> list[list[str]]:
    log(f"원본 시트 읽는 중: {SOURCE_SHEET_NAME}")
    source_sheet = client.open_by_key(SOURCE_SPREADSHEET_ID).worksheet(
        SOURCE_SHEET_NAME
    )
    values = source_sheet.get_all_values()

    if not values:
        raise ValueError("원본 OUT 시트에 업로드할 데이터가 없습니다.")

    column_count = max(len(row) for row in values)
    log(f"원본 읽기 완료: {len(values)}행, {column_count}열")
    return values


def upload_target_values(
    client: gspread.Client,
    values: list[list[str]],
) -> None:
    log(f"대상 시트 연결 중: {TARGET_SHEET_NAME}")
    target_sheet = client.open_by_key(TARGET_SPREADSHEET_ID).worksheet(
        TARGET_SHEET_NAME
    )

    log("대상 시트의 기존 값을 지우는 중")
    target_sheet.clear()

    log(f"대상 시트에 {len(values)}행 업로드 중")
    target_sheet.update(
        values=values,
        range_name="A1",
        value_input_option="RAW",
    )
    log("업로드 완료")


def main() -> None:
    try:
        client = create_google_client()
        values = read_source_values(client)
        upload_target_values(client, values)
    except gspread.WorksheetNotFound as exc:
        raise RuntimeError("지정한 시트 탭을 찾을 수 없습니다.") from exc
    except gspread.SpreadsheetNotFound as exc:
        raise RuntimeError(
            "스프레드시트를 찾을 수 없거나 인증 계정에 접근 권한이 없습니다."
        ) from exc
    except GoogleAuthError as exc:
        raise RuntimeError(f"Google API 인증에 실패했습니다: {exc}") from exc


if __name__ == "__main__":
    main()
