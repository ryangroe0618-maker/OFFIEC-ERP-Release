#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import csv
import io
import json
import sys
import time
import subprocess
import runpy
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QProcess, QProcessEnvironment, QTimer, Qt
from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)


def configure_script_output_encoding():
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is None or not hasattr(stream, "reconfigure"):
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


if len(sys.argv) >= 3 and sys.argv[1] == "--run-script":
    configure_script_output_encoding()
    script_path = Path(sys.argv[2]).resolve()
    sys.path.insert(0, str(script_path.parent.parent))
    sys.path.insert(0, str(script_path.parent))
    runpy.run_path(str(script_path), run_name="__main__")
    sys.exit(0)


if getattr(sys, "frozen", False):
    SEARCH_ROOT = Path(sys.executable).resolve().parent
else:
    SEARCH_ROOT = Path(__file__).resolve().parent


def resolve_project_root(start_dir: Path) -> Path:
    for candidate in (start_dir, *start_dir.parents):
        if (
            (candidate / ".venv" / "Scripts" / "python.exe").exists()
            or (candidate / ".venv" / "bin" / "python3").exists()
            or (candidate / ".venv" / "bin" / "python").exists()
        ):
            return candidate
    return start_dir.parent


def resolve_python_bin(project_root: Path) -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()

    candidates = [
        project_root / ".venv" / "Scripts" / "python.exe",
        project_root / ".venv" / "bin" / "python3",
        project_root / ".venv" / "bin" / "python",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0] if sys.platform.startswith("win") else candidates[1]


PROJECT_ROOT = SEARCH_ROOT if getattr(sys, "frozen", False) else resolve_project_root(SEARCH_ROOT)
BASE_DIR = SEARCH_ROOT
SCRIPT_DIR = BASE_DIR / "scripts"
PYTHON_BIN = resolve_python_bin(PROJECT_ROOT)
LAUNCHER_LOG_DIR = PROJECT_ROOT / ".launcher_logs"
OFFICE_USER_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vTw5vRa1eOkJkyOKPS-gz2rAqC4CcRqAO-qTGFh-wzQcP_p1c8SSB4V4CwqoLyOe1hZibDRsDEceomG/"
    "pub?gid=1550133707&single=true&output=csv"
)
LOGIN_SETTINGS_PATH = BASE_DIR / ".office_erp_login.json"
VERSION_FILE_PATH = BASE_DIR / ".office_erp_version"


def resolve_script_path(filename: str) -> Path:
    candidates = [
        SCRIPT_DIR / filename,
        BASE_DIR / filename,
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


@dataclass(frozen=True)
class ScriptItem:
    id: str
    title: str
    filename: str | None
    accent: str


@dataclass(frozen=True)
class PlatformItem:
    id: str
    title: str
    accent: str
    scripts: list[ScriptItem] = field(default_factory=list)


@dataclass(frozen=True)
class UserSession:
    user_id: str
    role: str

    @property
    def is_main(self) -> bool:
        return self.role.strip().lower() == "main"


def normalize_header(value: str) -> str:
    return (value or "").replace(" ", "").strip()


def load_office_users() -> list[dict[str, str]]:
    request = urllib.request.Request(
        OFFICE_USER_CSV_URL,
        headers={"User-Agent": "Office-ERP-Launcher"},
    )
    with urllib.request.urlopen(request, timeout=12) as response:
        raw_text = response.read().decode("utf-8-sig")

    reader = csv.DictReader(io.StringIO(raw_text))
    users: list[dict[str, str]] = []
    for row in reader:
        normalized = {normalize_header(key): (value or "").strip() for key, value in row.items()}
        users.append(
            {
                "id": normalized.get("아이디", ""),
                "password": (
                    normalized.get("비빌번호")
                    or normalized.get("비밀번호")
                    or normalized.get("비번")
                ),
                "role": normalized.get("구분", ""),
            }
        )
    return users


def authenticate_office_user(user_id: str, password: str) -> UserSession | None:
    wanted_id = user_id.strip()
    wanted_password = password.strip()
    if not wanted_id or not wanted_password:
        return None

    for user in load_office_users():
        if user["id"] == wanted_id and user["password"] == wanted_password:
            role = user["role"] or "Sub"
            return UserSession(user_id=wanted_id, role=role)
    return None


def build_app_structure(session: UserSession) -> dict[str, list[PlatformItem]]:
    if session.is_main:
        return APP_STRUCTURE
    return {
        section: platforms
        for section, platforms in APP_STRUCTURE.items()
        if section != "마감"
    }


def load_saved_login() -> dict[str, str]:
    if not LOGIN_SETTINGS_PATH.exists():
        return {}
    try:
        data = json.loads(LOGIN_SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(data, dict):
        return {}
    return {
        "id": str(data.get("id", "")),
        "password": str(data.get("password", "")),
    }


def save_login(user_id: str, password: str):
    LOGIN_SETTINGS_PATH.write_text(
        json.dumps(
            {
                "id": user_id,
                "password": password,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def clear_saved_login():
    try:
        LOGIN_SETTINGS_PATH.unlink(missing_ok=True)
    except OSError:
        pass


def read_installed_version() -> str:
    if not VERSION_FILE_PATH.exists():
        return "버전 확인 전"
    version = VERSION_FILE_PATH.read_text(encoding="utf-8", errors="replace").strip()
    return version or "버전 확인 전"


APP_STRUCTURE = {
    "현재고 업로드": [
        PlatformItem(
            id="stock",
            title="현재고",
            accent="#0F766E",
            scripts=[
                ScriptItem(
                    id="stock_all",
                    title="현재고 업로드",
                    filename="현재고 전체 실행.py",
                    accent="#0F766E",
                ),
            ],
        ),
    ],
    "출고 관리": [
        PlatformItem(
            id="poizon",
            title="POIZON",
            accent="#2563EB",
            scripts=[
                ScriptItem(
                    id="poizon_upload",
                    title="POIZON",
                    filename="poizon 리스트 업로드 수정중 최종.py",
                    accent="#2563EB",
                ),
                ScriptItem(
                    id="poizon_list",
                    title="POIZON 리스트 생성",
                    filename="poizon 리스트 생성.py",
                    accent="#1D4ED8",
                ),
                ScriptItem(
                    id="poizon_margin",
                    title="POIZON 마진",
                    filename="poizon 마진 정리 최종.py",
                    accent="#1E40AF",
                ),
            ],
        ),
        PlatformItem(
            id="kashion",
            title="KASHION",
            accent="#B45309",
            scripts=[
                ScriptItem(
                    id="kashion_upload",
                    title="KASHION",
                    filename="KASHION 리스트 업로드.py",
                    accent="#A16207",
                ),
                ScriptItem(
                    id="kashion_split",
                    title="KASHION 파일 구분",
                    filename="kashion 파일 구분.py",
                    accent="#B45309",
                ),
                ScriptItem(
                    id="kashion_margin",
                    title="KASHION 마진",
                    filename="kashion 정리.py",
                    accent="#C2410C",
                ),
                ScriptItem(
                    id="uno_vision",
                    title="Uno Vision",
                    filename="UNO 최종.py",
                    accent="#92400E",
                ),
            ],
        ),
        PlatformItem(
            id="buyma",
            title="BUYMA",
            accent="#DC2626",
            scripts=[
                ScriptItem(
                    id="buyma_upload",
                    title="BUYMA",
                    filename="buyma 리스트 업로드.py",
                    accent="#DC2626",
                ),
                ScriptItem(
                    id="buyma_split",
                    title="BUYMA 파일 구분",
                    filename="buyma_pdf_splitter.py",
                    accent="#B91C1C",
                ),
                ScriptItem(
                    id="buyma_margin",
                    title="BUYMA 마진",
                    filename="buyma 정리.py",
                    accent="#991B1B",
                ),
            ],
        ),
        PlatformItem(
            id="tmall",
            title="TMALL",
            accent="#7C3AED",
            scripts=[ScriptItem("tmall_pending", "TMALL", None, "#7C3AED")],
        ),
        PlatformItem(
            id="live",
            title="LIVE",
            accent="#C026D3",
            scripts=[
                ScriptItem(
                    id="live_beauty_one",
                    title="LIVE",
                    filename="beauty one.py",
                    accent="#C026D3",
                ),
                ScriptItem(
                    id="live_beauty_one_list",
                    title="LIVE 리스트 생성",
                    filename="beauty one 리스트 생성.py",
                    accent="#B91C1C",
                ),
                ScriptItem(
                    id="live_beauty_one_summary",
                    title="LIVE 정리",
                    filename="beauty one 정리.py",
                    accent="#A21CAF",
                ),
            ],
        ),
        PlatformItem(
            id="kream",
            title="KREAM",
            accent="#DC2626",
            scripts=[
                ScriptItem(
                    id="kream_upload",
                    title="KREAM",
                    filename="kream 리스트 업로드.py",
                    accent="#DC2626",
                ),
                ScriptItem(
                    id="kream_extra_upload",
                    title="KREAM 추가",
                    filename="kream 추가 리스트 업로드.py",
                    accent="#EF4444",
                ),
                ScriptItem(
                    id="kream_summary",
                    title="KREAM 정리",
                    filename="kream 정리.py",
                    accent="#991B1B",
                ),
                ScriptItem(
                    id="kream_list_create",
                    title="KREAM 리스트",
                    filename="kream 리스트 생성.py",
                    accent="#B91C1C",
                ),
                ScriptItem(
                    id="kream_shipping_register",
                    title="KREAM 출고 등록",
                    filename="kream 출고 등록.py",
                    accent="#7F1D1D",
                ),
                ScriptItem(
                    id="kream_extra_list",
                    title="KREAM 추가 리스트",
                    filename="kream 추가 리스트 업로드.py",
                    accent="#F87171",
                ),
            ],
        ),
        PlatformItem(
            id="brander",
            title="브랜더",
            accent="#EA580C",
            scripts=[
                ScriptItem(
                    id="brander_list_upload",
                    title="브랜더",
                    filename="brand 리스트 업로드.py",
                    accent="#EA580C",
                ),
                ScriptItem(
                    id="brander_list_create",
                    title="브랜더 리스트 생성",
                    filename="brand 리스트 생성.py",
                    accent="#C2410C",
                ),
                ScriptItem(
                    id="brander_shipment_collect",
                    title="브랜더 취합",
                    filename="brand 출고 취합.py",
                    accent="#9A3412",
                ),
                ScriptItem(
                    id="brander_margin_v2",
                    title="브랜더 마진",
                    filename="brand 마진 수정본.py",
                    accent="#7C2D12",
                ),
            ],
        ),
        PlatformItem(
            id="playmaker",
            title="풀메이커",
            accent="#059669",
            scripts=[
                ScriptItem(
                    id="playmaker_list_upload",
                    title="풀메이커",
                    filename="fullmake 리스트 업로드.py",
                    accent="#059669",
                ),
                ScriptItem(
                    id="playmaker_list_create",
                    title="풀메이커 리스트 생성",
                    filename="fullmake 리스트 생성.py",
                    accent="#0F766E",
                ),
                ScriptItem(
                    id="playmaker_shipment_collect",
                    title="풀메이커 취합",
                    filename="fullmake 출고 취합.py",
                    accent="#065F46",
                ),
                ScriptItem(
                    id="playmaker_margin_v2",
                    title="풀메이커마진",
                    filename="fullmake 마진 수정본.py",
                    accent="#064E3B",
                ),
            ],
        ),
        PlatformItem(
            id="japan",
            title="일본",
            accent="#0EA5E9",
            scripts=[
                ScriptItem(
                    id="japan_upload",
                    title="JAPAN",
                    filename="일본.py",
                    accent="#0EA5E9",
                ),
                ScriptItem(
                    id="japan_margin",
                    title="JAPAN 마진",
                    filename="일본 마진.py",
                    accent="#0284C7",
                ),
            ],
        ),
    ],
    "보관 판매": [
        PlatformItem(
            id="storage_poizon",
            title="POIZON 보관",
            accent="#2563EB",
            scripts=[
                ScriptItem(
                    id="storage_poizon_upload",
                    title="POIZON 보관",
                    filename="poizon 보관 정리.py",
                    accent="#2563EB",
                ),
                ScriptItem(
                    id="storage_poizon_margin",
                    title="POIZON 마진",
                    filename="poizon 보관 마진.py",
                    accent="#1E40AF",
                ),
            ],
        ),
        PlatformItem(
            id="storage_kream",
            title="KREAM",
            accent="#DC2626",
            scripts=[
                ScriptItem(
                    id="storage_kream_upload",
                    title="KREAM 보관",
                    filename="kream 보관.py",
                    accent="#DC2626",
                ),
                ScriptItem(
                    id="storage_kream_margin",
                    title="KREAM 마진",
                    filename="kream 보관 마진.py",
                    accent="#991B1B",
                ),
            ],
        ),
    ],
    "반품 관리": [
        PlatformItem(
            id="return_poizon",
            title="POIZON 반품",
            accent="#2563EB",
            scripts=[
                ScriptItem(
                    id="return_poizon_upload",
                    title="POIZON 반품",
                    filename="poizon 반품.py",
                    accent="#2563EB",
                ),
                ScriptItem(
                    id="return_poizon_list",
                    title="POIZON 반품 리스트",
                    filename="poizon 반품 펀칭 리스트.py",
                    accent="#1D4ED8",
                ),
                ScriptItem(
                    id="return_poizon_margin",
                    title="POIZON 반품 마진",
                    filename="poizon 반품 마진.py",
                    accent="#1E40AF",
                ),
            ],
        ),
        PlatformItem(
            id="return_kashion",
            title="KASHION 반품",
            accent="#B45309",
            scripts=[
                ScriptItem(
                    id="return_kashion_upload",
                    title="KASHION 반품",
                    filename="kashion 반품.py",
                    accent="#B45309",
                ),
                ScriptItem(
                    id="return_kashion_margin",
                    title="KASHION 반품 마진",
                    filename="kashion 반품 마진.py",
                    accent="#C2410C",
                ),
            ],
        ),
        PlatformItem(
            id="return_kream",
            title="KREAM 반품",
            accent="#DC2626",
            scripts=[
                ScriptItem(
                    id="return_kream_upload",
                    title="KREAM 반품",
                    filename="kream 반품.py",
                    accent="#DC2626",
                ),
                ScriptItem(
                    id="return_kream_margin",
                    title="KREAM 반품 마진",
                    filename="kream 반품 마진.py",
                    accent="#B91C1C",
                ),
            ],
        ),
    ],
    "플랫폼": [
        PlatformItem(
            id="platform",
            title="플랫폼",
            accent="#0891B2",
            scripts=[
                ScriptItem(
                    id="platform_poizon_upload",
                    title="POIZON",
                    filename="플랫폼 poizon 리스트 업로드.py",
                    accent="#0891B2",
                ),
                ScriptItem(
                    id="platform_kream_pending",
                    title="KREAM",
                    filename=None,
                    accent="#DC2626",
                ),
            ],
        ),
    ],
    "마감": [
        PlatformItem(
            id="closing_online",
            title="온라인",
            accent="#2563EB",
            scripts=[
                ScriptItem(
                    id="closing_online_daily",
                    title="일 마감",
                    filename="온라인 마감.py",
                    accent="#2563EB",
                ),
                ScriptItem(
                    id="closing_online_monthly",
                    title="월 마감",
                    filename="월마감.py",
                    accent="#1D4ED8",
                ),
            ],
        ),
        PlatformItem(
            id="closing_offline",
            title="오프라인",
            accent="#6B7280",
            scripts=[
                ScriptItem(
                    id="closing_offline_pending",
                    title="오프라인",
                    filename=None,
                    accent="#6B7280",
                ),
            ],
        ),
    ],
}


class ScriptCard(QFrame):
    def __init__(self, item: ScriptItem, run_callback):
        super().__init__()
        self.item = item
        self.run_callback = run_callback
        self.setObjectName("scriptCard")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)

        title = QPushButton(item.title)
        title.setEnabled(False)
        title.setObjectName("scriptBadge")
        title.setCursor(Qt.CursorShape.ArrowCursor)
        title.setStyleSheet(
            "QPushButton {"
            "background-color: #2563EB;"
            "color: white;"
            "border: none;"
            "border-radius: 16px;"
            "padding: 0 16px;"
            "min-height: 38px;"
            "font-size: 12px;"
            "font-weight: 700;"
            "text-align: center;"
            "}"
            "QPushButton:disabled {"
            "background-color: #2563EB;"
            "color: white;"
            "border: none;"
            "border-radius: 16px;"
            "}"
        )
        title.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)

        button = QPushButton("실행" if item.filename else "준비 중")
        button.setEnabled(bool(item.filename))
        button.clicked.connect(lambda: self.run_callback(item))
        button.setStyleSheet(
            "QPushButton {"
            "background: #1D4ED8;"
            "color: white;"
            "border: none;"
            "border-radius: 16px;"
            "padding: 12px 14px;"
            "font-weight: 800;"
            "font-size: 14px;"
            "}"
            "QPushButton:hover:!disabled {"
            "background: #2563EB;"
            "}"
            "QPushButton:disabled {"
            "background: #E5E7EB;"
            "color: #6B7280;"
            "}"
        )

        layout.addWidget(title)
        layout.addStretch(1)
        layout.addWidget(button)


class LoginDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.session: UserSession | None = None
        self.setWindowTitle("Office ERP 로그인")
        self.setModal(True)
        self.setFixedWidth(420)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 26, 28, 24)
        layout.setSpacing(14)

        title = QLabel("Office ERP")
        title.setObjectName("loginTitle")
        layout.addWidget(title)

        subtitle = QLabel("아이디와 비밀번호를 입력해 주세요.")
        subtitle.setObjectName("loginSubtitle")
        layout.addWidget(subtitle)

        self.id_input = QLineEdit()
        self.id_input.setPlaceholderText("아이디")
        self.id_input.setObjectName("loginInput")
        layout.addWidget(self.id_input)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("비밀번호")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setObjectName("loginInput")
        self.password_input.returnPressed.connect(self._try_login)
        layout.addWidget(self.password_input)

        self.remember_checkbox = QCheckBox("아이디 / 비밀번호 저장")
        self.remember_checkbox.setObjectName("loginCheck")
        layout.addWidget(self.remember_checkbox)

        saved_login = load_saved_login()
        if saved_login:
            self.id_input.setText(saved_login.get("id", ""))
            self.password_input.setText(saved_login.get("password", ""))
            self.remember_checkbox.setChecked(True)

        self.message_label = QLabel("")
        self.message_label.setObjectName("loginMessage")
        self.message_label.setWordWrap(True)
        layout.addWidget(self.message_label)

        button_row = QHBoxLayout()
        button_row.addStretch(1)

        cancel_btn = QPushButton("닫기")
        cancel_btn.setObjectName("loginCancelButton")
        cancel_btn.clicked.connect(self.reject)
        button_row.addWidget(cancel_btn)

        self.login_btn = QPushButton("로그인")
        self.login_btn.setObjectName("loginButton")
        self.login_btn.clicked.connect(self._try_login)
        button_row.addWidget(self.login_btn)

        layout.addLayout(button_row)
        self.setStyleSheet(
            """
            QDialog {
                background: #FFFFFF;
            }
            QLabel#loginTitle {
                color: #111827;
                font-size: 24px;
                font-weight: 900;
            }
            QLabel#loginSubtitle {
                color: #64748B;
                font-size: 12px;
                font-weight: 600;
            }
            QLabel#loginMessage {
                color: #DC2626;
                font-size: 12px;
                font-weight: 700;
                min-height: 18px;
            }
            QCheckBox#loginCheck {
                color: #475569;
                font-size: 12px;
                font-weight: 700;
                padding-top: 2px;
            }
            QCheckBox#loginCheck::indicator {
                width: 17px;
                height: 17px;
            }
            QLineEdit#loginInput {
                min-height: 42px;
                padding: 0 13px;
                border: 1px solid #D8DEE9;
                border-radius: 10px;
                color: #111827;
                font-size: 14px;
                background: #F8FAFC;
            }
            QLineEdit#loginInput:focus {
                border: 1px solid #2563EB;
                background: #FFFFFF;
            }
            QPushButton {
                min-width: 86px;
                min-height: 38px;
                border-radius: 10px;
                font-size: 13px;
                font-weight: 800;
            }
            QPushButton#loginButton {
                color: #FFFFFF;
                background: #2563EB;
                border: 1px solid #2563EB;
            }
            QPushButton#loginButton:hover {
                background: #1D4ED8;
                border: 1px solid #1D4ED8;
            }
            QPushButton#loginCancelButton {
                color: #2563EB;
                background: #FFFFFF;
                border: 1px solid #BFDBFE;
            }
            QPushButton#loginCancelButton:hover {
                background: #EFF6FF;
            }
            """
        )

    def _set_busy(self, busy: bool):
        self.login_btn.setEnabled(not busy)
        self.id_input.setEnabled(not busy)
        self.password_input.setEnabled(not busy)
        self.remember_checkbox.setEnabled(not busy)
        self.login_btn.setText("확인 중..." if busy else "로그인")
        QApplication.processEvents()

    def _try_login(self):
        self.message_label.setText("")
        self._set_busy(True)
        try:
            session = authenticate_office_user(
                self.id_input.text(),
                self.password_input.text(),
            )
        except Exception as exc:
            self.message_label.setText(f"로그인 정보를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요. ({exc})")
            self._set_busy(False)
            return

        if session is None:
            self.message_label.setText("아이디 또는 비밀번호가 올바르지 않습니다.")
            self._set_busy(False)
            return

        if self.remember_checkbox.isChecked():
            save_login(self.id_input.text().strip(), self.password_input.text().strip())
        else:
            clear_saved_login()

        self.session = session
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self, session: UserSession):
        super().__init__()
        self.session = session
        self.app_structure = build_app_structure(session)
        self.setWindowTitle(f"GROW ERP - {session.user_id}")
        self.resize(1420, 920)
        self.setMinimumSize(1180, 760)

        self.active_section = next(iter(self.app_structure))
        self.active_platform = self.app_structure[self.active_section][0]
        self.nav_buttons: dict[str, QPushButton] = {}
        self.platform_buttons: dict[str, QPushButton] = {}
        self.platform_logs = {
            platform.id: []
            for platforms in self.app_structure.values()
            for platform in platforms
        }
        self.pending_log_lines = {
            platform.id: []
            for platforms in self.app_structure.values()
            for platform in platforms
        }
        self.platform_status = {
            platform.id: ("대기 중", "statusIdle")
            for platforms in self.app_structure.values()
            for platform in platforms
        }
        self.platform_progress = {
            platform.id: "대기 중"
            for platforms in self.app_structure.values()
            for platform in platforms
        }
        self.processes: dict[str, QProcess] = {}
        self.process_meta: dict[str, dict] = {}
        self.log_view_platform_id = ""

        self.log_flush_timer = QTimer(self)
        self.log_flush_timer.setInterval(120)
        self.log_flush_timer.timeout.connect(self._flush_visible_logs)
        self.log_flush_timer.start()

        self._build_ui()
        self._apply_style()
        self._render_section()

    def _confirm_run(self, title: str) -> bool:
        dialog = QMessageBox(self)
        dialog.setWindowTitle("실행 확인")
        dialog.setIcon(QMessageBox.Icon.NoIcon)
        dialog.setText(f"{title} 작업을 실행할까요?")
        dialog.setInformativeText("로그 창에서 진행 상태를 바로 확인할 수 있습니다.")
        dialog.setTextFormat(Qt.TextFormat.PlainText)
        dialog.setStandardButtons(
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        run_button = dialog.button(QMessageBox.StandardButton.Yes)
        cancel_button = dialog.button(QMessageBox.StandardButton.No)
        if run_button is not None:
            run_button.setText("확인")
        if cancel_button is not None:
            cancel_button.setText("취소")
        dialog.setDefaultButton(QMessageBox.StandardButton.Yes)
        dialog.setMinimumWidth(420)
        dialog.setStyleSheet(
            """
            QMessageBox {
                background: #FFFFFF;
            }
            QMessageBox QLabel {
                color: #111827;
            }
            QMessageBox QLabel#qt_msgbox_label {
                font-size: 16px;
                font-weight: 800;
                min-width: 280px;
                line-height: 1.35;
                padding-top: 6px;
            }
            QMessageBox QLabel#qt_msgbox_informativelabel {
                color: #6B7280;
                font-size: 12px;
                font-weight: 500;
                margin-top: 2px;
                line-height: 1.45;
                padding-bottom: 8px;
            }
            QMessageBox QPushButton {
                min-width: 66px;
                min-height: 28px;
                padding: 3px 10px;
                border-radius: 9px;
                border: 1px solid #E5E7EB;
                font-size: 11px;
                font-weight: 700;
                margin-left: 6px;
            }
            QMessageBox QPushButton[text="확인"] {
                background: #2563EB;
                color: #FFFFFF;
                border: 1px solid #2563EB;
            }
            QMessageBox QPushButton[text="확인"]:hover {
                background: #1D4ED8;
                border: 1px solid #1D4ED8;
            }
            QMessageBox QPushButton[text="취소"] {
                background: #FFFFFF;
                color: #2563EB;
                border: 1px solid #BFDBFE;
            }
            QMessageBox QPushButton[text="취소"]:hover {
                background: #EFF6FF;
            }
            QMessageBox QWidget {
                background: #FFFFFF;
            }
            """
        )
        return dialog.exec() == int(QMessageBox.StandardButton.Yes)

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)

        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(280)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(20, 26, 20, 26)
        sidebar_layout.setSpacing(16)

        brand = QLabel("GROW ERP")
        brand.setObjectName("brand")
        sidebar_layout.addWidget(brand)

        menu_title = QLabel("업무 카테고리")
        menu_title.setObjectName("menuTitle")
        sidebar_layout.addWidget(menu_title)

        for section in self.app_structure:
            button = QPushButton(section)
            button.setCheckable(True)
            button.setObjectName("navButton")
            button.clicked.connect(lambda checked=False, name=section: self._set_section(name))
            sidebar_layout.addWidget(button)
            self.nav_buttons[section] = button

        sidebar_layout.addStretch(1)

        version_title = QLabel("현재 설치 버전")
        version_title.setObjectName("versionTitle")
        sidebar_layout.addWidget(version_title)

        self.version_label = QLabel(read_installed_version())
        self.version_label.setObjectName("versionLabel")
        self.version_label.setWordWrap(True)
        sidebar_layout.addWidget(self.version_label)

        outer.addWidget(sidebar)

        main_wrap = QWidget()
        main_layout = QVBoxLayout(main_wrap)
        main_layout.setContentsMargins(28, 28, 28, 28)
        main_layout.setSpacing(16)
        outer.addWidget(main_wrap, 1)

        header = QHBoxLayout()
        header.setSpacing(14)

        self.section_title = QLabel()
        self.section_title.setObjectName("sectionTitle")
        header.addWidget(self.section_title)
        header.addStretch(1)

        open_folder_btn = QPushButton("ERP 폴더 열기")
        open_folder_btn.setObjectName("ghostButton")
        open_folder_btn.clicked.connect(self._open_folder)

        self.stop_btn = QPushButton("실행 중지")
        self.stop_btn.setObjectName("ghostButton")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop_process)

        header.addWidget(open_folder_btn)
        header.addWidget(self.stop_btn)
        main_layout.addLayout(header)

        self.platform_tabs_wrap = QWidget()
        self.platform_tabs_layout = QHBoxLayout(self.platform_tabs_wrap)
        self.platform_tabs_layout.setContentsMargins(0, 0, 0, 0)
        self.platform_tabs_layout.setSpacing(12)
        main_layout.addWidget(self.platform_tabs_wrap)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setObjectName("scrollArea")

        scroll_host = QWidget()
        scroll_layout = QVBoxLayout(scroll_host)
        scroll_layout.setContentsMargins(0, 0, 0, 0)
        scroll_layout.setSpacing(14)

        self.platform_panel = QFrame()
        self.platform_panel.setObjectName("panel")
        platform_layout = QVBoxLayout(self.platform_panel)
        platform_layout.setContentsMargins(24, 24, 24, 24)
        platform_layout.setSpacing(20)

        self.platform_title = QLabel()
        self.platform_title.setObjectName("platformTitle")
        platform_layout.addWidget(self.platform_title)

        self.script_grid = QGridLayout()
        self.script_grid.setHorizontalSpacing(18)
        self.script_grid.setVerticalSpacing(18)
        platform_layout.addLayout(self.script_grid)

        scroll_layout.addWidget(self.platform_panel)

        self.log_panel = QFrame()
        self.log_panel.setObjectName("logPanel")
        log_layout = QVBoxLayout(self.log_panel)
        log_layout.setContentsMargins(18, 18, 18, 18)
        log_layout.setSpacing(12)

        log_top = QHBoxLayout()
        self.log_title = QLabel()
        self.log_title.setObjectName("logTitle")

        log_top.addWidget(self.log_title)
        self.progress_label = QLabel("진행 결과: 대기 중")
        self.progress_label.setObjectName("progressLabel")
        log_top.addWidget(self.progress_label)
        log_top.addStretch(1)

        clear_btn = QPushButton("로그 비우기")
        clear_btn.setObjectName("textButton")
        clear_btn.clicked.connect(self._clear_current_log)

        self.status_badge = QLabel("대기 중")
        self.status_badge.setObjectName("statusIdle")

        log_top.addWidget(clear_btn)
        log_top.addWidget(self.status_badge)

        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("실행 로그가 여기에 표시됩니다.")
        self.log_box.setMaximumHeight(240)

        log_layout.addLayout(log_top)
        log_layout.addWidget(self.log_box)
        scroll_layout.addWidget(self.log_panel)

        scroll.setWidget(scroll_host)
        main_layout.addWidget(scroll, 1)

    def _apply_style(self):
        self.setFont(QFont("Apple SD Gothic Neo", 11))
        self.setStyleSheet(
            """
            QMainWindow {
                background: #FFFFFF;
            }
            QFrame#sidebar {
                background: qlineargradient(
                    x1:0, y1:0, x2:0, y2:1,
                    stop:0 #111827,
                    stop:0.6 #1F2937,
                    stop:1 #374151
                );
            }
            QLabel#brand {
                color: white;
                font-size: 31px;
                font-weight: 800;
                letter-spacing: 0.5px;
            }
            QLabel#menuTitle {
                color: #C7D2E2;
                font-size: 12px;
                font-weight: 800;
                letter-spacing: 1px;
                text-transform: uppercase;
            }
            QLabel#versionTitle {
                color: #94A3B8;
                font-size: 11px;
                font-weight: 800;
            }
            QLabel#versionLabel {
                color: #FFFFFF;
                font-size: 12px;
                font-weight: 800;
                padding: 9px 10px;
                border-radius: 10px;
                background: rgba(255, 255, 255, 0.10);
            }
            QPushButton#navButton {
                background: rgba(255,255,255,0.05);
                color: rgba(255,255,255,0.94);
                border: 1px solid rgba(255,255,255,0.06);
                border-radius: 22px;
                text-align: left;
                padding: 16px 17px;
                font-size: 15px;
                font-weight: 700;
            }
            QPushButton#navButton:hover {
                background: rgba(255,255,255,0.09);
            }
            QPushButton#navButton:checked {
                background: #EFF6FF;
                color: #1D4ED8;
                border: 1px solid #DBEAFE;
            }
            QLabel#sectionTitle {
                color: #111827;
                font-size: 30px;
                font-weight: 800;
                letter-spacing: -0.5px;
            }
            QPushButton#ghostButton {
                background: #FFFFFF;
                color: #111827;
                border: 1px solid #E5E7EB;
                border-radius: 18px;
                padding: 12px 17px;
                font-size: 13px;
                font-weight: 800;
            }
            QPushButton#ghostButton:hover:!disabled {
                background: #F9FAFB;
            }
            QPushButton#ghostButton:disabled {
                color: #9CA3AF;
                background: #F9FAFB;
            }
            QPushButton#platformTab {
                background-color: #F3F4F6;
                color: #374151;
                border: none;
                border-radius: 20px;
                padding: 0 22px;
                min-height: 44px;
                font-size: 13px;
                font-weight: 700;
            }
            QPushButton#platformTab:hover {
                background-color: #E5E7EB;
            }
            QFrame#panel {
                background: #FFFFFF;
                border: 1px solid #F3F4F6;
                border-radius: 30px;
            }
            QLabel#platformTitle {
                color: #111827;
                font-size: 25px;
                font-weight: 800;
                letter-spacing: -0.4px;
            }
            QFrame#scriptCard {
                background: #FFFFFF;
                border: 1px solid #E5E7EB;
                border-radius: 26px;
            }
            QFrame#logPanel {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #111827,
                    stop:1 #1F2937
                );
                border: 1px solid #374151;
                border-radius: 28px;
            }
            QLabel#logTitle {
                color: white;
                font-size: 18px;
                font-weight: 800;
            }
            QLabel#progressLabel {
                color: #CBD5E1;
                font-size: 13px;
                font-weight: 700;
            }
            QLabel#statusIdle, QLabel#statusRunning, QLabel#statusError, QLabel#statusWarn, QLabel#statusSuccess {
                border-radius: 18px;
                padding: 10px 16px;
                min-width: 92px;
                font-size: 12px;
                font-weight: 700;
                qproperty-alignment: 'AlignCenter';
            }
            QLabel#statusIdle {
                background: #374151;
                color: #F9FAFB;
            }
            QLabel#statusRunning {
                background: #2563EB;
                color: #EFF6FF;
            }
            QLabel#statusError {
                background: #111827;
                color: #E5E7EB;
            }
            QLabel#statusWarn {
                background: #4B5563;
                color: #F3F4F6;
            }
            QLabel#statusSuccess {
                background: #047857;
                color: #ECFDF5;
            }
            QPushButton#textButton {
                background: transparent;
                color: #93C5FD;
                border: none;
                font-size: 13px;
                font-weight: 800;
                padding: 8px 10px;
            }
            QPlainTextEdit {
                background: rgba(17,24,39,0.92);
                color: #E5E7EB;
                border: 1px solid #374151;
                border-radius: 18px;
                padding: 14px;
                font-size: 13px;
                selection-background-color: #1D4ED8;
            }
            """
        )

    def _set_section(self, section_name: str):
        self.active_section = section_name
        self.active_platform = self.app_structure[section_name][0]
        self._render_section()

    def _render_section(self):
        for name, button in self.nav_buttons.items():
            button.setChecked(name == self.active_section)

        self.section_title.setText(self.active_section)
        self._render_platform_tabs()
        self._render_platform_panel()
        self._refresh_log_view()

    def _render_platform_tabs(self):
        self.platform_buttons.clear()
        while self.platform_tabs_layout.count():
            item = self.platform_tabs_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        for platform in self.app_structure[self.active_section]:
            button = QPushButton(platform.title)
            button.setObjectName("platformTab")
            button.clicked.connect(lambda checked=False, p=platform: self._set_platform(p))
            if platform.id == self.active_platform.id:
                button.setStyleSheet(
                    "background-color: #2563EB;"
                    "color: white;"
                    "border: none;"
                    "border-radius: 20px;"
                    "padding: 0 22px;"
                    "min-height: 44px;"
                    "font-size: 13px;"
                    "font-weight: 700;"
                )
            self.platform_tabs_layout.addWidget(button)
            self.platform_buttons[platform.id] = button

        self.platform_tabs_layout.addStretch(1)

    def _set_platform(self, platform: PlatformItem):
        self.active_platform = platform
        self._render_section()

    def _render_platform_panel(self):
        self.platform_title.setText(self.active_platform.title)
        while self.script_grid.count():
            item = self.script_grid.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        if not self.active_platform.scripts:
            placeholder = ScriptItem(
                id=f"{self.active_platform.id}_pending",
                title="준비 중",
                filename=None,
                accent=self.active_platform.accent,
            )
            self.script_grid.addWidget(ScriptCard(placeholder, self._run_script), 0, 0)
            return

        for idx, script in enumerate(self.active_platform.scripts):
            row, col = divmod(idx, 3)
            self.script_grid.addWidget(ScriptCard(script, self._run_script), row, col)

    def _append_log(self, platform_id: str, message: str):
        self.platform_logs.setdefault(platform_id, []).append(message)
        self._update_progress_from_log(platform_id, message)
        if self.active_platform.id == platform_id:
            self.pending_log_lines.setdefault(platform_id, []).append(message)

    def _now_text(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _elapsed_text(self, started_at: float):
        if not started_at:
            return "0.0초"
        elapsed = max(0.0, time.monotonic() - started_at)
        return f"{elapsed:.1f}초"

    def _refresh_log_view(self):
        self.log_title.setText(f"{self.active_platform.title} 로그")
        logs = self.platform_logs.get(self.active_platform.id, [])
        self.log_box.setPlainText("\n".join(logs))
        self.pending_log_lines[self.active_platform.id] = []
        self.log_view_platform_id = self.active_platform.id
        self.progress_label.setText(
            f"진행 결과: {self.platform_progress.get(self.active_platform.id, '대기 중')}"
        )
        cursor = self.log_box.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.log_box.setTextCursor(cursor)
        self.stop_btn.setEnabled(self._running_process_count(self.active_platform.id) > 0)
        self._apply_platform_status(self.active_platform.id)

    def _set_status(self, text: str, state_name: str):
        self.status_badge.setText(text)
        self.status_badge.setObjectName(state_name)
        self.status_badge.style().unpolish(self.status_badge)
        self.status_badge.style().polish(self.status_badge)

    def _update_platform_status(self, platform_id: str, text: str, state_name: str):
        self.platform_status[platform_id] = (text, state_name)
        if self.active_platform.id == platform_id:
            self._set_status(text, state_name)

    def _apply_platform_status(self, platform_id: str):
        text, state_name = self.platform_status.get(platform_id, ("대기 중", "statusIdle"))
        self._set_status(text, state_name)

    def _set_platform_progress(self, platform_id: str, text: str):
        cleaned = (text or "").strip()
        if not cleaned:
            cleaned = "대기 중"
        self.platform_progress[platform_id] = cleaned
        if self.active_platform.id == platform_id:
            self.progress_label.setText(f"진행 결과: {cleaned}")

    def _update_progress_from_log(self, platform_id: str, message: str):
        text = (message or "").strip()
        if not text:
            return

        if text.startswith("[START]"):
            self._set_platform_progress(platform_id, "실행 시작")
            return
        if text.startswith("[DONE]"):
            done_text = text.replace("[DONE]", "", 1).strip()
            self._set_platform_progress(platform_id, done_text or "실행 완료")
            return
        if text.startswith("[FAIL]"):
            fail_text = text.replace("[FAIL]", "", 1).strip()
            self._set_platform_progress(platform_id, fail_text or "실행 실패")
            return
        if text.startswith("[STOP]"):
            stop_text = text.replace("[STOP]", "", 1).strip()
            self._set_platform_progress(platform_id, stop_text or "실행 중지")
            return
        if text.startswith("[ERROR]"):
            error_text = text.replace("[ERROR]", "", 1).strip()
            self._set_platform_progress(platform_id, error_text or "오류 발생")
            return
        if text.startswith("[TIME]"):
            return

        self._set_platform_progress(platform_id, text)

    def _notify(self, message: str):
        if sys.platform != "darwin":
            return

        safe_message = message.replace("\\", "\\\\").replace('"', '\\"')
        safe_title = "GROW ERP"
        script = f'display notification "{safe_message}" with title "{safe_title}"'
        try:
            subprocess.Popen(
                ["osascript", "-e", script],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except Exception:
            pass

    def _flush_visible_logs(self):
        self._collect_process_logs()

        platform_id = self.active_platform.id
        pending_lines = self.pending_log_lines.get(platform_id, [])
        if not pending_lines:
            return

        if self.log_view_platform_id != platform_id:
            self._refresh_log_view()
            return

        self.log_box.appendPlainText("\n".join(pending_lines))
        self.pending_log_lines[platform_id] = []
        cursor = self.log_box.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.log_box.setTextCursor(cursor)

    def _collect_process_logs(self):
        for script_id in list(self.process_meta):
            self._read_process_log(script_id)

    def _read_process_log(self, script_id: str):
        meta = self.process_meta.get(script_id)
        if meta is None:
            return

        log_path = meta.get("log_path")
        if not isinstance(log_path, Path) or not log_path.exists():
            return

        offset = int(meta.get("log_offset", 0))
        with log_path.open("rb") as log_file:
            log_file.seek(offset)
            chunk = log_file.read()

        if not chunk:
            return

        meta["log_offset"] = offset + len(chunk)
        text = chunk.decode("utf-8", errors="replace")
        buffered = meta.get("log_buffer", "") + text

        if buffered.endswith("\n") or buffered.endswith("\r"):
            lines = buffered.splitlines()
            meta["log_buffer"] = ""
        else:
            lines = buffered.splitlines()
            if lines:
                meta["log_buffer"] = lines.pop()
            else:
                meta["log_buffer"] = buffered

        for line in lines:
            self._append_log(meta["platform_id"], line)

    def _running_process_count(self, platform_id: str | None = None) -> int:
        if platform_id is None:
            return len(self.processes)
        return sum(
            1
            for meta in self.process_meta.values()
            if meta["platform_id"] == platform_id
        )

    def _run_script(self, item: ScriptItem):
        if not item.filename:
            QMessageBox.information(self, "준비 중", "이 항목은 아직 추가되지 않았습니다.")
            return

        if not self._confirm_run(item.title):
            return

        script_path = resolve_script_path(item.filename)
        if not script_path.exists():
            QMessageBox.critical(self, "파일 없음", f"스크립트를 찾을 수 없습니다.\n\n{script_path}")
            return

        if not PYTHON_BIN.exists():
            QMessageBox.critical(
                self,
                "가상환경 없음",
                f".venv 파이썬을 찾을 수 없습니다.\n\n{PYTHON_BIN}",
            )
            return

        script_id = item.id
        if script_id in self.processes:
            QMessageBox.information(self, "실행 중", f"{item.title} 작업이 이미 실행 중입니다.")
            return

        process = QProcess(self)
        process.setWorkingDirectory(str(BASE_DIR))
        process_environment = QProcessEnvironment.systemEnvironment()
        process_environment.insert("PYTHONUNBUFFERED", "1")
        process_environment.insert("PYTHONUTF8", "1")
        process_environment.insert("PYTHONIOENCODING", "utf-8")
        process_environment.insert("PYTHONPATH", f"{BASE_DIR}{';' if sys.platform.startswith('win') else ':'}{SCRIPT_DIR}")
        process.setProcessEnvironment(process_environment)
        process.setProgram(str(PYTHON_BIN))
        if getattr(sys, "frozen", False):
            process.setArguments(["--run-script", str(script_path)])
        else:
            process.setArguments(["-u", str(script_path)])
        process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        LAUNCHER_LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_path = LAUNCHER_LOG_DIR / f"{script_id}.log"
        log_path.write_text("", encoding="utf-8")
        process.setStandardOutputFile(str(log_path))
        process.finished.connect(
            lambda exit_code, exit_status, sid=script_id: self._handle_finished(sid, exit_code, exit_status)
        )
        process.errorOccurred.connect(lambda _error, sid=script_id: self._handle_error(sid))

        self.processes[script_id] = process
        self.process_meta[script_id] = {
            "platform_id": self.active_platform.id,
            "title": item.title,
            "started_at": time.monotonic(),
            "stop_requested": False,
            "log_path": log_path,
            "log_offset": 0,
            "log_buffer": "",
        }
        self.stop_btn.setEnabled(True)
        self._append_log(
            self.active_platform.id,
            f"[START] {item.title} | {self._now_text()}",
        )
        running_count = self._running_process_count(self.active_platform.id)
        self._update_platform_status(
            self.active_platform.id,
            f"{running_count}개 실행 중",
            "statusRunning",
        )

        process.start()

        if not process.waitForStarted(3000):
            self._append_log(self.active_platform.id, "[ERROR] 프로세스를 시작하지 못했습니다.")
            self._cleanup_process(script_id)

    def _handle_finished(self, script_id: str, exit_code: int, exit_status):
        meta = self.process_meta.get(script_id)
        if meta is None:
            return
        self._read_process_log(script_id)
        if meta.get("log_buffer"):
            self._append_log(meta["platform_id"], meta["log_buffer"])
            meta["log_buffer"] = ""
        platform_id = meta["platform_id"]
        title = meta["title"] or "작업"
        started_at = meta["started_at"]
        stop_requested = meta["stop_requested"]

        if stop_requested:
            self._append_log(platform_id, f"[STOP] {title} 중지")
            self._append_log(
                platform_id,
                f"[TIME] 종료 {self._now_text()} | 소요 {self._elapsed_text(started_at)}",
            )
        elif exit_code == 0:
            self._append_log(platform_id, f"[DONE] {title} 완료")
            self._append_log(
                platform_id,
                f"[TIME] 완료 {self._now_text()} | 소요 {self._elapsed_text(started_at)}",
            )
        else:
            self._append_log(platform_id, f"[FAIL] {title} 실패 (exit code: {exit_code})")
            self._append_log(
                platform_id,
                f"[TIME] 종료 {self._now_text()} | 소요 {self._elapsed_text(started_at)}",
            )

        self._cleanup_process(script_id)

        running_count = self._running_process_count(platform_id)
        if running_count > 0:
            self._update_platform_status(platform_id, f"{running_count}개 실행 중", "statusRunning")
        elif stop_requested:
            self._update_platform_status(platform_id, "실행 중지", "statusWarn")
            self._notify(f"{title} 작업이 중지되었습니다.")
        elif exit_code == 0:
            self._update_platform_status(platform_id, "실행 완료", "statusIdle")
            self._notify(f"{title} 작업이 완료되었습니다.")
            QTimer.singleShot(
                0,
                lambda t=title: QMessageBox.information(
                    self,
                    "작업 완료",
                    f"{t} 작업이 완료되었습니다.",
                ),
            )
        else:
            self._update_platform_status(platform_id, "실행 실패", "statusError")
            self._notify(f"{title} 작업이 실패했습니다.")

    def _handle_error(self, script_id: str):
        meta = self.process_meta.get(script_id)
        process = self.processes.get(script_id)
        if meta is None or process is None:
            return
        if process.state() == QProcess.ProcessState.NotRunning:
            return
        self._append_log(meta["platform_id"], f"[ERROR] {meta['title']} 실행 중 오류가 발생했습니다.")

    def _cleanup_process(self, script_id: str):
        process = self.processes.pop(script_id, None)
        if process is not None:
            process.deleteLater()
        meta = self.process_meta.pop(script_id, None)
        if meta is not None:
            log_path = meta.get("log_path")
            if isinstance(log_path, Path):
                try:
                    log_path.unlink(missing_ok=True)
                except OSError:
                    pass
        self.stop_btn.setEnabled(self._running_process_count(self.active_platform.id) > 0)

    def _stop_process(self):
        active_script_ids = [
            script_id
            for script_id, meta in self.process_meta.items()
            if meta["platform_id"] == self.active_platform.id
        ]
        if not active_script_ids:
            return
        for script_id in active_script_ids:
            process = self.processes.get(script_id)
            meta = self.process_meta.get(script_id)
            if process is None or meta is None:
                continue
            meta["stop_requested"] = True
            process.terminate()
            QTimer.singleShot(
                2000,
                lambda sid=script_id: self._force_kill_if_running(sid),
            )

    def _force_kill_if_running(self, script_id: str):
        process = self.processes.get(script_id)
        if process is not None and process.state() != QProcess.ProcessState.NotRunning:
            process.kill()

    def _clear_current_log(self):
        self.platform_logs[self.active_platform.id] = []
        self.pending_log_lines[self.active_platform.id] = []
        self._refresh_log_view()

    def _open_folder(self):
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", str(BASE_DIR)])
            elif sys.platform.startswith("win"):
                subprocess.Popen(["explorer", str(BASE_DIR)])
            else:
                subprocess.Popen(["xdg-open", str(BASE_DIR)])
        except Exception as exc:
            QMessageBox.critical(self, "폴더 열기 실패", str(exc))


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("GROW ERP")
    login_dialog = LoginDialog()
    if login_dialog.exec() != int(QDialog.DialogCode.Accepted) or login_dialog.session is None:
        sys.exit(0)

    window = MainWindow(login_dialog.session)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
