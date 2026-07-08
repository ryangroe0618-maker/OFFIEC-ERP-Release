# -*- coding: utf-8 -*-

from __future__ import annotations

import os
from pathlib import Path


def desktop_dir() -> Path:
    candidates = [
        Path.home() / "Desktop",
        Path.home() / "OneDrive" / "Desktop",
        Path.home() / "OneDrive - Personal" / "Desktop",
        Path(os.environ.get("USERPROFILE", "")) / "Desktop",
    ]
    for candidate in candidates:
        if str(candidate) and candidate.exists():
            return candidate
    return Path.home() / "Desktop"


def base_dir() -> Path:
    path = desktop_dir()
    path.mkdir(parents=True, exist_ok=True)
    return path


def output_dir(name: str) -> Path:
    path = base_dir() / name
    path.mkdir(parents=True, exist_ok=True)
    return path


ROE_DIR = base_dir()
LIST_DIR = output_dir("LIST")
PDF_DIR = output_dir("PDF")
KASHION_DIR = output_dir("KASHION")
CLASSIFY_WAIT_DIR = output_dir("구분 대기")
