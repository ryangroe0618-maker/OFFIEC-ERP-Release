# -*- coding: utf-8 -*-
import runpy
import sys
import time
from pathlib import Path


if getattr(sys, "frozen", False):
    SEARCH_ROOT = Path(sys.executable).resolve().parent
else:
    SEARCH_ROOT = Path(__file__).resolve().parent


def resolve_project_root(start_dir: Path) -> Path:
    for candidate in (start_dir, *start_dir.parents):
        if (candidate / "현재고 가져오기 최종.py").exists() and (
            candidate / "현재고 변환 및 시트 업로드 최종.py"
        ).exists() and (
            candidate / "거래처 현재고 업로드.py"
        ).exists() and (
            candidate / "kashion 현재고 업로드.py"
        ).exists() and (
            candidate / "플랫폼 poizon 현재고 업로드.py"
        ).exists():
            return candidate
    return start_dir


BASE_DIR = resolve_project_root(SEARCH_ROOT)

SCRIPTS = [
    BASE_DIR / "현재고 가져오기 최종.py",
    BASE_DIR / "현재고 변환 및 시트 업로드 최종.py",
    BASE_DIR / "거래처 현재고 업로드.py",
    BASE_DIR / "kashion 현재고 업로드.py",
    BASE_DIR / "플랫폼 poizon 현재고 업로드.py",
]


def log(message: str, start_time: float):
    elapsed = time.perf_counter() - start_time
    print(f"[{elapsed:6.1f}s] {message}", flush=True)


def run_script(script_path: Path, total_start: float):
    if not script_path.exists():
        raise FileNotFoundError(f"파일이 없습니다: {script_path}")

    step_start = time.perf_counter()
    log(f"{script_path.name} 실행 시작", total_start)

    try:
        runpy.run_path(str(script_path), run_name="__main__")
    except Exception as e:
        raise RuntimeError(f"{script_path.name} 실행 실패: {e}") from e

    step_elapsed = time.perf_counter() - step_start
    log(f"{script_path.name} 실행 완료 ({step_elapsed:.1f}s)", total_start)


def main():
    total_start = time.perf_counter()
    log("현재고 전체 실행 시작", total_start)

    for script_path in SCRIPTS:
        run_script(script_path, total_start)

    log("현재고 전체 실행 완료", total_start)


if __name__ == "__main__":
    main()
