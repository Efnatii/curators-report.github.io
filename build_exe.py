"""Helper script to package the GUI tool into a Windows executable."""

from __future__ import annotations

import base64
import importlib.util
from pathlib import Path
import subprocess
import sys
import os
import urllib.request

from icon_data import ICON_BASE64

FONT_DOWNLOAD_URL = "https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans.ttf"


def ensure_pyinstaller() -> None:
    """Install PyInstaller on demand so the build works out of the box."""

    if importlib.util.find_spec("PyInstaller") is None:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])


def download_font(target_path: Path) -> Path:
    target_path.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(FONT_DOWNLOAD_URL) as response:
        target_path.write_bytes(response.read())
    return target_path


def locate_font() -> Path:
    script_dir = Path(__file__).resolve().parent
    candidates = [
        script_dir / "DejaVuSans.ttf",
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("C:/Windows/Fonts/DejaVuSans.ttf"),
        Path.home() / "Library/Fonts/DejaVuSans.ttf",
    ]

    for candidate in candidates:
        if candidate.exists():
            return candidate

    cache_font = Path.home() / ".cache" / "curators-report" / "DejaVuSans.ttf"
    if cache_font.exists():
        return cache_font

    try:
        return download_font(cache_font)
    except Exception as exc:  # pragma: no cover - relies on network availability
        raise FileNotFoundError(
            "Не удалось найти или скачать DejaVuSans.ttf. Скачайте файл вручную и повторите сборку."
        ) from exc


def build_executable() -> None:
    """Create a standalone .exe for combine_json_to_excel.py."""

    ensure_pyinstaller()

    script_path = Path(__file__).with_name("combine_json_to_excel.py")
    icon_path = script_path.with_name("combine_json_to_excel.ico")
    if not icon_path.exists():
        icon_bytes = base64.b64decode("".join(ICON_BASE64))
        icon_path.write_bytes(icon_bytes)
    if not script_path.exists():
        raise FileNotFoundError(f"Не найден файл {script_path}")
    if not icon_path.exists():
        raise FileNotFoundError(f"Не найден файл {icon_path}")

    font_path = locate_font()

    data_sep = os.pathsep

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name",
        "Combine Json to Excel",
        "--icon",
        str(icon_path),
        "--add-data",
        f"{icon_path}{data_sep}.",
        "--add-data",
        f"{font_path}{data_sep}.",
        str(script_path),
    ]

    subprocess.check_call(command)
    print("Готово! Проверьте папку dist/ для исполняемого файла.")


if __name__ == "__main__":
    build_executable()
