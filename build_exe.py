"""Helper script to package the GUI tool into a Windows executable."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import subprocess
import sys
import os
import base64

from icon_data import ICON_BASE64


def ensure_pyinstaller() -> None:
    """Install PyInstaller on demand so the build works out of the box."""

    if importlib.util.find_spec("PyInstaller") is None:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])


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
        str(script_path),
    ]

    subprocess.check_call(command)
    print("Готово! Проверьте папку dist/ для исполняемого файла.")


if __name__ == "__main__":
    build_executable()
