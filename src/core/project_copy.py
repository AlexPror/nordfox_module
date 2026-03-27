"""
Копирование проекта NordFox перед обновлением.
"""

from __future__ import annotations

import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict


logger = logging.getLogger("ProjectCopy")


def _sanitize_folder_name(raw_name: str) -> str:
    """Подготовить имя папки к ограничениям Windows."""
    name = (raw_name or "").strip()
    # Убираем запрещенные символы Windows и управляющие коды.
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", name)
    # Нормализуем пробелы и запрещенный завершающий суффикс.
    name = re.sub(r"\s+", " ", name).strip().rstrip(". ")
    return name or "project_copy"


def _pick_available_name(target_parent: Path, base_name: str) -> str:
    """Вернуть свободное имя папки, добавляя суффикс _N при коллизии."""
    if not (target_parent / base_name).exists():
        return base_name
    idx = 2
    while True:
        candidate = f"{base_name}_{idx}"
        if not (target_parent / candidate).exists():
            return candidate
        idx += 1


def _ignore_temp_files(_: str, names: list[str]) -> list[str]:
    ignored: list[str] = []
    for name in names:
        low = name.lower()
        if low.endswith(".bak"):
            ignored.append(name)
        elif low.endswith(".tmp"):
            ignored.append(name)
        elif low.endswith(".temp"):
            ignored.append(name)
        elif low.endswith(".lock"):
            ignored.append(name)
        elif low.endswith(".cd~"):
            ignored.append(name)
        elif name.startswith("~$"):
            ignored.append(name)
        elif name.startswith("~"):
            ignored.append(name)
        elif name in ("Thumbs.db", ".DS_Store"):
            ignored.append(name)
    return ignored


def copy_project_tree(source_root: Path, target_parent: Path, new_name: str | None = None) -> Dict[str, object]:
    """
    Скопировать проект в отдельную папку.
    Возвращает словарь с результатом операции.
    """
    source_root = source_root.resolve()
    target_parent = target_parent.resolve()

    result: Dict[str, object] = {
        "success": False,
        "source": str(source_root),
        "target": None,
        "copied_files": 0,
        "error": None,
    }

    if not source_root.exists() or not source_root.is_dir():
        result["error"] = f"Исходная папка не найдена: {source_root}"
        return result

    target_parent.mkdir(parents=True, exist_ok=True)

    if not new_name:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"{source_root.name}_copy_{stamp}"
    else:
        base_name = _sanitize_folder_name(new_name)

    new_name = _pick_available_name(target_parent, base_name)
    target_root = target_parent / new_name

    logger.info("Копирование проекта: %s -> %s", source_root, target_root)

    try:
        shutil.copytree(source_root, target_root, ignore=_ignore_temp_files)
    except Exception as exc:
        result["error"] = f"Ошибка копирования: {exc}"
        return result

    copied_files = sum(1 for p in target_root.rglob("*") if p.is_file())
    result["success"] = True
    result["target"] = str(target_root)
    result["copied_files"] = copied_files
    logger.info("Копирование завершено, файлов: %d", copied_files)
    return result

