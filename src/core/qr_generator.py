"""
Генерация QR-кодов в PNG для NordFox Module Manager.

Основывается на подходе из:
- C:\\Users\\Vorob\\PycharmProjects\\zvdProject\\kompas3d_project_manager\\components\\qr_generator.py

Задачи:
- сгенерировать PNG-файл QR-кода по заданной строке данных;
- использовать библиотеку segno;
- быть простой точкой входа для UI-слоя.
"""

from __future__ import annotations

import logging
from pathlib import Path

import segno  # type: ignore[import-untyped]


logger = logging.getLogger("QRGenerator")


def generate_qr_png(
    data: str,
    png_path: str | Path,
    scale: int = 10,
    border: int = 4,
) -> bool:
    """
    Сгенерировать QR-код в PNG.

    Args:
        data: строка данных для кодирования (NF;PRF=H20;L=2360;ID=0001 и т.п.).
        png_path: путь к PNG-файлу (будет создана соответствующая папка).
        scale: масштаб модуля (пикселей на модуль).
        border: quiet zone в модулях.
    """
    try:
        if not data:
            logger.error("Пустые данные для QR-кода")
            return False

        path = Path(png_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        logger.info(f"Генерация QR PNG: {path.name}")
        logger.info(f"  Данные: {data}")

        qr = segno.make(data, error="m")
        qr.save(
            str(path),
            kind="png",
            scale=scale,
            border=border,
        )

        logger.info(f"QR PNG создан: {path}")
        return True
    except Exception as exc:  # pragma: no cover - обёртка над библиотекой
        logger.error(f"Ошибка генерации QR PNG: {exc}")
        return False

