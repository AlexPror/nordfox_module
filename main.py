#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NordFox Module Manager
Десктопное приложение для обновления переменных, обозначений и наименований
сборок модулей фасадов NordFox и генерации QR-кодов (PNG).
"""

import sys
import os
import logging

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QFont


def setup_logging() -> None:
    """Базовая настройка логирования для приложения."""
    os.makedirs("logs", exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%H:%M:%S",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(
                os.path.join("logs", "nordfox_module.log"),
                encoding="utf-8",
            ),
        ],
    )


def main() -> None:
    """Точка входа в приложение NordFox Module Manager."""
    # Добавляем в sys.path корень проекта, чтобы можно было импортировать src.*
    project_root = os.path.dirname(os.path.abspath(__file__))
    if project_root not in sys.path:
        sys.path.insert(0, project_root)

    setup_logging()
    logger = logging.getLogger(__name__)
    logger.info("Запуск NordFox Module Manager")

    from src.ui.main_window import MainWindow  # импорт после настройки пути и логов

    app = QApplication(sys.argv)
    app.setApplicationName("NordFox Module Manager")
    from src.__version__ import __version__
    app.setApplicationVersion(__version__)

    # Стиль и шрифт — аналогично nordfox_specification
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 9))

    window = MainWindow()
    window.show()

    logger.info("Приложение запущено")
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

