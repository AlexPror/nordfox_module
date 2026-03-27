"""
Базовый модуль для подключения к КОМПАС-3D (API 5 и 7)
Адаптировано из:
- C:\\nordfox_specification\\src\\core\\kompas_connector.py
- C:\\Users\\Vorob\\PycharmProjects\\zvdProject\\kompas3d_project_manager\\components\\base_component.py
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any, Optional

import pythoncom


def get_dynamic_dispatch(prog_id: str) -> Any:
    """Получение COM-объекта через dynamic dispatch (без кэша win32com.gen_py)."""
    from win32com.client import dynamic  # type: ignore[import-untyped]

    return dynamic.Dispatch(prog_id)


class KompasConnector:
    """
    Унифицированное подключение к КОМПАС-3D.

    - API7: основное приложение и документы (Application.7)
    - API5: специфичные функции (например, сохранение в PDF, 2D-документы)
    """

    def __init__(self) -> None:
        self.logger = logging.getLogger(self.__class__.__name__)
        self.api5: Any | None = None
        self.api7: Any | None = None
        self.application: Any | None = None
        self._connected: bool = False
        # Флаг: активный документ уже был открыт в КОМПАС до нашего open_document().
        self._active_doc_was_preopened: bool = False

    # ------------------------------------------------------------------
    # Подключение / отключение
    # ------------------------------------------------------------------

    def connect(self, force_reconnect: bool = False) -> bool:
        """Подключение к КОМПАС-3D (инициализация API5 и API7)."""
        try:
            if force_reconnect or not self._connected:
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    # COM уже инициализирован в этом потоке
                    pass

                self.logger.info("Подключение к КОМПАС-3D...")

                # API5 и API7 через dynamic dispatch
                self.api5 = get_dynamic_dispatch("Kompas.Application.5")
                self.api7 = get_dynamic_dispatch("Kompas.Application.7")
                self.application = self.api7

                # Делаем окно видимым (важно для Home-версий)
                try:
                    self.application.Visible = True
                except Exception:
                    pass

                self.logger.info("Подключение к КОМПАС-3D выполнено успешно")
                self._connected = True
                return True

            # Уже подключены — проверяем, что соединение живое
            if self._connected and self.application is not None:
                try:
                    _ = self.application.Visible
                    return True
                except Exception:
                    self.logger.warning("Соединение с КОМПАС-3D потеряно, переподключение...")
                    self._connected = False
                    return self.connect(force_reconnect=True)

            return True

        except Exception as exc:  # pragma: no cover - обёртка над внешним API
            self.logger.error(f"Ошибка подключения к КОМПАС-3D: {exc}")
            self._connected = False
            return False

    def disconnect(self) -> None:
        """Отключение от КОМПАС-3D и освобождение COM."""
        try:
            self.application = None
            self.api5 = None
            self.api7 = None

            if self._connected:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

                self._connected = False
                self.logger.info("Отключение от КОМПАС-3D выполнено")
        except Exception as exc:  # pragma: no cover
            self.logger.error(f"Ошибка отключения от КОМПАС-3D: {exc}")

    # ------------------------------------------------------------------
    # Вспомогательные методы
    # ------------------------------------------------------------------

    @property
    def is_connected(self) -> bool:
        """Флаг подключения."""
        return self._connected and self.application is not None

    def get_api7(self) -> Any | None:
        """Вернуть API7 (Application.7) c автоподключением при необходимости."""
        if not self.is_connected:
            if not self.connect():
                return None
        return self.api7

    def get_api5(self) -> Any | None:
        """Вернуть API5 (Application.5) c автоподключением при необходимости."""
        if not self.is_connected:
            if not self.connect():
                return None
        return self.api5

    # ------------------------------------------------------------------
    # Работа с документами
    # ------------------------------------------------------------------

    def open_document(self, file_path: str) -> bool:
        """
        Открыть документ КОМПАС-3D по абсолютному пути.
        Возвращает True при успехе.
        """
        try:
            if not self.is_connected:
                if not self.connect():
                    return False

            path = Path(file_path).resolve()
            if not path.exists():
                self.logger.error(f"Файл не найден: {path}")
                return False

            app7 = self.get_api7()
            if app7 is None:
                self.logger.error("API7 недоступен")
                return False

            self.logger.info(f"Открытие документа: {path}")
            # Если документ уже открыт в КОМПАС, переиспользуем его, а не открываем дубликат.
            self._active_doc_was_preopened = False
            doc = None
            try:
                docs = app7.Documents
                for idx in range(docs.Count):
                    d = docs.Item(idx)
                    try:
                        full_name = str(getattr(d, "FullName", "") or "").strip()
                    except Exception:
                        full_name = ""
                    if full_name and Path(full_name).resolve() == path:
                        doc = d
                        self._active_doc_was_preopened = True
                        break
            except Exception:
                # Если перечислить открытые документы не удалось, просто идём по стандартному пути.
                pass

            if doc is None:
                doc = app7.Documents.Open(str(path), False)
            if not doc:
                self.logger.error(f"Не удалось открыть документ: {path}")
                return False

            # Небольшая задержка, чтобы документ полностью подгрузился
            import time

            time.sleep(0.5)

            try:
                app7.ActiveDocument = doc
            except Exception:
                pass

            try:
                self.logger.info(f"Документ открыт: {doc.Name}")
            except Exception:
                self.logger.info(f"Документ открыт: {path.name}")

            return True

        except Exception as exc:  # pragma: no cover
            self.logger.error(f"Ошибка открытия документа {file_path}: {exc}")
            return False

    def close_active_document(self, save: bool = False) -> bool:
        """Закрыть активный документ (с сохранением или без)."""
        try:
            if not self.is_connected:
                return False

            app7 = self.get_api7()
            if app7 is None:
                return False

            active_doc = app7.ActiveDocument
            if not active_doc:
                return False

            name = getattr(active_doc, "Name", "<unknown>")
            # Если документ был уже открыт пользователем до нашего вызова open_document(),
            # не закрываем его насильно. При необходимости просто сохраняем.
            if self._active_doc_was_preopened:
                if save:
                    try:
                        active_doc.Save()
                        self.logger.info(f"Документ сохранен (preopened): {name}")
                    except Exception as exc:
                        self.logger.warning(f"Не удалось сохранить preopened документ {name}: {exc}")
                else:
                    self.logger.info(f"Документ оставлен открытым (preopened): {name}")
                self._active_doc_was_preopened = False
                return True

            active_doc.Close(save)
            self.logger.info(f"Документ закрыт: {name} (save={save})")
            self._active_doc_was_preopened = False
            return True

        except Exception as exc:  # pragma: no cover
            self.logger.error(f"Ошибка закрытия документа: {exc}")
            return False

    def close_all_documents(self) -> bool:
        """Закрыть все открытые документы КОМПАС-3D."""
        try:
            if not self.is_connected:
                return False

            app7 = self.get_api7()
            if app7 is None:
                return False

            docs = app7.Documents
            count = docs.Count

            for idx in range(count - 1, -1, -1):
                try:
                    doc = docs.Item(idx)
                    if doc:
                        name = getattr(doc, "Name", f"doc_{idx}")
                        self.logger.info(f"Закрытие документа: {name}")
                        doc.Close(False)
                except Exception as inner_exc:  # pragma: no cover
                    self.logger.warning(f"Ошибка закрытия документа {idx}: {inner_exc}")

            import time

            time.sleep(0.5)
            self.logger.info("Все документы КОМПАС-3D закрыты")
            return True

        except Exception as exc:  # pragma: no cover
            self.logger.error(f"Ошибка закрытия всех документов: {exc}")
            return False

