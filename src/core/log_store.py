"""
JSONL-логирование действий NordFox Module Manager.

Формат:
- один JSONL-файл на сессию;
- каждая запись (event) добавляется append-only одной строкой JSON.
"""

from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Literal

logger = logging.getLogger("JsonLogStore")


class JsonLogStore:
    """Управление JSONL-логом одной сессии."""

    def __init__(self, app_name: str, app_version: str, project_root: Path) -> None:
        self.app_name = app_name
        self.app_version = app_version
        self.project_root = project_root

        self.logs_dir = Path(__file__).resolve().parents[2] / "logs"
        self.logs_dir.mkdir(parents=True, exist_ok=True)

        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"nordfox_module_{ts}.jsonl"
        self.path = self.logs_dir / filename

        self._started_at = datetime.now().isoformat(timespec="seconds")
        self._last_project_state: Dict[str, Any] = {}

        logger.info(f"JSON-лог сессии: {self.path}")
        self._write_event(
            "session_started",
            {
                "app": {
                    "name": self.app_name,
                    "version": self.app_version,
                },
                "session": {
                    "started_at": self._started_at,
                    "project_root": str(self.project_root),
                },
            },
        )

    # ------------------------------------------------------------------
    # Публичные методы
    # ------------------------------------------------------------------

    def set_project_state(self, state: Dict[str, Any]) -> None:
        """Сохранить сводное состояние проекта (assembly, documents, variables_index и т.п.)."""
        self._last_project_state = state
        self._write_event(
            "project_state_updated",
            {"project_state": state},
        )

    def add_action(
        self,
        type_: str,
        status: Literal["success", "error", "partial"],
        input_: Dict[str, Any],
        changes: Dict[str, Any],
        meta: Optional[Dict[str, Any]] = None,
    ) -> None:
        """Добавить запись действия в лог."""
        ts = datetime.now().isoformat(timespec="milliseconds")
        action_id = f"{ts}_{type_}"
        self._write_event(
            "action",
            {
                "id": action_id,
                "type": type_,
                "timestamp": ts,
                "status": status,
                "input": input_,
                "changes": changes,
                "meta": meta or {},
            },
        )

    def close(self) -> None:
        """Завершить сессию и обновить поле ended_at."""
        self._write_event(
            "session_ended",
            {
                "session": {
                    "started_at": self._started_at,
                    "ended_at": datetime.now().isoformat(timespec="seconds"),
                    "project_root": str(self.project_root),
                },
                # Последний снимок состояния даем в финальной записи
                # для удобства пост-анализа одной строкой.
                "project_state": self._last_project_state,
            },
        )

    # ------------------------------------------------------------------
    # Внутренние
    # ------------------------------------------------------------------

    def _write_event(self, event_type: str, payload: Dict[str, Any]) -> None:
        """Append-only запись события в JSONL файл."""
        try:
            record = {
                "event_type": event_type,
                "logged_at": datetime.now().isoformat(timespec="milliseconds"),
                **payload,
            }
            with self.path.open("a", encoding="utf-8") as f:
                f.write(json.dumps(record, ensure_ascii=False) + "\n")
        except Exception as exc:  # pragma: no cover
            logger.error(f"Ошибка записи JSONL-лога: {exc}")

