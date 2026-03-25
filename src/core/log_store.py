"""
JSON-логирование действий NordFox Module Manager.

Формат логов согласован в обсуждении:
- один JSON-файл на сессию;
- хранится сводное состояние проекта и список actions.
"""

from __future__ import annotations

import json
import logging
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

from .models import JsonSessionLog, JsonLogAction


logger = logging.getLogger("JsonLogStore")


class JsonLogStore:
    """Управление JSON-логом одной сессии."""

    def __init__(self, app_name: str, app_version: str, project_root: Path) -> None:
        self.app_name = app_name
        self.app_version = app_version
        self.project_root = project_root

        self.logs_dir = Path("logs")
        self.logs_dir.mkdir(parents=True, exist_ok=True)

        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"nordfox_module_{ts}.json"
        self.path = self.logs_dir / filename

        self._session = JsonSessionLog(
            app_name=self.app_name,
            app_version=self.app_version,
            started_at=datetime.now().isoformat(timespec="seconds"),
            ended_at=None,
            project_root=self.project_root,
        )

        logger.info(f"JSON-лог сессии: {self.path}")

    # ------------------------------------------------------------------
    # Публичные методы
    # ------------------------------------------------------------------

    def set_project_state(self, state: Dict[str, Any]) -> None:
        """Сохранить сводное состояние проекта (assembly, documents, variables_index и т.п.)."""
        self._session.project_state = state
        self._flush()

    def add_action(
        self,
        type_: str,
        status: str,
        input_: Dict[str, Any],
        changes: Dict[str, Any],
        meta: Optional[Dict[str, Any]] = None,
    ) -> None:
        """Добавить запись действия в лог."""
        ts = datetime.now().isoformat(timespec="milliseconds")
        action_id = f"{ts}_{type_}"

        action = JsonLogAction(
            id=action_id,
            type=type_,
            timestamp=ts,
            status=status,  # type: ignore[arg-type]
            input=input_,
            changes=changes,
            meta=meta or {},
        )
        self._session.actions.append(action)
        self._flush()

    def close(self) -> None:
        """Завершить сессию и обновить поле ended_at."""
        self._session.ended_at = datetime.now().isoformat(timespec="seconds")
        self._flush()

    # ------------------------------------------------------------------
    # Внутренние
    # ------------------------------------------------------------------

    def _flush(self) -> None:
        """Сериализовать JsonSessionLog в файл."""
        try:
            data = {
                "app": {
                    "name": self._session.app_name,
                    "version": self._session.app_version,
                },
                "session": {
                    "started_at": self._session.started_at,
                    "ended_at": self._session.ended_at,
                    "project_root": str(self._session.project_root),
                },
                "project_state": self._session.project_state,
                "actions": [asdict(a) for a in self._session.actions],
            }
            with self.path.open("w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as exc:  # pragma: no cover
            logger.error(f"Ошибка записи JSON-лога: {exc}")

