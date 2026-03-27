from __future__ import annotations

import logging
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import requests

from .stamp_updater import collect_drawings_for_stamps


logger = logging.getLogger("DrawingDwgExporter")


class DrawingDwgExporter:
    def __init__(self, base_url: str = "http://127.0.0.1:5001") -> None:
        self.base_url = base_url.rstrip("/")

    def ensure_service_running(self, timeout_sec: int = 15) -> bool:
        health_url = f"{self.base_url}/health"
        try:
            resp = requests.get(health_url, timeout=1)
            if resp.status_code == 200:
                try:
                    data = resp.json()
                except Exception:
                    data = {}
                if data.get("service") == "kompas-pdf":
                    return True
        except Exception:
            pass

        service_path = Path(__file__).resolve().parent / "kompas_pdf_service.py"
        if not service_path.exists():
            logger.error("Service script not found: %s", service_path)
            return False

        logger.info("Запуск локального CAD сервиса: %s", service_path)
        creation_flags = 0
        if sys.platform.startswith("win"):
            creation_flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
        subprocess.Popen(
            [sys.executable, str(service_path)],
            cwd=str(service_path.parent),
            creationflags=creation_flags,
        )

        for _ in range(timeout_sec):
            try:
                requests.get(health_url, timeout=1)
                return True
            except Exception:
                time.sleep(1)
        logger.error("Локальный CAD сервис не запустился за %s сек", timeout_sec)
        return False

    def export_one_cdw_to_dwg(self, cdw_path: Path, output_dwg: Path) -> dict[str, Any]:
        payload = {
            "input_path": str(cdw_path.resolve()),
            "output_path": str(output_dwg.resolve()),
        }
        try:
            response = requests.post(
                f"{self.base_url}/export_dwg",
                json=payload,
                timeout=90,
            )
        except Exception as exc:
            return {"success": False, "error": f"Service request failed: {exc}"}

        try:
            data = response.json()
        except Exception:
            data = {}

        if response.status_code == 200 and data.get("success"):
            return {
                "success": True,
                "output_path": data.get("output_path"),
                "method": data.get("method"),
            }
        error = data.get("error") or f"HTTP {response.status_code}"
        return {"success": False, "error": str(error)}

    def export_all_drawings_to_dwg(
        self,
        project_root: Path,
        output_folder: Path | None = None,
    ) -> dict[str, Any]:
        result: dict[str, Any] = {
            "success": False,
            "total_drawings": 0,
            "exported_dwgs": 0,
            "failed_drawings": 0,
            "dwg_files": [],
            "errors": [],
        }

        root = Path(project_root).resolve()
        drawings = collect_drawings_for_stamps(root)
        result["total_drawings"] = len(drawings)
        if not drawings:
            result["errors"].append("Чертежи .cdw не найдены")
            return result

        if not self.ensure_service_running():
            result["errors"].append("Не удалось запустить локальный CAD сервис")
            return result

        out_dir = Path(output_folder).resolve() if output_folder else (root / "DWG")
        out_dir.mkdir(parents=True, exist_ok=True)

        for drawing in drawings:
            output_dwg = out_dir / drawing.with_suffix(".dwg").name
            one = self.export_one_cdw_to_dwg(drawing, output_dwg)
            if one.get("success"):
                result["exported_dwgs"] += 1
                result["dwg_files"].append(str(output_dwg))
            else:
                result["failed_drawings"] += 1
                result["errors"].append(f"{drawing.name}: {one.get('error', 'Unknown error')}")

        result["success"] = result["exported_dwgs"] > 0
        return result
