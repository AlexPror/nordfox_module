from __future__ import annotations

import logging
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import requests

try:
    from pypdf import PdfMerger, PdfReader
except Exception:  # pragma: no cover
    PdfMerger = None
    PdfReader = None

from .stamp_updater import collect_drawings_for_stamps


logger = logging.getLogger("DrawingPdfExporter")


class DrawingPdfExporter:
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
            logger.error("PDF service script not found: %s", service_path)
            return False

        logger.info("Запуск локального PDF сервиса: %s", service_path)
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
        logger.error("Локальный PDF сервис не запустился за %s сек", timeout_sec)
        return False

    def export_one_cdw_to_pdf(self, cdw_path: Path, output_pdf: Path) -> dict[str, Any]:
        payload = {
            "input_path": str(cdw_path.resolve()),
            "output_path": str(output_pdf.resolve()),
        }
        try:
            response = requests.post(
                f"{self.base_url}/export",
                json=payload,
                timeout=60,
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

    def merge_pdf_files(self, pdf_files: list[Path], merged_path: Path) -> dict[str, Any]:
        local_pdf_merger = PdfMerger
        local_pdf_reader = PdfReader
        if local_pdf_merger is None or local_pdf_reader is None:
            # Lazy import: позволяет подхватить pypdf, установленный уже после старта UI.
            try:
                from pypdf import PdfMerger as _PdfMerger, PdfReader as _PdfReader

                local_pdf_merger = _PdfMerger
                local_pdf_reader = _PdfReader
            except Exception:
                try:
                    from PyPDF2 import PdfMerger as _PdfMerger, PdfReader as _PdfReader

                    local_pdf_merger = _PdfMerger
                    local_pdf_reader = _PdfReader
                except Exception:
                    return {
                        "success": False,
                        "error": f"pypdf is not installed (python: {sys.executable})",
                    }

        valid = [p for p in pdf_files if p.exists() and p.stat().st_size > 100]
        if not valid:
            return {"success": False, "error": "No valid PDF files to merge"}

        merger = local_pdf_merger()
        merged_count = 0
        skipped: list[str] = []
        try:
            for pdf in valid:
                try:
                    # Часть PDF из КОМПАС может создаваться с нестандартной структурой.
                    # Проверяем читаемость и число страниц, иначе пропускаем файл.
                    with pdf.open("rb") as f:
                        header = f.read(4)
                    if header != b"%PDF":
                        skipped.append(f"{pdf.name}: bad header")
                        continue

                    reader = local_pdf_reader(str(pdf))
                    if len(reader.pages) == 0:
                        skipped.append(f"{pdf.name}: no pages")
                        continue

                    merger.append(str(pdf))
                    merged_count += 1
                except Exception as exc:
                    skipped.append(f"{pdf.name}: {exc}")

            if merged_count == 0:
                return {
                    "success": False,
                    "error": "No readable PDF files to merge",
                    "skipped": skipped,
                }

            merged_path.parent.mkdir(parents=True, exist_ok=True)
            merger.write(str(merged_path))
            return {
                "success": True,
                "output_file": str(merged_path),
                "merged_count": merged_count,
                "skipped": skipped,
            }
        except Exception as exc:
            return {"success": False, "error": f"PDF merge failed: {exc}"}
        finally:
            merger.close()

    def export_all_drawings_to_pdf(
        self,
        project_root: Path,
        output_folder: Path | None = None,
        merge_into_one: bool = True,
        merged_output_name: str | None = None,
    ) -> dict[str, Any]:
        result: dict[str, Any] = {
            "success": False,
            "total_drawings": 0,
            "exported_pdfs": 0,
            "failed_drawings": 0,
            "pdf_files": [],
            "merged_pdf": None,
            "errors": [],
        }

        root = Path(project_root).resolve()
        drawings = collect_drawings_for_stamps(root)
        result["total_drawings"] = len(drawings)
        if not drawings:
            result["errors"].append("Чертежи .cdw не найдены")
            return result

        if not self.ensure_service_running():
            result["errors"].append("Не удалось запустить локальный PDF сервис")
            return result

        out_dir = Path(output_folder).resolve() if output_folder else (root / "PDF")
        out_dir.mkdir(parents=True, exist_ok=True)

        for drawing in drawings:
            output_pdf = out_dir / drawing.with_suffix(".pdf").name
            one = self.export_one_cdw_to_pdf(drawing, output_pdf)
            if one.get("success"):
                result["exported_pdfs"] += 1
                result["pdf_files"].append(str(output_pdf))
            else:
                # В некоторых версиях КОМПАС файл может успешно записаться,
                # но сервис возвращает ошибку по одному из fallback-методов.
                # Если PDF физически есть и не пустой — считаем экспорт успешным,
                # но оставляем предупреждение для диагностики.
                if output_pdf.exists() and output_pdf.stat().st_size > 100:
                    result["exported_pdfs"] += 1
                    result["pdf_files"].append(str(output_pdf))
                    result["errors"].append(
                        f"{drawing.name}: сервис вернул ошибку, но PDF создан ({one.get('error', 'Unknown error')})"
                    )
                else:
                    result["failed_drawings"] += 1
                    result["errors"].append(f"{drawing.name}: {one.get('error', 'Unknown error')}")

        if merge_into_one and result["pdf_files"]:
            merged_name = merged_output_name or f"{root.name} - все чертежи.pdf"
            if not merged_name.lower().endswith(".pdf"):
                merged_name = f"{merged_name}.pdf"
            merge_res = self.merge_pdf_files([Path(p) for p in result["pdf_files"]], out_dir / merged_name)
            if merge_res.get("success"):
                result["merged_pdf"] = merge_res.get("output_file")
                skipped = merge_res.get("skipped") or []
                for msg in skipped:
                    result["errors"].append(f"Merge skip: {msg}")
            else:
                result["errors"].append(merge_res.get("error", "PDF merge error"))
                skipped = merge_res.get("skipped") or []
                for msg in skipped:
                    result["errors"].append(f"Merge skip: {msg}")

        result["success"] = result["exported_pdfs"] > 0
        return result
