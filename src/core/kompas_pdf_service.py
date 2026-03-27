from __future__ import annotations

import logging
import os
import time
from pathlib import Path
from typing import Any

import pythoncom
from flask import Flask, jsonify, request


logger = logging.getLogger("KompasPdfService")
app = Flask(__name__)
SERVICE_VERSION = "v2"
_CACHED_RTP_PATH: Path | None = None


def _get_dynamic_dispatch(prog_id: str) -> Any:
    from win32com.client import dynamic  # type: ignore[import-untyped]

    return dynamic.Dispatch(prog_id)


def _find_pdf_converter_rtp(app7: Any) -> Path | None:
    """
    Поиск PdfConverter.rtp для разных установок КОМПАС.
    В некоторых версиях app7.Application.Path пустой, поэтому держим
    набор fallback-путей (как в образцовом проекте).
    """
    global _CACHED_RTP_PATH
    if _CACHED_RTP_PATH and _CACHED_RTP_PATH.exists():
        return _CACHED_RTP_PATH

    roots: list[Path] = []
    for attr_chain in (("Application", "Path"), ("Path",)):
        try:
            obj = app7
            for attr in attr_chain:
                obj = getattr(obj, attr)
            raw = str(obj or "").strip()
            if raw:
                roots.append(Path(raw))
        except Exception:
            pass

    standard_bases = [
        r"C:\Program Files\ASCON\KOMPAS-3D v24",
        r"C:\Program Files\ASCON\KOMPAS-3D v23 Home",
        r"C:\Program Files\ASCON\KOMPAS-3D v23",
        r"C:\Program Files\ASCON\KOMPAS-3D v22",
        r"C:\Program Files\ASCON\KOMPAS-3D v21",
        r"C:\Program Files (x86)\ASCON\KOMPAS-3D v21",
    ]
    for base in standard_bases:
        roots.append(Path(base))

    seen: set[str] = set()
    for root in roots:
        key = str(root).lower()
        if key in seen:
            continue
        seen.add(key)
        for sub in ("Bin", "Libs"):
            candidate = root / sub / "PdfConverter.rtp"
            if candidate.exists():
                _CACHED_RTP_PATH = candidate
                return candidate

    # Расширенный поиск по установочным папкам ASCON.
    for scan_root in (Path(r"C:\Program Files\ASCON"), Path(r"C:\Program Files (x86)\ASCON")):
        try:
            if not scan_root.exists():
                continue
            for candidate in scan_root.rglob("PdfConverter.rtp"):
                if candidate.exists():
                    _CACHED_RTP_PATH = candidate
                    return candidate
        except Exception:
            pass
    return None


def _convert_with_iconverter(input_path: str, output_path: str) -> tuple[bool, str]:
    try:
        app7 = _get_dynamic_dispatch("Kompas.Application.7")
        app7.Visible = True

        converter_path = _find_pdf_converter_rtp(app7)
        if converter_path is None:
            return False, "PdfConverter.rtp not found (all known locations)"

        converter = app7.Converter(str(converter_path))
        doc7 = app7.Documents.Open(input_path, False, False)
        if not doc7:
            return False, "Open failed (IConverter)"
        try:
            try:
                doc7.RebuildDocument()
                time.sleep(0.3)
            except Exception:
                pass
            params = converter.ConverterParameters(0)
            ok = bool(converter.Convert(doc7, output_path, False, params))
            return ok, "IConverter" if ok else "IConverter returned False"
        finally:
            try:
                doc7.Close(0)
            except Exception:
                pass
    except Exception as exc:
        return False, f"IConverter error: {exc}"


def _convert_with_api5_save_to_pdf(input_path: str, output_path: str) -> tuple[bool, str]:
    try:
        app5 = _get_dynamic_dispatch("Kompas.Application.5")
        doc2d = app5.Document2D()
        if not doc2d.ksOpenDocument(input_path, 0):
            return False, "API5 open failed"
        try:
            if not hasattr(doc2d, "ksSaveToPDF"):
                return False, "API5 has no ksSaveToPDF"
            ok = bool(doc2d.ksSaveToPDF(output_path))
            if ok and Path(output_path).exists() and Path(output_path).stat().st_size > 100:
                return True, "ksSaveToPDF"
            return False, "ksSaveToPDF returned False"
        finally:
            try:
                doc2d.ksCloseDocument()
            except Exception:
                pass
    except Exception as exc:
        return False, f"API5 error: {exc}"


def _convert_with_saveas(input_path: str, output_path: str) -> tuple[bool, str]:
    try:
        app7 = _get_dynamic_dispatch("Kompas.Application.7")
        app7.Visible = True
        doc7 = app7.Documents.Open(input_path, True, False)
        if not doc7:
            return False, "Open failed (SaveAs)"
        try:
            try:
                app7.ActiveDocument = doc7
            except Exception:
                pass
            try:
                doc7.RebuildDocument()
                time.sleep(0.3)
            except Exception:
                pass
            output_obj = Path(output_path)
            errors: list[str] = []

            def _is_pdf_ok() -> bool:
                return output_obj.exists() and output_obj.stat().st_size > 100

            # Вариант 1: классический SaveAs
            try:
                ok = bool(doc7.SaveAs(str(output_obj)))
                if (ok or _is_pdf_ok()) and _is_pdf_ok():
                    return True, "SaveAs"
                errors.append(f"SaveAs returned {ok}")
            except Exception as exc:
                errors.append(f"SaveAs exc: {exc}")

            # Вариант 2: SaveAs3 + пустой формат
            try:
                ok3 = bool(doc7.SaveAs3(str(output_obj), "", False, 0))
                if (ok3 or _is_pdf_ok()) and _is_pdf_ok():
                    return True, "SaveAs3-empty"
                errors.append(f"SaveAs3('',0) returned {ok3}")
            except Exception as exc:
                errors.append(f"SaveAs3('',0) exc: {exc}")

            # Вариант 3/4: SaveAs3 + явный pdf/PDF
            for fmt in ("pdf", "PDF"):
                try:
                    okf = bool(doc7.SaveAs3(str(output_obj), fmt, False, 0))
                    if (okf or _is_pdf_ok()) and _is_pdf_ok():
                        return True, f"SaveAs3-{fmt}"
                    errors.append(f"SaveAs3('{fmt}',0) returned {okf}")
                except Exception as exc:
                    errors.append(f"SaveAs3('{fmt}',0) exc: {exc}")

            return False, " | ".join(errors) if errors else "SaveAs/SaveAs3 failed"
        finally:
            try:
                doc7.Close(0)
            except Exception:
                pass
    except Exception as exc:
        return False, f"SaveAs error: {exc}"


def _export_direct_dwg_with_saveas(input_path: str, output_path: str) -> tuple[bool, str]:
    """
    Прямой экспорт в DWG без промежуточного DXF.
    Основан на SaveAs/SaveAs3 у документа чертежа.
    """
    try:
        app7 = _get_dynamic_dispatch("Kompas.Application.7")
        app7.Visible = True
        doc7 = app7.Documents.Open(input_path, True, False)
        if not doc7:
            return False, "Open failed (DWG SaveAs)"
        try:
            try:
                app7.ActiveDocument = doc7
            except Exception:
                pass
            try:
                doc7.RebuildDocument()
                time.sleep(0.3)
            except Exception:
                pass

            ok = bool(doc7.SaveAs(output_path))
            if ok and Path(output_path).exists() and Path(output_path).stat().st_size > 100:
                return True, "SaveAs(DWG)"

            try:
                ok3 = bool(doc7.SaveAs3(output_path, "", False, 0))
            except Exception:
                ok3 = False
            if ok3 and Path(output_path).exists() and Path(output_path).stat().st_size > 100:
                return True, "SaveAs3(DWG)"
            return False, "Direct DWG export failed (SaveAs/SaveAs3)"
        finally:
            try:
                doc7.Close(0)
            except Exception:
                pass
    except Exception as exc:
        return False, f"Direct DWG error: {exc}"


@app.get("/health")
def health_check():
    return jsonify({"status": "running", "service": "kompas-pdf", "version": SERVICE_VERSION})


@app.post("/export")
def export_pdf():
    pythoncom.CoInitialize()
    try:
        data = request.get_json(silent=True) or {}
        input_path = str(data.get("input_path", "") or "").strip()
        output_path = str(data.get("output_path", "") or "").strip()
        if not input_path or not output_path:
            return jsonify({"success": False, "error": "input_path/output_path required"}), 400

        input_file = Path(input_path).resolve()
        output_file = Path(output_path).resolve()
        if not input_file.exists():
            return jsonify({"success": False, "error": f"Input not found: {input_file}"}), 404
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists():
            try:
                output_file.unlink()
            except Exception:
                pass

        # Практика показывает, что IConverter в некоторых окружениях может подвисать.
        # Ставим стабильные методы первыми; IConverter оставляем последним как fallback.
        attempts = (
            _convert_with_api5_save_to_pdf,
            _convert_with_saveas,
            _convert_with_iconverter,
        )
        errors: list[str] = []
        for fn in attempts:
            ok, method_or_error = fn(str(input_file), str(output_file))
            if ok and output_file.exists() and output_file.stat().st_size > 100:
                return (
                    jsonify(
                        {
                            "success": True,
                            "method": method_or_error,
                            "output_path": str(output_file),
                        }
                    ),
                    200,
                )
            errors.append(f"{fn.__name__}: {method_or_error}")

        return jsonify({"success": False, "error": " | ".join(errors)}), 500
    except Exception as exc:
        logger.exception("Unhandled /export error")
        return jsonify({"success": False, "error": str(exc)}), 500
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


@app.post("/export_dwg")
def export_dwg():
    pythoncom.CoInitialize()
    try:
        data = request.get_json(silent=True) or {}
        input_path = str(data.get("input_path", "") or "").strip()
        output_path = str(data.get("output_path", "") or "").strip()
        if not input_path or not output_path:
            return jsonify({"success": False, "error": "input_path/output_path required"}), 400

        input_file = Path(input_path).resolve()
        output_file = Path(output_path).resolve()
        if not input_file.exists():
            return jsonify({"success": False, "error": f"Input not found: {input_file}"}), 404
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists():
            try:
                output_file.unlink()
            except Exception:
                pass

        ok, method_or_error = _export_direct_dwg_with_saveas(str(input_file), str(output_file))
        if ok and output_file.exists() and output_file.stat().st_size > 100:
            return (
                jsonify(
                    {
                        "success": True,
                        "method": method_or_error,
                        "output_path": str(output_file),
                    }
                ),
                200,
            )
        return jsonify({"success": False, "error": method_or_error}), 500
    except Exception as exc:
        logger.exception("Unhandled /export_dwg error")
        return jsonify({"success": False, "error": str(exc)}), 500
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    )
    # Используем новый порт по умолчанию, чтобы не конфликтовать со "старыми"
    # уже запущенными экземплярами сервиса из предыдущих версий.
    port = int(os.environ.get("NORDFOX_PDF_SERVICE_PORT", "5001"))
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
