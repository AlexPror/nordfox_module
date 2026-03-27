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


def _get_dynamic_dispatch(prog_id: str) -> Any:
    from win32com.client import dynamic  # type: ignore[import-untyped]

    return dynamic.Dispatch(prog_id)


def _convert_with_iconverter(input_path: str, output_path: str) -> tuple[bool, str]:
    try:
        app7 = _get_dynamic_dispatch("Kompas.Application.7")
        app7.Visible = True

        base_path = str(getattr(app7.Application, "Path", "") or getattr(app7, "Path", "") or "").strip()
        if not base_path:
            return False, "API7 path is empty"

        converter_path = None
        for sub in ("Bin", "Libs"):
            candidate = Path(base_path) / sub / "PdfConverter.rtp"
            if candidate.exists():
                converter_path = candidate
                break
        if converter_path is None:
            return False, "PdfConverter.rtp not found"

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
            ok = bool(doc7.SaveAs(output_path))
            if ok and Path(output_path).exists() and Path(output_path).stat().st_size > 100:
                return True, "SaveAs"
            try:
                ok3 = bool(doc7.SaveAs3(output_path, "", False, 0))
            except Exception:
                ok3 = False
            if ok3 and Path(output_path).exists() and Path(output_path).stat().st_size > 100:
                return True, "SaveAs3"
            return False, "SaveAs/SaveAs3 failed"
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
    return jsonify({"status": "running", "service": "kompas-pdf"})


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

        attempts = (
            _convert_with_iconverter,
            _convert_with_api5_save_to_pdf,
            _convert_with_saveas,
        )
        last_error = "Unknown conversion error"
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
            last_error = method_or_error

        return jsonify({"success": False, "error": last_error}), 500
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
    port = int(os.environ.get("NORDFOX_PDF_SERVICE_PORT", "5000"))
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
