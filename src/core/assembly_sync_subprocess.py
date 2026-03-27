from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from pathlib import Path
from typing import Any


logger = logging.getLogger("AssemblySyncSubprocess")


def _read_prop(obj: object, low_name: str, pascal_name: str) -> object:
    val = None
    try:
        val = getattr(obj, low_name, None)
    except Exception:
        val = None
    if val in (None, ""):
        try:
            val = getattr(obj, pascal_name, None)
        except Exception:
            val = None
    return val


def _write_prop(obj: object, low_name: str, pascal_name: str, value: str) -> bool:
    try:
        setattr(obj, low_name, value)
        return True
    except Exception:
        pass
    try:
        setattr(obj, pascal_name, value)
        return True
    except Exception:
        return False


def _read_component_source_stem(comp: object) -> str:
    """
    Пытаемся получить stem исходного файла компонента (без расширения).
    Это fallback для случаев, когда mark/name вхождения не совпадают с деталями.
    """
    candidates = ("FileName", "filename", "fileName", "PathName", "pathName", "FullName", "fullName")
    for attr in candidates:
        try:
            raw = getattr(comp, attr, None)
        except Exception:
            raw = None
        if not raw:
            continue
        try:
            stem = Path(str(raw)).stem.strip().lower()
        except Exception:
            stem = ""
        if stem:
            return stem
    return ""


def run_sync(payload_path: Path) -> dict[str, Any]:
    from src.core.kompas_connector import KompasConnector

    payload = json.loads(payload_path.read_text(encoding="utf-8"))
    assembly_path = Path(str(payload.get("assembly_path", "")))
    items = list(payload.get("items") or [])
    result: dict[str, Any] = {
        "success": False,
        "documents_updated": 0,
        "fields_updated": 0,
        "errors": [],
    }
    errors: list[str] = []

    if not assembly_path.exists():
        result["errors"] = [f"Сборка не найдена: {assembly_path}"]
        return result

    by_marking: dict[str, dict[str, str]] = {}
    by_pair: dict[tuple[str, str], dict[str, str]] = {}
    by_file_stem: dict[str, dict[str, str]] = {}
    for item in items:
        old_mark = str(item.get("old_marking", "") or "").strip()
        old_name = str(item.get("old_name", "") or "").strip()
        new_mark = str(item.get("new_marking", "") or "").strip()
        new_name = str(item.get("new_name", "") or "").strip()
        file_stem = str(item.get("file_stem", "") or "").strip().lower()
        if not new_mark and not new_name:
            continue
        rec = {"marking": new_mark, "name": new_name}
        if old_mark:
            by_marking.setdefault(old_mark, rec)
        if old_mark or old_name:
            by_pair.setdefault((old_mark, old_name), rec)
        if file_stem:
            by_file_stem.setdefault(file_stem, rec)

    conn = KompasConnector()
    if not conn.open_document(str(assembly_path)):
        result["errors"] = [f"Не удалось открыть сборку: {assembly_path}"]
        return result

    updated_components = 0
    updated_fields = 0
    visited_components = 0
    matched_components = 0
    save_on_close = False
    try:
        api5 = conn.get_api5()
        api7 = conn.get_api7()
        if api5 is None or api7 is None:
            return {"success": False, "documents_updated": 0, "fields_updated": 0, "errors": ["API5/API7 недоступны"]}
        i_doc3d = getattr(api5, "ActiveDocument3D", None)
        if not i_doc3d:
            return {"success": False, "documents_updated": 0, "fields_updated": 0, "errors": ["ActiveDocument3D не найден"]}
        i_asm = i_doc3d.GetPart(-1)
        if not i_asm:
            return {"success": False, "documents_updated": 0, "fields_updated": 0, "errors": ["GetPart(-1) вернул пусто"]}

        for idx in range(200):
            try:
                comp = i_asm.GetPart(idx)
            except Exception:
                break
            if not comp:
                break
            visited_components += 1
            cur_mark = str(_read_prop(comp, "marking", "Marking") or "").strip()
            cur_name = str(_read_prop(comp, "name", "Name") or "").strip()
            item = by_marking.get(cur_mark) or by_pair.get((cur_mark, cur_name))
            if not item:
                src_stem = _read_component_source_stem(comp)
                if src_stem:
                    item = by_file_stem.get(src_stem)
            if not item:
                continue
            matched_components += 1

            changed = 0
            new_mark = str(item.get("marking", "") or "").strip()
            new_name = str(item.get("name", "") or "").strip()
            if new_mark:
                if _write_prop(comp, "marking", "Marking", new_mark):
                    changed += 1
                else:
                    errors.append(f"comp[{idx}]: не удалось записать marking")
            if new_name:
                if _write_prop(comp, "name", "Name", new_name):
                    changed += 1
                else:
                    errors.append(f"comp[{idx}]: не удалось записать name")
            if changed > 0:
                try:
                    comp.Update()
                except Exception:
                    pass
                try:
                    comp.RebuildModel()
                except Exception:
                    pass
                updated_components += 1
                updated_fields += changed
                save_on_close = True

        if save_on_close:
            try:
                i_asm.Update()
            except Exception:
                pass
            try:
                i_doc3d.RebuildDocument()
            except Exception:
                pass
            time.sleep(0.2)
            try:
                api7.ActiveDocument.Save()
            except Exception as exc:
                errors.append(f"Save error: {exc}")
            time.sleep(0.2)
    except Exception as exc:
        errors.append(f"Синхронизация компонентов сборки: {exc}")
    finally:
        conn.close_active_document(save=save_on_close)

    result["success"] = len(errors) == 0
    result["documents_updated"] = updated_components
    result["fields_updated"] = updated_fields
    result["components_visited"] = visited_components
    result["components_matched"] = matched_components
    result["errors"] = errors
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--payload", required=True)
    parser.add_argument("--result", required=True)
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%H:%M:%S",
    )
    repo_root = Path(__file__).resolve().parents[2]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))

    payload_path = Path(args.payload)
    result_path = Path(args.result)
    try:
        result = run_sync(payload_path)
    except Exception as exc:
        result = {
            "success": False,
            "documents_updated": 0,
            "fields_updated": 0,
            "errors": [f"fatal: {exc}"],
        }
    result_path.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
    return 0 if result.get("success") else 1


if __name__ == "__main__":
    raise SystemExit(main())
