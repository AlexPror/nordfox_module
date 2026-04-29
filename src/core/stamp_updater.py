from __future__ import annotations

"""
Обновление штампов чертежей в проекте NordFox по образцу kompas3d_project_manager.

Логика:
- найти все чертежи (*.cdw) в папке проекта (сначала в корне, при необходимости рекурсивно);
- исключить развертки (имя содержит 'развертка' или 'razvertka');
- для каждого чертежа открыть документ через API7, найти первый лист и его штамп;
- обновить только те поля, значения для которых заданы (не пустые);
- сохранить чертёж.

Номера ячеек для ГОСТ 2.104 — см. модуль stamp_cells.
"""

import logging
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional, Sequence, Tuple

from . import stamp_cells as SC
from .kompas_connector import KompasConnector


logger = logging.getLogger("StampUpdater")

SheetMode = Literal["none", "manual", "batch"]

# KompasAPIObjectTypeEnum.ksObjectDrawingDocumentSettings — IDrawingDocumentSettings
# (у IKompasDocument.DocumentSettings только IDocumentSettings, без SetSheetAutoNumber в dispinterface)
_KS_OBJECT_DRAWING_DOCUMENT_SETTINGS = 10042

# LIBID библиотеки типов KOMPAS-3D API7 (gencache: IDrawingDocumentSettings с корректным vtable)
_KOMPAS_API7_TLB_GUID = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"

# RPC_S_SERVER_UNAVAILABLE — процесс КОМПАС закрыт или COM-сессия оборвалась при длинной пакетной обработке
_RPC_UNAVAILABLE_HRESULT = -2147023174
_RPC_UNAVAILABLE_HRESULT_U = 0x800706BA


def _com_invoke_target(obj: Any) -> Any:
    """
    Дойти до объекта, у которого корректно вызывается COM Invoke.

    GetInterface / dynamic Dispatch иногда возвращает цепочку CDispatch → CDispatch → PyIDispatch.
    У gen_py DispatchBaseClass ожидается _oleobj_ : PyIDispatch; иначе Invoke получает обёртку и
    падает «The Python instance can not be converted to a COM object» даже для int.
    """
    cur: Any = obj
    for _ in range(20):
        if cur is None:
            return None
        nxt = getattr(cur, "_oleobj_", None)
        if nxt is None:
            return cur
        if nxt is cur:
            return cur
        cur = nxt
    return cur


def _unwrap_variant_value(v: Any) -> Any:
    """Вынуть .value у pywin32 VARIANT (иначе bool(VARIANT) может врать)."""
    for _ in range(3):
        if v is None:
            return None
        val = getattr(v, "value", v)
        if val is v:
            return v
        v = val
    return v


def _com_scalar(v: Any) -> Any:
    """Один скаляр из VARIANT / tuple из COM."""
    v = _unwrap_variant_value(v)
    item = getattr(v, "item", None)
    if callable(item):
        try:
            v = item()
        except Exception:
            pass
    if isinstance(v, (list, tuple)) and len(v) == 1:
        return _com_scalar(v[0])
    return v


def _sheet_autonumber_enabled(raw: Any) -> bool:
    """
    True = в КОМПАС включена автонумерация листов.
    VARIANT_BOOL: 0 = выкл, -1 (или любой ненулевой int) часто = вкл.
    """
    v = _unwrap_variant_value(raw)
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return int(v) != 0
    if isinstance(v, str):
        s = v.strip().lower()
        if s in ("0", "", "false", "нет", "off", "no"):
            return False
        if s in ("-1", "1", "true", "yes", "on"):
            return True
        try:
            return int(s, 10) != 0
        except ValueError:
            return len(s) > 0
    return bool(v)


def _drawing_document_settings(doc_2d: Any) -> tuple[Any | None, str | None]:
    """
    Интерфейс настроек чертежа (IDrawingDocumentSettings) для COM API7.

    GetInterface(10042) даёт указатель без нужного dispinterface в dynamic dispatch —
    обязательная обёртка: mod.IDrawingDocumentSettings(raw).
    """
    errors: list[str] = []
    try:
        from win32com.client import gencache

        mod = gencache.EnsureModule(_KOMPAS_API7_TLB_GUID, 0, 1, 0)
    except Exception as exc:
        return None, f"type library API7 ({_KOMPAS_API7_TLB_GUID}): {exc}"

    raw: Any | None = None
    if hasattr(doc_2d, "GetInterface"):
        try:
            raw = doc_2d.GetInterface(_KS_OBJECT_DRAWING_DOCUMENT_SETTINGS)
        except Exception as exc:
            errors.append(f"GetInterface({_KS_OBJECT_DRAWING_DOCUMENT_SETTINGS}): {exc}")

    if raw is None:
        try:
            doc1 = mod.IKompasDocument1(doc_2d)
            raw = doc1.GetInterface(_KS_OBJECT_DRAWING_DOCUMENT_SETTINGS)
        except Exception as exc:
            errors.append(f"IKompasDocument1.GetInterface: {exc}")

    if raw is None:
        try:
            raw = doc_2d.DocumentSettings
        except Exception as exc:
            errors.append(f"DocumentSettings: {exc}")

    if raw is None:
        return None, "; ".join(errors) if errors else "не удалось получить объект настроек чертежа"

    inner = _com_invoke_target(raw)
    if inner is None:
        return None, "; ".join(errors) if errors else "пустой указатель настроек чертежа"

    try:
        return mod.IDrawingDocumentSettings(inner), None
    except Exception as exc:
        try:
            from win32com.client import Dispatch

            return Dispatch(inner), None
        except Exception as exc2:
            tail = "; ".join(errors)
            return None, (
                f"IDrawingDocumentSettings(inner): {exc}; Dispatch(inner): {exc2}"
                + (f"; {tail}" if tail else "")
            )


# Как записали SheetAutoNumber в этом процессе (ускоряет следующие файлы)
_sheet_autonumber_put_strategy: str | None = None


def _set_sheet_auto_number(ds: Any, value: bool) -> None:
    """
    Записать признак автонумерации листов (TLB: свойство SheetAutoNumber, dispid 1).

    DispatchBaseClass.__setattr__ передаёт значение в PyIDispatch.Invoke; Python bool
    даёт «can not be converted to a COM object» — нужны VARIANT_BOOL: 0 / -1.
    Сначала Invoke на развёрнутом PyIDispatch (обход битой цепочки _oleobj_).
    """
    global _sheet_autonumber_put_strategy
    vb_false, vb_true = 0, -1
    vb = vb_true if value else vb_false

    put_map = getattr(ds, "_prop_map_put_", None) or getattr(type(ds), "_prop_map_put_", None)

    def invoke_core_put() -> None:
        import pythoncom

        core = _com_invoke_target(ds)
        if core is None:
            raise RuntimeError("пустой COM-указатель")
        if put_map and "SheetAutoNumber" in put_map:
            args, tail = put_map["SheetAutoNumber"]
            core.Invoke(*(args + (vb,) + tail))
        else:
            core.Invoke(1, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, vb)

    strats: list[tuple[str, Any]] = [("PyIDispatch.Invoke через развёрнутый core", invoke_core_put)]
    strats.extend(
        [
            ("VARIANT_BOOL int через свойство", lambda: setattr(ds, "SheetAutoNumber", vb)),
            ("python bool через свойство", lambda: setattr(ds, "SheetAutoNumber", bool(value))),
        ]
    )

    try:
        import pythoncom
        from win32com.client import VARIANT

        strats.append(
            (
                "VARIANT(VT_BOOL)",
                lambda: setattr(
                    ds,
                    "SheetAutoNumber",
                    VARIANT(pythoncom.VT_BOOL, bool(value)),
                ),
            )
        )
    except Exception:
        pass

    def invoke_via_ds_oleobj() -> None:
        if not put_map or "SheetAutoNumber" not in put_map:
            raise KeyError("нет _prop_map_put_['SheetAutoNumber']")
        args, tail = put_map["SheetAutoNumber"]
        getattr(ds, "_oleobj_").Invoke(*(args + (vb,) + tail))

    strats.append(("Invoke через ds._oleobj_", invoke_via_ds_oleobj))

    errs: list[str] = []
    preferred = _sheet_autonumber_put_strategy
    ordered = strats
    if preferred:
        ordered = [(n, f) for n, f in strats if n == preferred] + [(n, f) for n, f in strats if n != preferred]

    for name, fn in ordered:
        try:
            fn()
            _sheet_autonumber_put_strategy = name
            return
        except Exception as exc:
            errs.append(f"{name}: {exc!s}")

    raise RuntimeError("; ".join(errs))


def _get_sheet_auto_number(ds: Any) -> Any:
    """Сырое значение SheetAutoNumber (для проверки см. _sheet_autonumber_enabled)."""
    errs: list[str] = []
    import pythoncom

    core = _com_invoke_target(ds)
    if core is not None:
        try:
            return _com_scalar(
                core.InvokeTypes(1, 0, pythoncom.DISPATCH_PROPERTYGET, (11, 0), ())
            )
        except Exception as exc:
            errs.append(f"InvokeTypes PROPERTYGET: {exc}")
        try:
            return _com_scalar(core.Invoke(1, 0, pythoncom.DISPATCH_PROPERTYGET, 1))
        except Exception as exc:
            errs.append(f"Invoke PROPERTYGET: {exc}")
    try:
        return _com_scalar(ds.SheetAutoNumber)
    except Exception as exc:
        errs.append(f"SheetAutoNumber read: {exc}")
    try:
        return _com_scalar(ds.IsSheetAutoNumber())
    except Exception as exc:
        errs.append(f"IsSheetAutoNumber: {exc}")
    try:
        g = getattr(ds, "_prop_map_get_", None) or getattr(type(ds), "_prop_map_get_", None)
        if g and "SheetAutoNumber" in g:
            return _com_scalar(ds._ApplyTypes_(*g["SheetAutoNumber"]))
    except Exception as exc:
        errs.append(f"_ApplyTypes_ GET: {exc}")
    raise RuntimeError("; ".join(errs))


def _idrawing_property_put(ds: Any, dispid: int, value: Any) -> None:
    """Запись свойства IDrawingDocumentSettings по DISPID (gen_py: 1 SheetAutoNumber, 2 First, 3 AutoCount, 4 SheetsCount)."""
    import pythoncom

    put_map = getattr(ds, "_prop_map_put_", None) or getattr(type(ds), "_prop_map_put_", None)
    core = _com_invoke_target(ds)
    if core is None:
        raise RuntimeError("пустой COM-указатель для IDrawingDocumentSettings")
    did = int(dispid)
    if put_map:
        for _name, spec in put_map.items():
            args, tail = spec
            if args and int(args[0]) == did:
                core.Invoke(*(args + (value,) + tail))
                return
    core.Invoke(did, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, value)


def _ensure_sheet_auto_number_disabled(
    doc_2d: Any,
    *,
    context: str = "",
    predefined_sheets_total: int | None = None,
    sheet_first_number: int | None = None,
    disable_auto_sheet_count: bool = False,
) -> tuple[bool, str | None]:
    """
    Подготовить документ к ручным полям «лист / листов» в штампе (API7).

    - SheetAutoNumber = выкл (иначе КОМПАС перезаписывает ячейки штампа).
    - SheetAutoCount = выкл и SheetsCount = N — иначе «листов» подменяется на число
      листов в одном файле (часто 1), даже если в штамп записали другое значение.
    - Для комплекта из отдельных .cdw: на каждый файл нужно выставить SheetFirstNumber
      равным номеру листа в комплекте (иначе при одном листе в документе КОМПАС
      показывает всегда «лист 1» из N при обновлении штампа из настроек).
    """
    prefix = f"{context}: " if context else ""
    ds, how_err = _drawing_document_settings(doc_2d)
    if ds is None:
        return False, f"{prefix}{how_err or 'не удалось получить настройки чертежа'}"

    try:
        _set_sheet_auto_number(ds, False)
    except Exception as exc:
        return False, f"{prefix}отключение SheetAutoNumber: {exc}"

    try:
        if predefined_sheets_total is not None:
            t = max(1, int(predefined_sheets_total))
            first = 1
            if sheet_first_number is not None:
                first = max(1, min(t, int(sheet_first_number)))
            _idrawing_property_put(ds, 3, 0)
            _idrawing_property_put(ds, 2, first)
            _idrawing_property_put(ds, 4, t)
        elif disable_auto_sheet_count:
            _idrawing_property_put(ds, 3, 0)
    except Exception as exc:
        return False, f"{prefix}настройки количества листов (SheetAutoCount/SheetsCount): {exc}"

    time.sleep(0.1)
    ds2, _ = _drawing_document_settings(doc_2d)
    read_ds = ds2 if ds2 is not None else ds

    raw_after: Any = None
    for attempt in (1, 2):
        try:
            raw_after = _get_sheet_auto_number(read_ds)
        except Exception as exc:
            return False, f"{prefix}чтение SheetAutoNumber: {exc}"
        if not _sheet_autonumber_enabled(raw_after):
            return True, None
        if attempt == 1:
            logger.warning(
                "%sчтение SheetAutoNumber после выключения даёт «вкл» (%r), повторная запись",
                prefix,
                raw_after,
            )
            try:
                _set_sheet_auto_number(read_ds, False)
            except Exception as exc:
                return False, f"{prefix}повтор отключения SheetAutoNumber: {exc}"
            time.sleep(0.1)
            ds3, _ = _drawing_document_settings(doc_2d)
            if ds3 is not None:
                read_ds = ds3

    logger.warning(
        "%sпо чтению API SheetAutoNumber всё ещё «вкл» (%r). "
        "PUT 0 выполнен без ошибки — для КОМПАС чтение иногда не совпадает с фактом до Save; "
        "продолжаем обработку штампа.",
        prefix,
        raw_after,
    )
    return True, None


def _is_rpc_unavailable(exc: BaseException) -> bool:
    """Обрыв связи с сервером автоматизации (часто после многих Open/Close подряд)."""
    try:
        from pywintypes import com_error  # type: ignore[import-untyped]

        if isinstance(exc, com_error) and exc.args:
            hr = int(exc.args[0])
            if hr == _RPC_UNAVAILABLE_HRESULT:
                return True
            if (hr & 0xFFFFFFFF) == _RPC_UNAVAILABLE_HRESULT_U:
                return True
    except Exception:
        pass
    text = str(exc).lower()
    return (
        "rpc" in text
        or "0x800706ba" in text
        or "-2147023174" in text
        or "сервер rpc недоступен" in text
    )


def collect_drawings_for_stamps(project_root: Path) -> List[Path]:
    """
    Список чертежей для обработки (без развёрток, без временных ~$).
    """
    project_root = Path(project_root).resolve()
    all_drawings = list(project_root.glob("*.cdw"))
    if not all_drawings:
        all_drawings = list(project_root.rglob("*.cdw"))
        all_drawings = [d for d in all_drawings if not d.name.startswith("~$")]

    out: List[Path] = []
    for drawing in all_drawings:
        name_lower = drawing.name.lower()
        if "развертка" in name_lower or "razvertka" in name_lower:
            continue
        out.append(drawing)
    return out


def sort_drawings_for_sheet_numbering(drawings: List[Path]) -> List[Path]:
    """
    Сортировка: сначала по номеру из «(лист N)» в имени файла, иначе по имени.
    """

    def sort_key(p: Path) -> tuple:
        m = re.search(r"\(лист\s*(\d+)\)", p.name, re.IGNORECASE)
        if m:
            return (0, int(m.group(1)), p.name.lower())
        return (1, p.name.lower())

    return sorted(drawings, key=sort_key)


def is_drawing_title_sheet(path: Path) -> bool:
    """
    Титульный лист не участвует в нумерации листов комплекта: в штампе поля листа очищаются.

    Определение — по имени файла (без учёта регистра).
    """
    stem_l = Path(path).stem.lower()
    name_l = Path(path).name.lower()
    return "титул" in stem_l or "титул" in name_l or "title" in stem_l


def is_drawing_node_sheet(path: Path) -> bool:
    """Чертёж узла(ов): в штампе в ячейке материала (1) перечислены номера узлов на листе — для перечня."""
    s = (Path(path).stem + " " + Path(path).name).lower()
    return "узел" in s or "узлы" in s or "uzel" in s or "uzly" in s


def read_stamp_cell_str(
    conn: KompasConnector,
    cdw_path: Path,
    cell_id: int,
    *,
    layout_sheet_index: int = 0,
) -> Optional[str]:
    """
    Прочитать текст одной ячейки штампа первого листа оформления. Документ не сохраняется.
    """
    path = Path(cdw_path).resolve()
    if not path.is_file() or path.suffix.lower() != ".cdw":
        return None
    if not conn.connect(force_reconnect=False):
        return None
    api7 = conn.get_api7()
    if api7 is None:
        return None
    doc7 = None
    try:
        doc7 = api7.Documents.Open(str(path), False, False)
        if not doc7:
            return None
        time.sleep(0.4)
        doc2d = api7.ActiveDocument
        if not doc2d:
            return None
        ok_auto, _ = _ensure_sheet_auto_number_disabled(doc2d, context=path.name)
        if not ok_auto:
            logger.debug("read_stamp_cell_str: автонумерация листов для %s", path.name)

        layout_sheets = doc2d.LayoutSheets
        if not layout_sheets or layout_sheets.Count == 0:
            return None
        if layout_sheet_index < 0 or layout_sheet_index >= int(layout_sheets.Count):
            return None
        sheet = layout_sheets.Item(layout_sheet_index)
        if not sheet:
            return None
        stamp = sheet.Stamp
        if not stamp:
            return None
        try:
            text_item = stamp.Text(int(cell_id))
            val = str(text_item.Str or "").strip()
            return val or None
        except Exception:
            return None
    except Exception as exc:
        logger.debug("read_stamp_cell_str %s: %s", path, exc)
        return None
    finally:
        if doc7 is not None:
            try:
                doc7.Close(False)
            except Exception:
                pass
            time.sleep(0.15)


def scan_stamp_cells_non_empty(
    conn: KompasConnector,
    cdw_path: Path,
    *,
    cell_index_min: int = 1,
    cell_index_max: int = 220,
    layout_sheet_index: int = 0,
) -> Dict[str, Any]:
    """
    Один чертёж: прочитать штамп первого (или выбранного) листа оформления.
    Возвращает список пар (индекс ячейки, текст) для непустых значений.
    Документ не сохраняется.
    """
    path = Path(cdw_path).resolve()
    out: Dict[str, Any] = {
        "success": False,
        "error": None,
        "path": str(path),
        "cells": [],
        "layout_sheet_index": layout_sheet_index,
    }

    if not path.is_file() or path.suffix.lower() != ".cdw":
        out["error"] = f"Нужен файл .cdw: {path}"
        return out

    if not conn.connect(force_reconnect=False):
        out["error"] = "Не удалось подключиться к КОМПАС-3D."
        return out

    api7 = conn.get_api7()
    if api7 is None:
        out["error"] = "API7 недоступен."
        return out

    doc7 = None
    try:
        logger.info("Сканирование штампа: %s", path.name)
        doc7 = api7.Documents.Open(str(path), False, False)
        if not doc7:
            out["error"] = "Documents.Open вернул пусто."
            return out

        time.sleep(0.5)
        doc2d = api7.ActiveDocument
        if not doc2d:
            out["error"] = "ActiveDocument недоступен после открытия."
            return out

        ok_auto, auto_msg = _ensure_sheet_auto_number_disabled(doc2d, context=path.name)
        if not ok_auto:
            logger.warning("Сканирование штампа: автонумерация листов — %s", auto_msg)

        layout_sheets = doc2d.LayoutSheets
        if not layout_sheets or layout_sheets.Count == 0:
            out["error"] = "Листы оформления не найдены."
            return out

        if layout_sheet_index < 0 or layout_sheet_index >= int(layout_sheets.Count):
            out["error"] = f"Неверный индекс листа оформления: {layout_sheet_index} (всего {layout_sheets.Count})."
            return out

        sheet = layout_sheets.Item(layout_sheet_index)
        if not sheet:
            out["error"] = f"Лист оформления {layout_sheet_index} не получен."
            return out

        stamp = sheet.Stamp
        if not stamp:
            out["error"] = "Штамп на листе не найден."
            return out

        cells: List[Tuple[int, str]] = []
        lo = max(1, int(cell_index_min))
        hi = min(500, int(cell_index_max))
        for i in range(lo, hi + 1):
            try:
                text_item = stamp.Text(i)
                val = ""
                try:
                    val = str(text_item.Str or "").strip()
                except Exception:
                    val = ""
                if val:
                    cells.append((i, val))
                    logger.info("  штамп ячейка %s: %s", i, val)
            except Exception:
                continue

        out["cells"] = cells
        out["success"] = True
        logger.info("Сканирование штампа: непустых ячеек %s (диапазон %s…%s)", len(cells), lo, hi)
        return out
    except Exception as exc:
        logger.exception("Сканирование штампа: %s", exc)
        out["error"] = str(exc)
        return out
    finally:
        if doc7 is not None:
            try:
                doc7.Close(False)
            except Exception:
                pass
            time.sleep(0.2)


def update_all_drawing_stamps(
    conn: KompasConnector,
    project_root: Path,
    *,
    developer: str | None = None,
    checker: str | None = None,
    organization: str | None = None,
    material: str | None = None,
    tech_control: str | None = None,
    norm_control: str | None = None,
    approved: str | None = None,
    date: str | None = None,
    role_dates: Dict[int, str] | None = None,
    designation: str | None = None,
    name: str | None = None,
    document_letter: str | None = None,
    sheet_mode: SheetMode = "none",
    sheet_current: int | None = None,
    sheet_total: int | None = None,
    sheet_batch_paths: Sequence[Path] | None = None,
) -> Dict[str, Any]:
    """
    Обновить штампы всех чертежей в указанной папке проекта.

    Поля соответствуют ячейкам КОМПАС (см. stamp_cells):
      - designation, name — основная надпись;
      - sheet_mode batch: нумерация 1…M только для чертежей без титула; файлы с «титул»
        в имени — поля «лист / листов» очищаются; M = число таких чертежей;
      - sheet_mode manual: одинаковые sheet_current / sheet_total для всех чертежей;
      - document_letter: литера (ячейка DOCUMENT_LETTER2, см. stamp_cells);
      - developer: 110, checker: 111, …
    """
    result: Dict[str, Any] = {
        "success": False,
        "drawings_total": 0,
        "drawings_updated": 0,
        "drawings_failed": 0,
        "updated_files": [],
        "errors": [],
    }

    project_root = Path(project_root).resolve()
    if not project_root.exists() or not project_root.is_dir():
        msg = f"Папка проекта не найдена или не является папкой: {project_root}"
        logger.error(msg)
        result["errors"].append(msg)
        return result

    logger.info("=" * 60)
    logger.info("ОБНОВЛЕНИЕ ШТАМПОВ ЧЕРТЕЖЕЙ")
    logger.info("=" * 60)
    logger.info(f"Папка проекта: {project_root}")

    all_drawings = collect_drawings_for_stamps(project_root)
    result["drawings_total"] = len(all_drawings)

    if not all_drawings:
        msg = "Чертежи для обновления штампов не найдены."
        logger.warning(msg)
        result["errors"].append(msg)
        return result

    sorted_for_sheets = sort_drawings_for_sheet_numbering(all_drawings)
    sheet_index_map: Dict[Path, tuple[int, int]] = {}
    if sheet_mode == "batch" and sorted_for_sheets:
        all_resolved = {p.resolve() for p in all_drawings}
        if sheet_batch_paths:
            ordered = [Path(p).resolve() for p in sheet_batch_paths]
            if set(ordered) != all_resolved:
                msg = (
                    "Порядок листов: список путей не совпадает с набором чертежей в папке "
                    f"(ожидалось {len(all_drawings)} файлов)."
                )
                logger.error(msg)
                result["errors"].append(msg)
                return result
            numbered_only = [p for p in ordered if not is_drawing_title_sheet(p)]
            total = len(numbered_only)
            for i, p in enumerate(numbered_only, start=1):
                sheet_index_map[p] = (i, total)
        else:
            numbered_only = [p for p in sorted_for_sheets if not is_drawing_title_sheet(p)]
            total = len(numbered_only)
            for i, p in enumerate(numbered_only, start=1):
                sheet_index_map[p.resolve()] = (i, total)

    logger.info(f"Найдено чертежей для обработки: {len(all_drawings)}")

    max_retries_per_file = 3

    for drawing in all_drawings:
        last_error: str | None = None
        for attempt in range(max_retries_per_file):
            api7 = conn.get_api7()
            if api7 is None:
                msg = "API7 недоступен для обновления штампов чертежей"
                logger.error(msg)
                result["errors"].append(msg)
                return result

            try:
                logger.info("=" * 60)
                logger.info(drawing.name)
                if attempt > 0:
                    logger.info("  (повтор %s/%s после обрыва COM)", attempt + 1, max_retries_per_file)
                logger.info("=" * 60)

                logger.info("  Открытие чертежа (API7)...")
                doc7 = api7.Documents.Open(str(drawing), False, False)
                if not doc7:
                    msg = f"Не удалось открыть чертёж: {drawing}"
                    logger.error(msg)
                    result["drawings_failed"] += 1
                    result["errors"].append(msg)
                    last_error = None
                    break

                time.sleep(1.5)

                kompas_document_2d = api7.ActiveDocument
                if not kompas_document_2d:
                    msg = f"Документ не открыт (ActiveDocument = None): {drawing}"
                    logger.error(msg)
                    result["drawings_failed"] += 1
                    result["errors"].append(msg)
                    try:
                        doc7.Close(False)
                    except Exception:
                        pass
                    last_error = None
                    break

                cur_sheet: int | None = None
                tot_sheet: int | None = None
                clear_title_sheet_nums = False
                if sheet_mode == "batch":
                    if is_drawing_title_sheet(drawing):
                        clear_title_sheet_nums = True
                    else:
                        pair = sheet_index_map.get(drawing.resolve())
                        if pair:
                            cur_sheet, tot_sheet = pair
                elif sheet_mode == "manual":
                    cur_sheet = sheet_current
                    tot_sheet = sheet_total

                will_touch_sheet_fields = bool(
                    clear_title_sheet_nums or (cur_sheet is not None and tot_sheet is not None)
                )
                ok_auto, auto_msg = _ensure_sheet_auto_number_disabled(
                    kompas_document_2d,
                    context=drawing.name,
                    predefined_sheets_total=tot_sheet
                    if cur_sheet is not None and tot_sheet is not None
                    else None,
                    sheet_first_number=cur_sheet
                    if cur_sheet is not None and tot_sheet is not None
                    else None,
                    disable_auto_sheet_count=clear_title_sheet_nums,
                )
                if not ok_auto:
                    logger.warning("  Автонумерация листов: %s", auto_msg)
                    if will_touch_sheet_fields:
                        result["errors"].append(str(auto_msg))
                        result["drawings_failed"] += 1
                        try:
                            doc7.Close(False)
                        except Exception:
                            pass
                        time.sleep(0.35)
                        last_error = None
                        break

                any_field = any(
                    [
                        developer,
                        checker,
                        organization,
                        material,
                        tech_control,
                        norm_control,
                        approved,
                        date,
                        bool(role_dates),
                        designation,
                        name,
                        document_letter,
                        cur_sheet is not None and tot_sheet is not None,
                        clear_title_sheet_nums,
                    ]
                )
                fields_updated = 0

                if any_field:
                    layout_sheets = kompas_document_2d.LayoutSheets
                    if not layout_sheets or layout_sheets.Count == 0:
                        logger.warning("  Листы оформления не найдены")
                    else:
                        sheet = layout_sheets.Item(0)
                        if not sheet:
                            logger.warning("  Лист не найден")
                        else:
                            stamp = sheet.Stamp
                            if stamp:
                                fields_to_update: dict[int, str] = {}

                                if designation:
                                    fields_to_update[SC.DESIGNATION] = designation
                                if name:
                                    fields_to_update[SC.NAME] = name
                                if document_letter:
                                    fields_to_update[SC.DOCUMENT_LETTER2] = document_letter
                                if clear_title_sheet_nums:
                                    fields_to_update[SC.SHEET_CURRENT] = ""
                                    fields_to_update[SC.SHEET_TOTAL] = ""
                                    for _ph in SC.SHEET_PHANTOM_CELLS:
                                        fields_to_update[int(_ph)] = ""
                                    logger.info("  Титульный лист — очистка полей листа в штампе")
                                elif cur_sheet is not None and tot_sheet is not None:
                                    fields_to_update[SC.SHEET_CURRENT] = str(cur_sheet)
                                    fields_to_update[SC.SHEET_TOTAL] = str(tot_sheet)
                                    for _ph in SC.SHEET_PHANTOM_CELLS:
                                        fields_to_update[int(_ph)] = ""

                                if developer:
                                    fields_to_update[SC.DEVELOPER] = developer
                                if checker:
                                    fields_to_update[SC.CHECKER] = checker
                                if tech_control:
                                    fields_to_update[SC.TECH_CONTROL] = tech_control
                                if norm_control:
                                    fields_to_update[SC.NORM_CONTROL] = norm_control
                                if approved:
                                    fields_to_update[SC.APPROVED] = approved
                                if organization:
                                    fields_to_update[SC.ORGANIZATION] = organization

                                is_assembly = "сб" in drawing.name.lower() or "sb" in drawing.name.lower()
                                if material and not is_assembly:
                                    fields_to_update[SC.MATERIAL] = material
                                elif material and is_assembly:
                                    logger.info("  Материал пропущен (сборочный чертеж)")

                                if date:
                                    fields_to_update[SC.DATE_MAIN] = date
                                if role_dates:
                                    for cell_id, value in role_dates.items():
                                        if value:
                                            fields_to_update[int(cell_id)] = str(value)

                                for cell_id, value in fields_to_update.items():
                                    try:
                                        text_item = stamp.Text(cell_id)
                                        text_item.Str = str(value)
                                        logger.info(f"    Ячейка {cell_id}: {value}")
                                        fields_updated += 1
                                    except Exception as exc:
                                        logger.warning(f"    Ошибка записи ячейки {cell_id}: {exc}")

                                try:
                                    stamp.Update()
                                    sheet.Update()
                                except Exception as exc:
                                    logger.warning(f"    Ошибка обновления штампа/листа: {exc}")

                                if (
                                    SC.SHEET_CURRENT in fields_to_update
                                    or SC.SHEET_TOTAL in fields_to_update
                                ):
                                    time.sleep(0.25)
                                    for _cid in (SC.SHEET_CURRENT, SC.SHEET_TOTAL):
                                        if _cid in fields_to_update:
                                            try:
                                                stamp.Text(_cid).Str = str(fields_to_update[_cid])
                                                logger.info(
                                                    "    Повтор записи листа, ячейка %s: %s",
                                                    _cid,
                                                    fields_to_update[_cid],
                                                )
                                            except Exception as exc:
                                                logger.warning(
                                                    "    Повтор ячейки %s: %s", _cid, exc
                                                )
                                    try:
                                        stamp.Update()
                                        sheet.Update()
                                    except Exception as exc:
                                        logger.warning(
                                            "    Ошибка обновления штампа после повтора: %s", exc
                                        )

                                if SC.SHEET_CURRENT in fields_to_update or SC.SHEET_TOTAL in fields_to_update:
                                    for _vid in (SC.SHEET_CURRENT, SC.SHEET_TOTAL):
                                        try:
                                            got = str(stamp.Text(_vid).Str or "").strip()
                                            logger.info(
                                                "    После записи, чтение ячейки %s: %r",
                                                _vid,
                                                got,
                                            )
                                        except Exception as exc:
                                            logger.warning(
                                                "    Не удалось прочитать ячейку %s: %s", _vid, exc
                                            )

                                time.sleep(1.5)
                            else:
                                logger.warning("  Штамп не найден")

                try:
                    kompas_document_2d = api7.ActiveDocument
                    if kompas_document_2d:
                        kompas_document_2d.Save()
                        time.sleep(0.5)
                except Exception as exc:
                    logger.warning(f"  Ошибка сохранения чертежа {drawing.name}: {exc}")

                try:
                    doc7.Close(False)
                except Exception:
                    pass

                time.sleep(0.35)

                if fields_updated > 0:
                    result["drawings_updated"] += 1
                    result["updated_files"].append(str(drawing))
                    logger.info(f"  ✓ Штамп обновлён ({fields_updated} полей)")
                else:
                    logger.info("  Штамп не изменялся (значения не заданы или совпадают)")

                last_error = None
                break

            except Exception as exc:  # pragma: no cover
                last_error = f"Неожиданная ошибка при обработке {drawing}: {exc}"
                logger.warning(last_error)
                if _is_rpc_unavailable(exc) and attempt + 1 < max_retries_per_file:
                    logger.warning(
                        "  Обрыв COM/RPC — переподключение к КОМПАС и повтор (%s/%s)",
                        attempt + 2,
                        max_retries_per_file,
                    )
                    try:
                        conn.reconnect()
                    except Exception as reconnect_exc:
                        logger.error("  Переподключение не удалось: %s", reconnect_exc)
                    time.sleep(2.5)
                    continue
                result["drawings_failed"] += 1
                result["errors"].append(last_error)
                break

    logger.info("=" * 60)
    logger.info(
        f"ИТОГ ОБНОВЛЕНИЯ ШТАМПОВ: всего={result['drawings_total']}, "
        f"обновлено={result['drawings_updated']}, ошибок={result['drawings_failed']}"
    )
    logger.info("=" * 60)

    result["success"] = len(result["errors"]) == 0
    return result
