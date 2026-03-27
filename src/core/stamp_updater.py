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
from typing import Any, Dict, List, Literal

from . import stamp_cells as SC
from .kompas_connector import KompasConnector


logger = logging.getLogger("StampUpdater")

SheetMode = Literal["none", "manual", "batch"]


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
    sheet_mode: SheetMode = "none",
    sheet_current: int | None = None,
    sheet_total: int | None = None,
) -> Dict[str, Any]:
    """
    Обновить штампы всех чертежей в указанной папке проекта.

    Поля соответствуют ячейкам КОМПАС (см. stamp_cells):
      - designation, name — основная надпись;
      - sheet_mode batch: нумерация листов 1…N по списку sort_drawings_for_sheet_numbering;
      - sheet_mode manual: одинаковые sheet_current / sheet_total для всех чертежей;
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
        total = len(sorted_for_sheets)
        for i, p in enumerate(sorted_for_sheets, start=1):
            sheet_index_map[p.resolve()] = (i, total)

    logger.info(f"Найдено чертежей для обработки: {len(all_drawings)}")

    api7 = conn.get_api7()
    if api7 is None:
        msg = "API7 недоступен для обновления штампов чертежей"
        logger.error(msg)
        result["errors"].append(msg)
        return result

    for drawing in all_drawings:
        try:
            logger.info("=" * 60)
            logger.info(drawing.name)
            logger.info("=" * 60)

            logger.info("  Открытие чертежа (API7)...")
            doc7 = api7.Documents.Open(str(drawing), False, False)
            if not doc7:
                msg = f"Не удалось открыть чертёж: {drawing}"
                logger.error(msg)
                result["drawings_failed"] += 1
                result["errors"].append(msg)
                continue

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
                continue

            cur_sheet: int | None = None
            tot_sheet: int | None = None
            if sheet_mode == "batch":
                pair = sheet_index_map.get(drawing.resolve())
                if pair:
                    cur_sheet, tot_sheet = pair
            elif sheet_mode == "manual":
                cur_sheet = sheet_current
                tot_sheet = sheet_total

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
                    cur_sheet is not None and tot_sheet is not None,
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
                            if cur_sheet is not None and tot_sheet is not None:
                                fields_to_update[SC.SHEET_CURRENT] = str(cur_sheet)
                                fields_to_update[SC.SHEET_TOTAL] = str(tot_sheet)

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

            if fields_updated > 0:
                result["drawings_updated"] += 1
                result["updated_files"].append(str(drawing))
                logger.info(f"  ✓ Штамп обновлён ({fields_updated} полей)")
            else:
                logger.info("  Штамп не изменялся (значения не заданы или совпадают)")

        except Exception as exc:  # pragma: no cover
            msg = f"Неожиданная ошибка при обработке {drawing}: {exc}"
            logger.warning(msg)
            result["drawings_failed"] += 1
            result["errors"].append(msg)

    logger.info("=" * 60)
    logger.info(
        f"ИТОГ ОБНОВЛЕНИЯ ШТАМПОВ: всего={result['drawings_total']}, "
        f"обновлено={result['drawings_updated']}, ошибок={result['drawings_failed']}"
    )
    logger.info("=" * 60)

    result["success"] = len(result["errors"]) == 0
    return result
