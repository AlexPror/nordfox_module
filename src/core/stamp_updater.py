from __future__ import annotations

"""
Обновление штампов чертежей в проекте NordFox по образцу kompas3d_project_manager.

Логика:
- найти все чертежи (*.cdw) в папке проекта (сначала в корне, при необходимости рекурсивно);
- исключить развертки (имя содержит 'развертка' или 'razvertka');
- для каждого чертежа открыть документ через API7, найти первый лист и его штамп;
- обновить только те поля, значения для которых заданы (не пустые);
- сохранить чертёж.
"""

import logging
import time
from pathlib import Path
from typing import Dict, Any

from .kompas_connector import KompasConnector


logger = logging.getLogger("StampUpdater")


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
    order_number: str | None = None,
) -> Dict[str, Any]:
    """
    Обновить штампы всех чертежей в указанной папке проекта.

    Поля соответствуют типовым ячейкам КОМПАС:
      - developer: 110
      - checker: 111
      - tech_control: 112
      - norm_control: 114
      - approved: 115
      - organization: 9
      - material: 3 (для несборочных чертежей)
      - date: 130 (общая дата разработки, если не передан role_dates)
      - role_dates: словарь {ячейка_даты: строка_даты} для проставления дат напротив ролей
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

    # Поиск чертежей: сначала в корне, затем при необходимости рекурсивно
    all_drawings = list(project_root.glob("*.cdw"))
    if not all_drawings:
        logger.info("Чертежи в корне не найдены, пробуем рекурсивный поиск...")
        all_drawings = list(project_root.rglob("*.cdw"))
        all_drawings = [d for d in all_drawings if not d.name.startswith("~$")]

    # Исключаем развертки
    drawings_to_process: list[Path] = []
    unfoldings_skipped = 0
    for drawing in all_drawings:
        name_lower = drawing.name.lower()
        if "развертка" in name_lower or "razvertka" in name_lower:
            unfoldings_skipped += 1
            continue
        drawings_to_process.append(drawing)

    all_drawings = drawings_to_process
    result["drawings_total"] = len(all_drawings)

    if unfoldings_skipped:
        logger.info(f"Пропущено разверток: {unfoldings_skipped}")

    logger.info(f"Найдено чертежей для обработки: {len(all_drawings)}")
    if not all_drawings:
        msg = "Чертежи для обновления штампов не найдены."
        logger.warning(msg)
        result["errors"].append(msg)
        return result

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

            # Если ни одно поле не задано — просто пересобираем/сохраняем чертёж
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

                            if developer:
                                fields_to_update[110] = developer
                            if checker:
                                fields_to_update[111] = checker
                            if tech_control:
                                fields_to_update[112] = tech_control
                            if norm_control:
                                fields_to_update[114] = norm_control
                            if approved:
                                fields_to_update[115] = approved
                            if organization:
                                fields_to_update[9] = organization

                            # Материал не заполняем для сборочных чертежей (СБ)
                            is_assembly = "сб" in drawing.name.lower() or "sb" in drawing.name.lower()
                            if material and not is_assembly:
                                fields_to_update[3] = material
                            elif material and is_assembly:
                                logger.info("  Материал пропущен (сборочный чертеж)")

                            if date:
                                fields_to_update[130] = date
                            if role_dates:
                                # Явные даты по ячейкам имеют приоритет (могут дополнять общую дату)
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

            # Сохраняем и закрываем документ
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

