"""
Каскадное обновление переменных проекта NordFox в КОМПАС-3D:
сборка -> детали -> финальная пересборка сборки.
"""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Dict, List, Tuple

from .kompas_connector import KompasConnector
from .models import KompasDocumentInfo


logger = logging.getLogger("VariablesUpdater")


def _is_instance_variable_name(name: str) -> bool:
    if not name.startswith("v") or "_" not in name:
        return False
    prefix = name.split("_", 1)[0]
    return len(prefix) > 1 and prefix[1:].isdigit()


def _wait_short(seconds: float) -> None:
    time.sleep(seconds)


def _rebuild_3d_document(i_doc3d, i_part, cycles: int = 1) -> None:
    for _ in range(max(cycles, 1)):
        i_part.Update()
        _wait_short(0.2)
        i_part.RebuildModel()
        _wait_short(0.3)
        i_doc3d.RebuildDocument()
        _wait_short(0.3)


def _collect_main_table_variables(var_collection) -> Dict[str, float]:
    values: Dict[str, float] = {}
    for idx in range(300):
        try:
            var = var_collection.GetByIndex(idx)
            if not var:
                break
            name = getattr(var, "name", None)
            if not name or _is_instance_variable_name(name):
                continue
            val = getattr(var, "value", None)
            if val is None:
                continue
            values[name] = float(val)
        except Exception:
            break
    return values


def _update_assembly_variables(
    conn: KompasConnector,
    assembly_path: Path,
    changed_values: Dict[str, float],
) -> Tuple[int, List[str], List[Tuple[str, float | None, float]], Dict[str, float]]:
    updated_count = 0
    errors: List[str] = []
    updated_vars: List[Tuple[str, float | None, float]] = []
    assembly_values: Dict[str, float] = {}

    if not conn.open_document(str(assembly_path)):
        msg = f"Не удалось открыть сборку: {assembly_path}"
        logger.error(msg)
        errors.append(msg)
        return updated_count, errors, updated_vars, assembly_values

    api5 = conn.get_api5()
    api7 = conn.get_api7()
    if api5 is None or api7 is None:
        msg = "API5/API7 недоступны для обновления переменных"
        logger.error(msg)
        errors.append(msg)
        conn.close_active_document(save=False)
        return updated_count, errors, updated_vars, assembly_values

    try:
        i_doc3d = api5.ActiveDocument3D
        if not i_doc3d:
            msg = "ActiveDocument3D не найден"
            logger.error(msg)
            errors.append(msg)
            conn.close_active_document(save=False)
            return updated_count, errors, updated_vars, assembly_values

        i_part = i_doc3d.GetPart(-1)
        var_collection = i_part.VariableCollection()
    except Exception as exc:  # pragma: no cover
        msg = f"Ошибка доступа к VariableCollection: {exc}"
        logger.error(msg)
        errors.append(msg)
        conn.close_active_document(save=False)
        return updated_count, errors, updated_vars, assembly_values

    logger.info("  Сборка: входных переменных к обновлению: %d", len(changed_values))
    for name, new_val in changed_values.items():
        try:
            var = var_collection.GetByName(name)
            if not var:
                continue
            old_val = getattr(var, "value", None)
            # Пишем всегда, но логируем только если реально изменилось
            var.Expression = str(new_val)
            var.value = float(new_val)
            if hasattr(var, "External"):
                var.External = True
            if old_val is None or abs(float(old_val) - float(new_val)) > 1e-3:
                logger.info(f"{assembly_path.name}: {name}: {old_val} → {new_val}")
                old_as_float = float(old_val) if old_val is not None else None
                updated_vars.append((name, old_as_float, float(new_val)))
            updated_count += 1
        except Exception as exc:
            msg = f"{assembly_path.name}: ошибка обновления {name}: {exc}"
            logger.warning(msg)
            errors.append(msg)

    try:
        # Даем зависимым формулам в сборке пересчитаться.
        _rebuild_3d_document(i_doc3d, i_part, cycles=2)
        assembly_values = _collect_main_table_variables(var_collection)
        logger.info("  Сборка: прочитано переменных после пересчета: %d", len(assembly_values))
        for key in sorted(assembly_values.keys()):
            if key in changed_values:
                logger.info("    [DRV] %s = %s", key, assembly_values[key])
        api7.ActiveDocument.Save()
        _wait_short(0.2)
    except Exception as exc:  # pragma: no cover
        msg = f"{assembly_path.name}: ошибка пересборки/сохранения: {exc}"
        logger.warning(msg)
        errors.append(msg)

    conn.close_active_document(save=False)
    return updated_count, errors, updated_vars, assembly_values


def _is_formula_expression(expr: str) -> bool:
    return any(op in expr for op in ("+", "-", "*", "/", "if", "?"))


def _is_hyperlink_expression(expr: str) -> bool:
    return (":\\" in expr) or ("|" in expr)


def _cascade_update_part_variables(
    conn: KompasConnector,
    doc_path: Path,
    source_values: Dict[str, float],
) -> Tuple[int, List[str], List[Tuple[str, float | None, float]]]:
    updated = 0
    errors: List[str] = []
    updated_vars: List[Tuple[str, float | None, float]] = []

    if not conn.open_document(str(doc_path)):
        msg = f"Не удалось открыть документ: {doc_path}"
        logger.error(msg)
        errors.append(msg)
        return updated, errors, updated_vars

    api5 = conn.get_api5()
    api7 = conn.get_api7()
    if api5 is None or api7 is None:
        msg = "API5/API7 недоступны для обновления переменных"
        logger.error(msg)
        errors.append(msg)
        conn.close_active_document(save=False)
        return updated, errors, updated_vars

    try:
        i_doc3d = api5.ActiveDocument3D
        if not i_doc3d:
            msg = "ActiveDocument3D не найден"
            logger.error(msg)
            errors.append(msg)
            conn.close_active_document(save=False)
            return updated, errors, updated_vars

        i_part = i_doc3d.GetPart(-1)
        var_collection = i_part.VariableCollection()
    except Exception as exc:  # pragma: no cover
        msg = f"Ошибка доступа к VariableCollection: {exc}"
        logger.error(msg)
        errors.append(msg)
        conn.close_active_document(save=False)
        return updated, errors, updated_vars

    logger.info("  Деталь %s: сопоставление с переменными сборки (%d шт.)", doc_path.name, len(source_values))
    for name, new_val in source_values.items():
        if _is_instance_variable_name(name):
            continue
        try:
            var = var_collection.GetByName(name)
            if not var:
                continue
            old_val = getattr(var, "value", None)
            expr = str(getattr(var, "Expression", "") or "")

            if _is_hyperlink_expression(expr):
                logger.info("    %s: Expression гиперссылка -> число (%s)", name, new_val)
                var.Expression = str(new_val)
                var.value = float(new_val)
            elif _is_formula_expression(expr):
                logger.info("    %s: formula trigger через reset/restore Expression", name)
                saved_expr = expr
                var.Expression = ""
                _wait_short(0.05)
                var.Expression = saved_expr
            else:
                logger.info("    %s: прямое числовое обновление (%s)", name, new_val)
                var.Expression = str(new_val)
                var.value = float(new_val)

            if hasattr(var, "External"):
                var.External = True

            if old_val is None or abs(float(old_val) - float(new_val)) > 1e-3:
                logger.info(f"{doc_path.name}: {name}: {old_val} → {new_val}")
                old_as_float = float(old_val) if old_val is not None else None
                updated_vars.append((name, old_as_float, float(new_val)))
            updated += 1
        except Exception as exc:
            msg = f"{doc_path.name}: ошибка обновления {name}: {exc}"
            logger.warning(msg)
            errors.append(msg)

    if updated > 0:
        try:
            _rebuild_3d_document(i_doc3d, i_part, cycles=1)
            api7.ActiveDocument.Save()
            _wait_short(0.2)
        except Exception as exc:  # pragma: no cover
            msg = f"{doc_path.name}: ошибка пересборки/сохранения: {exc}"
            logger.warning(msg)
            errors.append(msg)

    # Закрываем документ без дополнительных диалогов
    conn.close_active_document(save=False)

    return updated, errors, updated_vars


def _final_rebuild_assembly(conn: KompasConnector, assembly_path: Path, cycles: int = 2) -> List[str]:
    errors: List[str] = []
    if not conn.open_document(str(assembly_path)):
        return [f"Не удалось открыть сборку для финальной пересборки: {assembly_path}"]

    api5 = conn.get_api5()
    api7 = conn.get_api7()
    if api5 is None or api7 is None:
        conn.close_active_document(save=False)
        return ["API5/API7 недоступны для финальной пересборки"]

    try:
        i_doc3d = api5.ActiveDocument3D
        i_part = i_doc3d.GetPart(-1)
        _rebuild_3d_document(i_doc3d, i_part, cycles=cycles)
        api7.ActiveDocument.Save()
        _wait_short(0.2)
    except Exception as exc:
        errors.append(f"{assembly_path.name}: ошибка финальной пересборки: {exc}")
    finally:
        conn.close_active_document(save=False)

    return errors


def update_project_variables(
    conn: KompasConnector,
    assembly_info: KompasDocumentInfo,
    documents: List[KompasDocumentInfo],
    new_values: Dict[str, float],
    drawing_comments: Dict[str, str] | None = None,
) -> Dict[str, object]:
    """
    Обновить переменные проекта:
    - new_values: имя переменной -> новое значение (из UI);
    - обновляются только те документы, где эти переменные присутствуют.

    Возвращает словарь с результатами:
        {
          "success": bool,
          "documents_updated": int,
          "variables_updated": int,
          "errors": [...]
        }
    """
    result = {
        "success": False,
        "documents_updated": 0,
        "variables_updated": 0,
        "errors": [],  # type: ignore[list-item]
    }

    all_errors: List[str] = []
    docs_updated = 0  # количество 3D документов, прошедших каскад
    vars_updated_total = 0
    per_doc_changes: Dict[Path, List[Tuple[str, float | None, float]]] = {}
    drawing_updates: Dict[Path, List[Tuple[str, str | None, str]]] = {}

    logger.info("=" * 60)
    logger.info("КАСКАДНОЕ ОБНОВЛЕНИЕ ПЕРЕМЕННЫХ")
    logger.info("=" * 60)
    logger.info("ЭТАП 1/3: обновление сборки")
    updated, errors, updated_vars, assembly_values = _update_assembly_variables(
        conn, assembly_info.path, new_values
    )
    all_errors.extend(errors)
    docs_updated += 1
    if updated > 0:
        vars_updated_total += updated
        per_doc_changes[assembly_info.path] = updated_vars

    logger.info("ЭТАП 2/3: каскад в детали")
    part_docs = [d for d in documents if d.doc_type == "part"]
    for doc in part_docs:
        updated_part, errors_part, updated_part_vars = _cascade_update_part_variables(
            conn, doc.path, assembly_values
        )
        all_errors.extend(errors_part)
        docs_updated += 1
        if updated_part > 0:
            vars_updated_total += updated_part
            per_doc_changes[doc.path] = updated_part_vars

    logger.info("ЭТАП 3/3: финальная пересборка сборки")
    all_errors.extend(_final_rebuild_assembly(conn, assembly_info.path, cycles=2))

    # Обновляем комментарии переменных чертежей, если заданы
    if drawing_comments:
        for doc in documents:
            if doc.doc_type != "drawing":
                continue
            path = doc.path
            try:
                if not conn.open_document(str(path)):
                    msg = f"Не удалось открыть чертёж для обновления комментариев: {path}"
                    logger.warning(msg)
                    all_errors.append(msg)
                    continue

                api5 = conn.get_api5()
                if api5 is None:
                    msg = "API5 недоступен для обновления переменных чертежа"
                    logger.error(msg)
                    all_errors.append(msg)
                    conn.close_active_document(save=False)
                    continue

                i_doc2d = getattr(api5, "ActiveDocument2D", None)
                if not i_doc2d:
                    msg = "ActiveDocument2D не найден для чертежа при обновлении комментариев"
                    logger.error(msg)
                    all_errors.append(msg)
                    conn.close_active_document(save=False)
                    continue

                var_collection = i_doc2d.VariableCollection()
                per_doc_list: List[Tuple[str, str | None, str]] = []

                for name, new_comment in drawing_comments.items():
                    try:
                        var = var_collection.GetByName(name)
                        if not var:
                            continue
                        old_comment = getattr(var, "Comment", None) if hasattr(var, "Comment") else None
                        if old_comment == new_comment:
                            continue
                        if hasattr(var, "Comment"):
                            var.Comment = new_comment
                        per_doc_list.append((name, old_comment, new_comment))
                        logger.info(f"{path.name}: комментарий {name}: {old_comment} → {new_comment}")
                    except Exception as exc:
                        msg = f"{path.name}: ошибка обновления комментария {name}: {exc}"
                        logger.warning(msg)
                        all_errors.append(msg)

                # Сохраняем и закрываем чертёж, если были изменения
                if per_doc_list:
                    try:
                        api7 = conn.get_api7()
                        if api7 is not None and getattr(api7, "ActiveDocument", None):
                            api7.ActiveDocument.Save()
                            time.sleep(0.3)
                    except Exception as exc:  # pragma: no cover
                        msg = f"{path.name}: ошибка сохранения чертежа после обновления комментариев: {exc}"
                        logger.warning(msg)
                        all_errors.append(msg)

                    drawing_updates[path] = per_doc_list

                conn.close_active_document(save=False)
            except Exception as exc:  # pragma: no cover
                msg = f"{path.name}: общая ошибка при обновлении комментариев: {exc}"
                logger.warning(msg)
                all_errors.append(msg)

    # Детализированный лог по каждому 3D-документу
    for doc_path, changes in per_doc_changes.items():
        logger.info("=" * 60)
        logger.info(f"Документ обновлён: {doc_path}")
        for name, old, new in changes:
            logger.info(f"  {name}: {old} → {new}")

    # Детализированный лог по комментариям чертежей
    for doc_path, changes in drawing_updates.items():
        logger.info("=" * 60)
        logger.info(f"Чертёж обновлён (комментарии): {doc_path}")
        for name, old, new in changes:
            logger.info(f"  {name}: {old} → {new}")

    logger.info("=" * 60)
    logger.info("ИТОГ ОБНОВЛЕНИЯ ПЕРЕМЕННЫХ")
    logger.info(f"  Документов обработано (3D): {docs_updated}")
    logger.info(f"  Переменных обновлено (3D): {vars_updated_total}")
    logger.info(f"  Чертежей с обновлёнными комментариями: {len(drawing_updates)}")
    for doc_path, changes in per_doc_changes.items():
        logger.info(f"    {doc_path.name}: переменных {len(changes)}")
    for doc_path, changes in drawing_updates.items():
        logger.info(f"    {doc_path.name}: комментариев {len(changes)}")
    logger.info("=" * 60)

    result["documents_updated"] = docs_updated
    result["variables_updated"] = vars_updated_total
    result["errors"] = all_errors
    result["success"] = len(all_errors) == 0

    return result

