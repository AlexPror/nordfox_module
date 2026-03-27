"""
Сканирование проекта NordFox в КОМПАС-3D:
- поиск сборки, деталей и чертежей в выбранной папке;
- чтение переменных из сборки и деталей (первый этап);
- подготовка структур для динамического UI.

Часть логики (получение VariableCollection, фильтрация переменных экземпляров)
адаптирована из CascadingVariablesUpdater из kompas3d_project_manager.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Tuple
import re

from .kompas_connector import KompasConnector
from .models import KompasDocumentInfo, KompasVariable


logger = logging.getLogger("VariablesScanner")


def _normalize_block_id(raw: str | None) -> str | None:
    if raw is None:
        return None
    normalized = raw.rstrip("_").strip()
    return normalized or None


def _is_instance_variable_name(name: str) -> bool:
    """
    Определить, является ли имя переменной переменной экземпляра (v1234_A3 и т.п.),
    которые не нужно копировать в файлы деталей.
    """
    if not name.startswith("v") or "_" not in name:
        return False
    prefix = name.split("_", 1)[0]
    return len(prefix) > 1 and prefix[1:].isdigit()


def _read_variables_from_assembly(conn: KompasConnector, asm_path: Path) -> Dict[str, KompasVariable]:
    """
    Прочитать переменные из сборки (.a3d) через API5.
    Возвращает словарь name -> KompasVariable.
    """
    if not conn.open_document(str(asm_path)):
        logger.error(f"Не удалось открыть сборку: {asm_path}")
        return {}

    api5 = conn.get_api5()
    if api5 is None:
        logger.error("API5 недоступен для чтения переменных сборки")
        return {}

    try:
        i_doc3d = api5.ActiveDocument3D
        if not i_doc3d:
            logger.error("ActiveDocument3D не найден для сборки")
            return {}

        i_part = i_doc3d.GetPart(-1)
        var_collection = i_part.VariableCollection()
    except Exception as exc:  # pragma: no cover
        logger.error(f"Ошибка получения VariableCollection для сборки: {exc}")
        return {}

    vars_dict: Dict[str, KompasVariable] = {}
    current_block: str | None = None

    # Паттерн пустой переменной-блока: Stoiki________, Kronshtein_MacFox____ и т.п.
    # Допускаем "_" внутри имени блока.
    header_pattern = re.compile(r"^([A-Za-zА-Яа-я0-9_]+)_+$")

    # Итерируем ограниченное количество переменных (как в оригинальном коде)
    for idx in range(200):
        try:
            var = var_collection.GetByIndex(idx)
            if not var:
                break
            name = getattr(var, "name", None)
            if not name:
                continue
            if _is_instance_variable_name(name):
                continue

            value = getattr(var, "value", None)
            expression = getattr(var, "Expression", None) if hasattr(var, "Expression") else None
            is_external = bool(getattr(var, "External", False))

            m = header_pattern.match(name)
            if m:
                # Это пустая переменная-заголовок блока (Stoiki____, Kronshtein_MacFox____ и т.п.)
                block_name = _normalize_block_id(m.group(1)) or m.group(1)
                current_block = block_name
                kv = KompasVariable(
                    name=name,
                    value=value,
                    original_value=value,
                    document_type="assembly",
                    document_path=asm_path,
                    comment=None,
                    expression=expression,
                    block_id=block_name,
                    is_block_header=True,
                    is_external=is_external,
                )
            else:
                block_id = _normalize_block_id(current_block)
                # Специальное правило: все переменные, в имени которых есть "MacFox",
                # попадают в блок Kronshtein_MacFox, даже если они физически в другом месте
                if "macfox" in name.lower():
                    block_id = "Kronshtein_MacFox"
                kv = KompasVariable(
                    name=name,
                    value=value,
                    original_value=value,
                    document_type="assembly",
                    document_path=asm_path,
                    comment=None,
                    expression=expression,
                    block_id=block_id,
                    is_block_header=False,
                    is_external=is_external,
                )

            # В индекс документа добавляем по имени; при совпадении имени из того же файла
            # считаем, что это одна и та же переменная.
            vars_dict[name] = kv

        except Exception:
            break

    logger.info(f"Переменные сборки: {len(vars_dict)}")
    return vars_dict


def _read_variables_from_part(conn: KompasConnector, part_path: Path) -> Dict[str, KompasVariable]:
    """
    Прочитать переменные из детали (.m3d) через API5.
    """
    if not conn.open_document(str(part_path)):
        logger.warning(f"Не удалось открыть деталь: {part_path}")
        return {}

    api5 = conn.get_api5()
    if api5 is None:
        logger.error("API5 недоступен для чтения переменных детали")
        return {}

    try:
        i_doc3d = api5.ActiveDocument3D
        if not i_doc3d:
            logger.error("ActiveDocument3D не найден для детали")
            return {}

        i_part = i_doc3d.GetPart(-1)
        var_collection = i_part.VariableCollection()
    except Exception as exc:  # pragma: no cover
        logger.error(f"Ошибка получения VariableCollection для детали: {exc}")
        return {}

    vars_dict: Dict[str, KompasVariable] = {}
    current_block: str | None = None
    header_pattern = re.compile(r"^([A-Za-zА-Яа-я0-9_]+)_+$")

    for idx in range(200):
        try:
            var = var_collection.GetByIndex(idx)
            if not var:
                break
            name = getattr(var, "name", None)
            if not name:
                continue
            if _is_instance_variable_name(name):
                continue

            value = getattr(var, "value", None)
            expression = getattr(var, "Expression", None) if hasattr(var, "Expression") else None
            is_external = bool(getattr(var, "External", False))

            m = header_pattern.match(name)
            if m:
                block_name = _normalize_block_id(m.group(1)) or m.group(1)
                current_block = block_name
                kv = KompasVariable(
                    name=name,
                    value=value,
                    original_value=value,
                    document_type="part",
                    document_path=part_path,
                    comment=None,
                    expression=expression,
                    block_id=block_name,
                    is_block_header=True,
                    is_external=is_external,
                )
            else:
                block_id = _normalize_block_id(current_block)
                if "macfox" in name.lower():
                    block_id = "Kronshtein_MacFox"
                kv = KompasVariable(
                    name=name,
                    value=value,
                    original_value=value,
                    document_type="part",
                    document_path=part_path,
                    comment=None,
                    expression=expression,
                    block_id=block_id,
                    is_block_header=False,
                    is_external=is_external,
                )

            vars_dict[name] = kv

        except Exception:
            break

    logger.info(f"Переменные детали {part_path.name}: {len(vars_dict)}")
    return vars_dict


def _read_marking_and_name(conn: KompasConnector, path: Path) -> Tuple[str | None, str | None]:
    """
    Прочитать обозначение (marking) и наименование (name) из 3D-документа (.a3d или .m3d).
    Используется для предварительного автозаполнения блоков в UI.
    """
    if not conn.open_document(str(path)):
        logger.warning(f"Не удалось открыть документ для чтения обозначения/имени: {path}")
        return None, None

    api5 = conn.get_api5()
    if api5 is None:
        logger.error("API5 недоступен для чтения обозначения/имени")
        conn.close_active_document(save=False)
        return None, None

    marking = None
    name = None

    try:
        i_doc3d = api5.ActiveDocument3D
        if not i_doc3d:
            logger.error("ActiveDocument3D не найден при чтении обозначения/имени")
        else:
            i_part = i_doc3d.GetPart(-1)
            # В рабочем образце используется lower-case; PascalCase оставляем как fallback.
            marking = getattr(i_part, "marking", None)
            if marking in (None, ""):
                marking = getattr(i_part, "Marking", None)
            name = getattr(i_part, "name", None)
            if name in (None, ""):
                name = getattr(i_part, "Name", None)
    except Exception as exc:  # pragma: no cover
        logger.error(f"Ошибка чтения обозначения/имени для {path}: {exc}")
    finally:
        conn.close_active_document(save=False)

    return marking, name


def _read_variables_from_drawing(conn: KompasConnector, drw_path: Path) -> Dict[str, KompasVariable]:
    """
    Прочитать переменные из чертежа (.cdw) через API.
    Основной интерес — значения и комментарии для блока \"Переменные чертежа\".
    """
    if not conn.open_document(str(drw_path)):
        logger.warning(f"Не удалось открыть чертёж: {drw_path}")
        return {}

    api5 = conn.get_api5()
    if api5 is None:
        logger.error("API5 недоступен для чтения переменных чертежа")
        conn.close_active_document(save=False)
        return {}

    vars_dict: Dict[str, KompasVariable] = {}
    current_block: str | None = None
    header_pattern = re.compile(r"^([A-Za-zА-Яа-я0-9_]+)_+$")

    try:
        i_doc2d = getattr(api5, "ActiveDocument2D", None)
        if not i_doc2d:
            logger.error("ActiveDocument2D не найден для чертежа")
            return {}
        # VariableCollection может быть недоступен для некоторых типов чертежей.
        try:
            var_collection = i_doc2d.VariableCollection()
        except Exception as exc:
            logger.warning(f"{drw_path.name}: VariableCollection недоступен, переменные чертежа будут пропущены: {exc}")
            return {}

        for idx in range(200):
            try:
                var = var_collection.GetByIndex(idx)
                if not var:
                    break
                name = getattr(var, "name", None)
                if not name:
                    continue

                value = getattr(var, "value", None)
                expression = getattr(var, "Expression", None) if hasattr(var, "Expression") else None
                comment = getattr(var, "Comment", None) if hasattr(var, "Comment") else None
                is_external = bool(getattr(var, "External", False))

                m = header_pattern.match(name)
                if m:
                    block_name = _normalize_block_id(m.group(1)) or m.group(1)
                    current_block = block_name
                    kv = KompasVariable(
                        name=name,
                        value=value,
                        original_value=value,
                        document_type="drawing",
                        document_path=drw_path,
                        comment=comment,
                        expression=expression,
                        block_id=block_name,
                        is_block_header=True,
                        is_external=is_external,
                    )
                else:
                    kv = KompasVariable(
                        name=name,
                        value=value,
                        original_value=value,
                        document_type="drawing",
                        document_path=drw_path,
                        comment=comment,
                        expression=expression,
                        block_id=_normalize_block_id(current_block),
                        is_block_header=False,
                        is_external=is_external,
                    )

                vars_dict[name] = kv
            except Exception:
                break
    finally:
        conn.close_active_document(save=False)

    logger.info(f"Переменные чертежа {drw_path.name}: {len(vars_dict)}")
    return vars_dict


def scan_project(root: Path, conn: KompasConnector) -> Tuple[KompasDocumentInfo, List[KompasDocumentInfo], Dict[str, KompasVariable]]:
    """
    Просканировать проект в указанной папке.

    Возвращает:
      - информацию о сборке (KompasDocumentInfo),
      - список других документов (детали и т.п.),
      - индекс переменных по имени (для динамического UI):
        name -> один "эталонный" KompasVariable (значение для поля ввода).
    """
    root = root.resolve()
    if not root.is_dir():
        raise ValueError(f"Папка проекта не найдена: {root}")

    # Ищем сборку рекурсивно в проекте (можно выбрать корневую папку проекта)
    asm_files = list(root.rglob("*.a3d"))
    if not asm_files:
        raise RuntimeError(f"В папке {root} не найдена сборка (*.a3d)")

    assembly_path = asm_files[0]
    logger.info(f"Найдена сборка: {assembly_path}")

    # Детали и чертежи — ТОЛЬКО рядом со сборкой (без обхода подпапок)
    assembly_dir = assembly_path.parent
    part_files = [p for p in assembly_dir.glob("*.m3d") if not p.stem.startswith("-")]
    drawing_files = list(assembly_dir.glob("*.cdw"))

    logger.info(f"Найдено деталей: {len(part_files)}, чертежей: {len(drawing_files)}")

    # Читаем переменные сборки и деталей
    asm_vars = _read_variables_from_assembly(conn, assembly_path)
    asm_marking, asm_name = _read_marking_and_name(conn, assembly_path)
    assembly_info = KompasDocumentInfo(
        path=assembly_path,
        doc_type="assembly",
        designation=asm_marking,
        name=asm_name,
        variables=asm_vars,
    )

    documents: List[KompasDocumentInfo] = []

    for part_path in part_files:
        vars_part = _read_variables_from_part(conn, part_path)
        part_marking, part_name = _read_marking_and_name(conn, part_path)
        documents.append(
            KompasDocumentInfo(
                path=part_path,
                doc_type="part",
                designation=part_marking,
                name=part_name,
                variables=vars_part,
            )
        )

    for drw_path in drawing_files:
        vars_drw = _read_variables_from_drawing(conn, drw_path)
        documents.append(
            KompasDocumentInfo(
                path=drw_path,
                doc_type="drawing",
                designation=None,
                name=None,
                variables=vars_drw,
            )
        )

    # Индекс переменных по имени для динамического UI:
    # берем ТОЛЬКО переменные из сборки — пользователь редактирует именно их.
    var_index: Dict[str, KompasVariable] = {}

    for name, kv in asm_vars.items():
        var_index[name] = kv

    logger.info(
        "Индекс переменных (var_index) сформирован из сборки: %s; переменных: %d",
        assembly_path,
        len(var_index),
    )

    return assembly_info, documents, var_index

