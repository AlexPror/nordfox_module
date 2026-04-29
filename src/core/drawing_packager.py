"""
Нумерация чертежей комплекта: согласование префикса в имени и «(лист N)»,
двухфазное переименование без коллизий имён.
"""

from __future__ import annotations

import logging
import re
import uuid
from pathlib import Path
from typing import List, Optional, Tuple

logger = logging.getLogger(__name__)

_SHEET_RE = re.compile(r"\(лист\s*(\d+)\)\s*$", re.IGNORECASE)


def parse_cdw_stem(stem: str) -> tuple[Optional[int], str, Optional[int]]:
    """
    Разбор имени без расширения:
    - ведущий номер файла (если есть),
    - средняя часть,
    - номер из «(лист N)» в конце (если есть).
    """
    m_sheet = _SHEET_RE.search(stem)
    sheet_n = int(m_sheet.group(1)) if m_sheet else None
    base = stem[: m_sheet.start()].rstrip() if m_sheet else stem
    m_lead = re.match(r"^(\d+)\s+(.+)$", base)
    if m_lead:
        return int(m_lead.group(1)), m_lead.group(2).strip(), sheet_n
    return None, base.strip(), sheet_n


def build_new_filename(middle: str, index: int) -> str:
    """«{index} {middle} (лист {index}).cdw»."""
    mid = (middle or "").strip()
    if mid:
        return f"{index} {mid} (лист {index}).cdw"
    return f"{index} (лист {index}).cdw"


def format_register_name_from_middle(middle: str) -> str:
    """
    Текст для колонки «Наименование» в перечне чертежей (.frw).

    Из средней части имени (как в parse_cdw_stem, без номера файла и «(лист N)»),
    отбрасывается обозначение после разделителя « _ » и для модулей применяется вид::

        Модуль М1-1.1-СП-5 _ 10-23-КП-Р-КМД1.1-1-01 → Модуль: М1-1.1-СП-5;

    Иначе возвращается часть до « _ » или исходная строка без лишних пробелов.
    """
    s = (middle or "").strip()
    if not s:
        return ""
    if " _ " in s:
        s = s.split(" _ ", 1)[0].strip()
    m = re.match(r"^Модуль\s+(.+)$", s, re.IGNORECASE)
    if m:
        code = m.group(1).strip()
        return f"Модуль: {code};"
    return s


def append_material_to_register_line(register_name: str, material_line: str) -> str:
    """Добавить к наименованию в перечне текст из ячейки «материал» штампа (например «Узлы 1, 1.1, …»)."""
    mat = (material_line or "").strip()
    if not mat:
        return (register_name or "").strip()
    base = (register_name or "").strip()
    if not base:
        return mat
    if mat.lower() in base.lower():
        return base
    return f"{base} {mat}"


def plan_renames_for_order(
    ordered_paths: List[Path],
    *,
    middle_parts: Optional[List[str]] = None,
) -> List[Tuple[Path, Path]]:
    """Упорядоченный список .cdw → пары (как сейчас, как должно быть)."""
    if middle_parts is not None and len(middle_parts) != len(ordered_paths):
        raise ValueError("middle_parts и список файлов должны быть одной длины")
    out: List[Tuple[Path, Path]] = []
    for i, p in enumerate(ordered_paths, start=1):
        if middle_parts is not None:
            mid = middle_parts[i - 1]
        else:
            _lead, mid, _sheet = parse_cdw_stem(p.stem)
        new_name = build_new_filename(mid, i)
        new_path = (p.parent / new_name).resolve()
        out.append((p.resolve(), new_path))
    return out


def apply_renames_two_phase(plans: List[Tuple[Path, Path]]) -> tuple[bool, List[str]]:
    """Два прохода через уникальные временные имена."""
    errors: List[str] = []
    work = [(o, n) for o, n in plans if o != n]
    if not work:
        return True, []

    for old, new in work:
        if not old.exists():
            errors.append(f"Файл не найден: {old}")
            return False, errors
        if new.exists() and new.resolve() != old.resolve():
            errors.append(f"Целевое имя уже занято: {new.name}")
            return False, errors

    tmp_tag = uuid.uuid4().hex[:10]
    tmps: List[Path] = []
    try:
        for idx, (old, _new) in enumerate(work):
            tmp = old.parent / f"__nfx_{tmp_tag}_{idx:04d}.cdw"
            try:
                old.rename(tmp)
            except OSError as exc:
                errors.append(f"Временное переименование {old.name}: {exc}")
                return False, errors
            tmps.append(tmp)

        for tmp, (_old, new) in zip(tmps, work):
            try:
                tmp.rename(new)
            except OSError as exc:
                errors.append(f"Финальное имя {new.name}: {exc}")
                return False, errors

        logger.info("Переименовано чертежей комплекта: %s", len(work))
        return True, []
    except Exception as exc:  # pragma: no cover
        errors.append(str(exc))
        return False, errors


def sheet_number_from_name(path: Path) -> Optional[int]:
    _l, _m, s = parse_cdw_stem(path.stem)
    return s
