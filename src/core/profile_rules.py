"""
Правила для профилей NordFox и логика разбора наименований профилей.

Задачи модуля:
- разобрать строку вида "Профиль H20.1" на family/size/digit;
- предоставить ElementDesignationRule для ролей (стойка средняя, крайняя, ригель и т.п.);
- быть единой точкой расширения, когда добавятся новые профили и переменные.
"""

from __future__ import annotations

import re
import logging
from typing import Dict, Optional

from .models import KompasVariable, ProfileInfo, ElementDesignationRule

logger = logging.getLogger("ProfileRules")


# Базовые правила ролей элементов
ELEMENT_RULES: Dict[str, ElementDesignationRule] = {
    "stoika_srednyaya": ElementDesignationRule(
        role="stoika_srednyaya",
        prefix="СС",
        length_variable="Visota_srednei_stoiki",
    ),
    "stoika_kraynaya": ElementDesignationRule(
        role="stoika_kraynaya",
        prefix="СК",
        length_variable="Visota_stoiki",
    ),
    "rigel": ElementDesignationRule(
        role="rigel",
        prefix="Р",
        length_variable="Dlina_rigelya",
    ),
    # Для L-профиля пока только в плане: префикс зависит от размера
    "l_profile": ElementDesignationRule(
        role="l_profile",
        prefix_template="L{size}",
        length_variable="Dlina_L_profilya",  # будет уточнено позже
    ),
}


def parse_profile_name(name: str) -> Optional[ProfileInfo]:
    """
    Разобрать наименование профиля вида:
        "Профиль H20.1", "Профиль H20", "Профиль DT23",
        "Профиль H Hat22", "Профиль T15", "Профиль L15"

    Возвращает ProfileInfo или None, если строка не похожа на профиль.
    """
    text = name.strip()
    if not text.lower().startswith("профиль "):
        return None

    raw_suffix = text[len("Профиль ") :].strip()
    raw_suffix = " ".join(raw_suffix.split())

    # Специальные семейства с пробелом внутри, например "H Hat"
    special_families = ["H Hat", "T Hat"]
    for fam in special_families:
        if raw_suffix.startswith(fam + " "):
            family = fam
            rest = raw_suffix[len(fam) :].strip()
            break
    else:
        # Общее правило: family — всё до первой цифры
        match = re.match(r"([^\d]+)([\d].*)?$", raw_suffix)
        if not match:
            return None
        family = match.group(1).strip()
        rest = (match.group(2) or "").strip()

    # В rest ищем целое число размера (до первой нецифры или конца строки)
    size_match = re.match(r"(\d+)", rest)
    if not size_match:
        return None

    size = int(size_match.group(1))
    digit = size % 10

    return ProfileInfo(
        full_name=text,
        family=family,
        size=size,
        digit=digit,
        raw_suffix=raw_suffix,
    )


def get_element_rule(role: str) -> Optional[ElementDesignationRule]:
    """Вернуть правило обозначения для заданной роли элемента."""
    return ELEMENT_RULES.get(role)


def profile_short_code(full_profile_name: str) -> Optional[str]:
    """Краткий код профиля для суффикса обозначения, напр. «Профиль H20.1» → «H20.1»."""
    p = parse_profile_name(full_profile_name)
    if not p:
        return None
    return "".join(p.raw_suffix.split())


def _length_keys_for_role(role: str, rule: ElementDesignationRule) -> list[str]:
    keys: list[str] = []
    if rule.length_variable:
        keys.append(rule.length_variable)
    if role == "rigel":
        for k in ("Dlina_rigelya", "Dlina_rigelya_verhnego"):
            if k not in keys:
                keys.append(k)
    return keys


def length_mm_for_role(role: str, var_values: Dict[str, float]) -> Optional[float]:
    """Длина в мм по правилу роли и переменным сборки."""
    rule = ELEMENT_RULES.get(role)
    if not rule:
        return None
    for name in _length_keys_for_role(role, rule):
        if name in var_values:
            return float(var_values[name])
    return None


def collect_assembly_numeric_values(var_index: Dict[str, KompasVariable]) -> Dict[str, float]:
    """Числовые переменные сборки из индекса UI (имя → значение).

    Приоритет источников:
    1) kv.value (колонка "Значение");
    2) kv.expression (если в выражении есть числовой литерал).
    """
    def _to_float_or_none(raw: object) -> Optional[float]:
        if raw is None:
            return None
        if isinstance(raw, (int, float)):
            return float(raw)
        txt = str(raw).strip()
        if not txt:
            return None
        # Прямое преобразование (в т.ч. с десятичной запятой).
        try:
            return float(txt.replace(",", "."))
        except ValueError:
            pass
        # Fallback: извлекаем первое число из выражения.
        m = re.search(r"[-+]?\d+(?:[.,]\d+)?", txt)
        if not m:
            return None
        try:
            return float(m.group(0).replace(",", "."))
        except ValueError:
            return None

    out: Dict[str, float] = {}
    for name, kv in var_index.items():
        if kv.document_type != "assembly" or kv.is_block_header:
            continue
        value_num = _to_float_or_none(kv.value)
        if value_num is not None:
            out[name] = value_num
            logger.info("[Rules] %s: source=value, number=%s", name, value_num)
            continue
        expr_num = _to_float_or_none(kv.expression)
        if expr_num is not None:
            out[name] = expr_num
            logger.info("[Rules] %s: source=expression, number=%s", name, expr_num)
            continue
        logger.info("[Rules] %s: source=none, number=<skip>", name)
    return out


def infer_role_from_part_name(part_name: str) -> Optional[str]:
    """Угадать роль детали по наименованию (для автоподстановки обозначений)."""
    n = (part_name or "").lower()
    if "ригель" in n or "rigel" in n:
        return "rigel"
    if "стойка" in n or "stoik" in n:
        if "средн" in n or "sred" in n:
            return "stoika_srednyaya"
        return "stoika_kraynaya"
    return None


def build_element_designation(
    role: str,
    series_num: int,
    profile_full_name: str,
    var_values: Dict[str, float],
) -> Optional[str]:
    """
    Обозначение вида «СК-3-4000» (префикс — номер серии — длина).

    series_num: 1…4 и т.д. (выбор в интерфейсе).
    """
    rule = ELEMENT_RULES.get(role)
    if not rule or series_num < 1:
        return None
    length = length_mm_for_role(role, var_values)
    if length is None:
        return None
    prefix = (rule.prefix or "").strip()
    if rule.prefix_template:
        p = parse_profile_name(profile_full_name)
        if not p:
            return None
        prefix = rule.prefix_template.format(size=p.size)
    if not prefix:
        return None
    return f"{prefix}-{series_num}-{int(round(length))}"


def build_element_name(
    role: str,
    profile_full_name: str,
    var_values: Dict[str, float],
) -> Optional[str]:
    """
    Наименование вида «Профиль H21 (4000)» на основе роли, профиля и длины.
    """
    length = length_mm_for_role(role, var_values)
    short = profile_short_code(profile_full_name)
    if length is None or not short:
        return None
    return f"Профиль {short} ({int(round(length))})"

