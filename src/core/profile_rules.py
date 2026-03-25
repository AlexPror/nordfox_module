"""
Правила для профилей NordFox и логика разбора наименований профилей.

Задачи модуля:
- разобрать строку вида "Профиль H20.1" на family/size/digit;
- предоставить ElementDesignationRule для ролей (стойка средняя, крайняя, ригель и т.п.);
- быть единой точкой расширения, когда добавятся новые профили и переменные.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Optional

from .models import ProfileInfo, ElementDesignationRule


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
        length_variable="Dlina_rigelya_verhnego",
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

