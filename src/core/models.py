"""
Модели данных для NordFox Module Manager.

Здесь описаны структуры, в которых мы храним информацию:
- о документах КОМПАС-3D (сборка, детали, чертежи);
- о переменных;
- о профилях и обозначениях;
- о действиях для JSON-логов.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal, Any, Dict, List, Optional


DocumentType = Literal["assembly", "part", "drawing"]


@dataclass
class KompasVariable:
    """Одна переменная в документе КОМПАС-3D."""

    name: str
    value: Any
    original_value: Any
    document_type: DocumentType
    document_path: Path
    comment: Optional[str] = None
    expression: Optional[str] = None
    block_id: Optional[str] = None          # Имя блока (Stoiki, Rigel, Obolochka, ...)
    is_block_header: bool = False           # True для пустой переменной-заголовка блока
    is_external: bool = False
    is_designation_related: bool = False
    is_name_related: bool = False


@dataclass
class KompasDocumentInfo:
    """Информация о конкретном файле КОМПАС-3D."""

    path: Path
    doc_type: DocumentType
    designation: Optional[str] = None
    name: Optional[str] = None
    variables: Dict[str, KompasVariable] = field(default_factory=dict)


@dataclass
class ProfileInfo:
    """
    Информация о типе профиля, извлеченная из наименования.

    Примеры:
        "Профиль H20.1" -> family="H", size=20, digit=0
        "Профиль DT23"  -> family="DT", size=23, digit=3
    """

    full_name: str              # Полное имя, например "Профиль H20.1"
    family: str                 # "H", "DT", "H Hat", "T", "L", ...
    size: int                   # 20, 23, 15 ...
    digit: int                  # size % 10
    raw_suffix: str             # Часть после "Профиль ", например "H20.1"


@dataclass
class ElementDesignationRule:
    """
    Правило формирования обозначения и привязка к переменной длины.

    Примеры ролей:
        "stoika_srednyaya", "stoika_kraynaya", "rigel", "l_profile".
    """

    role: str                         # Внутреннее имя роли
    prefix: Optional[str] = None      # "СС", "СК", "Р" и т.п.
    prefix_template: Optional[str] = None  # Например "L{size}" для L15-400
    length_variable: str = ""         # Имя переменной длины


@dataclass
class JsonLogAction:
    """
    Описание одного действия для JSON-лога.

    Поля совпадают с ранее обсужденной структурой:
    - id, type, timestamp, status, input, changes, meta.
    """

    id: str
    type: str
    timestamp: str
    status: Literal["success", "error", "partial"]
    input: Dict[str, Any] = field(default_factory=dict)
    changes: Dict[str, Any] = field(default_factory=dict)
    meta: Dict[str, Any] = field(default_factory=dict)


@dataclass
class JsonSessionLog:
    """Полный JSON-лог сессии."""

    app_name: str
    app_version: str
    started_at: str
    ended_at: Optional[str]
    project_root: Path
    project_state: Dict[str, Any] = field(default_factory=dict)
    actions: List[JsonLogAction] = field(default_factory=list)

