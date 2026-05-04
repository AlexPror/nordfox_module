"""
Экспорт таблицы «Перечень чертежей комплекта» в фрагмент .frw (КОМПАС-3D).

Две таблицы в одном файле, если строк данных не помещаются в лимит.
Паттерн построения таблицы — как в nordfox_specification (ksTable / ksLineSeg / ksText).
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from .kompas_connector import KompasConnector

logger = logging.getLogger(__name__)

# Максимум строк данных в одном фрагменте .frw (продолжение — следующий блок).
_FRW_MAX_DATA_ROWS_PER_TABLE = 38

TITLE = "Перечень чертежей комплекта"
HEADER = ("Лист", "Наименование", "Примечание")


def _project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def resolve_frw_template_path() -> Optional[Path]:
    """Пустой шаблон-фрагмент Table.frw (как в Specification Builder)."""
    candidates = [
        _project_root() / "table_frw" / "Table.frw",
        Path(os.environ.get("NORDFOX_TABLE_FRW", "")).expanduser(),
        _project_root().parent / "nordfox_specification" / "table_frw" / "Table.frw",
    ]
    for p in candidates:
        if p and p.is_file() and p.suffix.lower() == ".frw":
            return p.resolve()
    return None


def _get_constants(api5):
    try:
        from win32com.client import gencache

        mod = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0)
        return mod.constants
    except Exception as exc:
        logger.error("KOMPAS constants: %s", exc)
        return None


def _try_open_template(api7, template: Path):
    try:
        return api7.Documents.Open(str(template), False)
    except Exception as exc:
        logger.error("Documents.Open(%s): %s", template, exc)
        return None


def _try_create_fragment(api7, constants) -> Optional[object]:
    if constants is None:
        return None
    for attr in (
        "ksDocumentFragment",
        "ksDocumentDrawing",
        "ksDocumentDrawings",
        "ksDocumentLayout",
    ):
        if not hasattr(constants, attr):
            continue
        dt = getattr(constants, attr)
        try:
            doc = api7.Documents.Add(dt, True)
            if doc:
                logger.info("Создан новый документ через Documents.Add(%s)", attr)
                return doc
        except Exception as exc:
            logger.debug("Documents.Add(%s): %s", attr, exc)
    return None


def _save_frw(doc7, file_path: Path) -> bool:
    target = str(file_path)
    for args in ((target,), (target, True), (target, False)):
        try:
            if doc7.SaveAs(*args):
                return True
        except Exception:
            continue
    return False


def _text_height_mm(font_pt: float) -> float:
    return max(1.8, font_pt * (25.4 / 72.0) * 0.88)


def _place_text_cell(
    doc2d,
    x_left: float,
    y_top: float,
    y_bottom: float,
    text: str,
    font_pt: float,
    pad_x: float,
    pad_y: float,
) -> None:
    if not (text or "").strip():
        return
    text_h = _text_height_mm(font_pt)
    line_pitch = text_h * 1.08
    y_text = y_top - pad_y - text_h
    for line in str(text).replace("\r", "").split("\n"):
        if not line.strip():
            continue
        if y_text <= y_bottom + pad_y:
            break
        try:
            doc2d.ksText(x_left + pad_x, y_text, 0.0, text_h, 1.0, 0, line)
        except Exception:
            pass
        y_text -= line_pitch


def _build_single_table(
    doc2d,
    *,
    x0: float,
    y0_bottom: float,
    col_widths: Sequence[float],
    title_row_h: float,
    header_row_h: float,
    data_row_h: float,
    font_pt: float,
    data_rows: List[Tuple[str, str, str]],
    table_title: str = TITLE,
    merge_title: bool = True,
) -> float:
    """
    Строит одну таблицу; возвращает total_height (мм).
    data_rows — только строки данных (лист, наименование, примечание).
    """
    cols = 3
    rows = 2 + len(data_rows)
    row_heights = [title_row_h, header_row_h] + [data_row_h] * len(data_rows)
    total_height = sum(row_heights)
    total_width = sum(col_widths)

    try:
        table_id = doc2d.ksTable()
    except Exception as exc:
        logger.warning("ksTable: %s", exc)
        table_id = 0

    style_thin = 2
    y = y0_bottom
    for i in range(rows + 1):
        try:
            doc2d.ksLineSeg(x0, y, x0 + total_width, y, style_thin)
        except Exception:
            pass
        if i < rows:
            y += row_heights[rows - 1 - i]

    x = x0
    for c in range(cols + 1):
        try:
            doc2d.ksLineSeg(x, y0_bottom, x, y0_bottom + total_height, style_thin)
        except Exception:
            pass
        x += col_widths[c] if c < cols else 0.0

    prefix_h = [0.0]
    for h in row_heights:
        prefix_h.append(prefix_h[-1] + h)

    pad_x = 0.8
    pad_y = 0.85

    matrix: List[List[str]] = []
    matrix.append([table_title, "", ""])
    matrix.append([HEADER[0], HEADER[1], HEADER[2]])
    for r in data_rows:
        matrix.append([r[0], r[1], r[2]])

    for r in range(rows):
        for c in range(cols):
            val = matrix[r][c] if c < len(matrix[r]) else ""
            x_left = x0 + sum(col_widths[:c])
            y_cell_top = y0_bottom + total_height - prefix_h[r]
            y_cell_bottom = y0_bottom + total_height - prefix_h[r + 1]
            _place_text_cell(doc2d, x_left, y_cell_top, y_cell_bottom, val, font_pt, pad_x, pad_y)

    if table_id:
        try:
            table_id = doc2d.ksEndObj()
        except Exception:
            pass

    if merge_title and table_id:
        try:
            if doc2d.ksOpenTable(table_id):
                base = 1
                doc2d.ksCombineTwoTableItems(base, base + 1)
                doc2d.ksCombineTwoTableItems(base, base + 2)
                doc2d.ksRebuildTableVirtualGrid()
            doc2d.ksEndObj()
        except Exception as exc:
            logger.warning("Объединение ячеек заголовка: %s", exc)
            try:
                doc2d.ksEndObj()
            except Exception:
                pass

    return total_height


def export_register_frw(
    data_rows: List[Tuple[str, str, str]],
    output_path: Path,
    connector: KompasConnector,
    *,
    max_data_rows_per_table: int = 28,
    data_row_height_mm: float = 5.0,
    title_row_height_mm: float = 6.5,
    header_row_height_mm: float = 5.0,
    col_widths: Tuple[float, float, float] = (14.0, 102.0, 32.0),
    gap_between_tables_mm: float = 6.0,
    font_pt: float = 3.6,
    continuation_suffix: str = " (продолжение)",
) -> tuple[bool, str]:
    """
    data_rows: список (Лист, Наименование, Примечание) без шапки.
    """
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not data_rows:
        return False, "Нет строк для таблицы перечня."

    if not connector.connect(force_reconnect=False):
        return False, "Не удалось подключиться к КОМПАС-3D."

    api5 = connector.api5
    api7 = connector.api7
    if api5 is None or api7 is None:
        return False, "API5/API7 недоступны."

    constants = _get_constants(api5)
    doc7 = None
    template = resolve_frw_template_path()
    if template:
        doc7 = _try_open_template(api7, template)
    if doc7 is None:
        doc7 = _try_create_fragment(api7, constants)
    if doc7 is None:
        return (
            False,
            "Не найден шаблон table_frw/Table.frw и не удалось создать фрагмент через API. "
            "Скопируйте Table.frw из проекта Specification Builder в "
            f"{_project_root() / 'table_frw'} или задайте NORDFOX_TABLE_FRW.",
        )

    import time

    time.sleep(0.2)
    try:
        api7.ActiveDocument = doc7
    except Exception:
        pass

    doc2d = api5.ActiveDocument2D
    if doc2d is None:
        return False, "ActiveDocument2D недоступен для фрагмента."

    chunks: List[List[Tuple[str, str, str]]] = []
    n = max(1, min(_FRW_MAX_DATA_ROWS_PER_TABLE, int(max_data_rows_per_table)))
    for i in range(0, len(data_rows), n):
        chunks.append(data_rows[i : i + n])

    y_cursor = 0.0
    for idx, chunk in enumerate(chunks):
        title = TITLE if idx == 0 else TITLE + continuation_suffix
        h = _build_single_table(
            doc2d,
            x0=0.0,
            y0_bottom=y_cursor,
            col_widths=col_widths,
            title_row_h=title_row_height_mm,
            header_row_h=header_row_height_mm,
            data_row_h=data_row_height_mm,
            font_pt=font_pt,
            data_rows=chunk,
            table_title=title,
            merge_title=True,
        )
        y_cursor += h + (gap_between_tables_mm if idx < len(chunks) - 1 else 0.0)

    try:
        if hasattr(doc2d, "ksReDrawDocPart"):
            doc2d.ksReDrawDocPart(0, 0, 0, 0)
    except Exception:
        pass

    saved = False
    try:
        if hasattr(doc2d, "ksSaveDocument"):
            saved = bool(doc2d.ksSaveDocument(str(output_path)))
    except Exception:
        saved = False
    if not saved:
        saved = _save_frw(doc7, output_path)

    if not saved or not output_path.exists():
        return False, f"Не удалось сохранить {output_path}"

    logger.info("Перечень чертежей экспортирован в FRW: %s", output_path)
    return True, str(output_path)
