"""
Главное окно NordFox Module Manager (PyQt6).

Задачи:
- выбор папки проекта (кнопка "Открыть");
- отображение обозначения и наименования сборки;
- динамическая зона для переменных (будет заполняться после сканирования);
- базовая вкладка для QR-кодов (генерация PNG);
- панель логов и статус-бар.

Детальная логика сканирования проекта и обновления переменных/обозначений
будет реализована в отдельных core-модулях и вызываться из этого окна.
"""

from __future__ import annotations

import logging
import json
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Dict, Optional

from PyQt6.QtCore import Qt, QDateTime
from PyQt6.QtGui import QAction, QFont
from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QLabel,
    QLineEdit,
    QPushButton,
    QStatusBar,
    QSplitter,
    QPlainTextEdit,
    QTabWidget,
    QScrollArea,
    QFormLayout,
    QComboBox,
    QToolBox,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QCheckBox,
    QSpinBox,
)
from PyQt6.QtWidgets import QApplication, QProgressDialog

from ..core.qr_generator import generate_qr_png
from ..core.kompas_connector import KompasConnector
from ..core.variables_scanner import scan_project
from ..core.variables_updater import update_project_variables
from ..core.log_store import JsonLogStore
from ..core.models import KompasDocumentInfo, KompasVariable
from ..core.stamp_updater import update_all_drawing_stamps
from ..core.drawing_pdf_exporter import DrawingPdfExporter
from ..core.drawing_dwg_exporter import DrawingDwgExporter
from ..core.project_copy import copy_project_tree
from ..core.profile_rules import (
    build_element_designation,
    build_element_name,
    collect_assembly_numeric_values,
    infer_role_from_part_name,
)


logger = logging.getLogger(__name__)


class QtTextLogHandler(logging.Handler):
    """
    Обработчик логов, который пишет сообщения в QPlainTextEdit.
    Ожидается, что используется из GUI-потока.
    """

    def __init__(self, widget: QPlainTextEdit) -> None:
        super().__init__()
        self._widget = widget
        formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%H:%M:%S",
        )
        self.setFormatter(formatter)

    def emit(self, record: logging.LogRecord) -> None:  # pragma: no cover - GUI
        msg = self.format(record)
        # Добавляем строку в конец лога
        self._widget.appendPlainText(msg)
        # Прокрутка вниз
        cursor = self._widget.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self._widget.setTextCursor(cursor)


class MainWindow(QMainWindow):
    """Главное окно приложения."""

    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle("NordFox Module Manager")
        self.setMinimumSize(1200, 800)

        self._current_project_root: Optional[Path] = None
        self._assembly_info: Optional[KompasDocumentInfo] = None
        self._documents: list[KompasDocumentInfo] = []
        self._var_index: Dict[str, KompasVariable] = {}
        # Поля для переменных сборки (значения)
        self._var_inputs: Dict[str, QLineEdit] = {}
        # Поля для переменных чертежей (комментарии): имя переменной -> QLineEdit
        self._drawing_comment_inputs: Dict[str, QLineEdit] = {}
        # Таблица элементов: путь документа -> KompasDocumentInfo.
        # Это исключает "схлопывание" разных файлов с одинаковыми mark/name.
        self._assembly_items: Dict[Path, KompasDocumentInfo] = {}
        self._assembly_item_new_marking: Dict[Path, QLineEdit] = {}
        self._assembly_item_new_name: Dict[Path, QLineEdit] = {}

        self._kompas = KompasConnector()
        self._json_log: Optional[JsonLogStore] = None
        self._updating_in_progress: bool = False

        self._setup_ui()
        self._setup_menu()
        self._setup_statusbar()

        logger.info("Главное окно NordFox Module Manager создано")

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _setup_ui(self) -> None:
        central = QWidget(self)
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # Блок выбора проекта и данных сборки
        top_group = QGroupBox("Проект NordFox")
        top_layout = QVBoxLayout(top_group)

        # Строка выбора папки проекта
        project_row = QHBoxLayout()
        lbl_project = QLabel("Папка проекта:")
        self.project_path_edit = QLineEdit()
        self.project_path_edit.setPlaceholderText("Выберите папку с КОМПАС-проектом (сборка модулей NordFox)...")
        self.project_path_edit.setReadOnly(True)
        btn_browse = QPushButton("Открыть...")
        btn_browse.clicked.connect(self._browse_project_folder)

        project_row.addWidget(lbl_project)
        project_row.addWidget(self.project_path_edit)
        project_row.addWidget(btn_browse)
        top_layout.addLayout(project_row)

        # Строка режима "копия + обновление"
        copy_row = QHBoxLayout()
        self.copy_mode_check = QCheckBox("Работать с копией проекта")
        self.copy_mode_check.setChecked(False)
        self.copy_mode_check.setVisible(False)
        self.copy_target_edit = QLineEdit()
        self.copy_target_edit.setPlaceholderText("Папка назначения для копии проекта...")
        self.btn_copy_target_browse = QPushButton("Папка копии...")
        self.btn_copy_target_browse.clicked.connect(self._browse_copy_target_folder)
        self.btn_copy_and_update = QPushButton("Создать проект")
        self.btn_copy_and_update.setEnabled(False)
        self.btn_copy_and_update.clicked.connect(self._on_copy_and_update_clicked)
        copy_row.addWidget(self.copy_mode_check)
        copy_row.addWidget(self.copy_target_edit)
        copy_row.addWidget(self.btn_copy_target_browse)
        copy_row.addWidget(self.btn_copy_and_update)
        top_layout.addLayout(copy_row)

        hint_copy = QLabel(
            "Кнопка «Создать проект» делает копию в указанной папке. "
            "Кнопки «Обновить переменные» и «Пересканировать» работают только с текущей «Папкой проекта»."
        )
        hint_copy.setWordWrap(True)
        hint_copy.setStyleSheet("color: palette(mid); font-size: 11px;")
        top_layout.addWidget(hint_copy)

        # Строка обозначение / наименование сборки
        assembly_row = QHBoxLayout()
        self.assembly_designation_edit = QLineEdit()
        self.assembly_name_edit = QLineEdit()
        self.assembly_designation_edit.setPlaceholderText("Обозначение сборки (будет прочитано из КОМПАС)...")
        self.assembly_name_edit.setPlaceholderText("Наименование сборки (будет прочитано из КОМПАС)...")

        lbl_assm_des = QLabel("Обозначение сборки:")
        lbl_assm_name = QLabel("Наименование сборки:")
        assembly_row.addWidget(lbl_assm_des)
        assembly_row.addWidget(self.assembly_designation_edit)
        assembly_row.addWidget(lbl_assm_name)
        assembly_row.addWidget(self.assembly_name_edit)

        top_layout.addLayout(assembly_row)

        main_layout.addWidget(top_group)

        # Основной сплиттер: слева вкладки (переменные / QR), справа логи
        splitter = QSplitter(Qt.Orientation.Horizontal, self)

        # Левая часть: вкладки
        left_tabs = QTabWidget()
        left_tabs.setTabPosition(QTabWidget.TabPosition.North)
        left_tabs.setDocumentMode(True)
        left_tabs.setStyleSheet("QTabBar::tab { height: 28px; }")

        # Вкладка "Переменные" (форма будет наполняться после сканирования)
        self.vars_tab = QWidget()
        vars_layout = QVBoxLayout(self.vars_tab)

        self.vars_scroll = QScrollArea()
        self.vars_scroll.setWidgetResizable(True)
        self.vars_blocks_container = QWidget()
        self.vars_blocks_layout = QVBoxLayout(self.vars_blocks_container)
        self.vars_blocks_layout.setContentsMargins(0, 0, 0, 0)
        self.vars_blocks_layout.setSpacing(8)
        self.vars_toolbox = QToolBox()
        self.vars_blocks_layout.addWidget(self.vars_toolbox)
        self.vars_scroll.setWidget(self.vars_blocks_container)

        vars_layout.addWidget(self.vars_scroll)

        buttons_row = QHBoxLayout()
        self.btn_update_variables = QPushButton("Обновить переменные")
        self.btn_update_variables.setEnabled(False)
        self.btn_update_variables.clicked.connect(self._on_update_variables_clicked)

        self.btn_rescan = QPushButton("Пересканировать")
        self.btn_rescan.setEnabled(False)
        self.btn_rescan.clicked.connect(self._on_rescan_clicked)

        buttons_row.addWidget(self.btn_update_variables)
        buttons_row.addWidget(self.btn_rescan)

        vars_layout.addLayout(buttons_row)

        left_tabs.addTab(self.vars_tab, "Переменные")

        # Вкладка "Детали/подсборки" (таблица обозначений/наименований)
        self.items_tab = QWidget()
        items_layout = QVBoxLayout(self.items_tab)

        items_toolbar = QHBoxLayout()
        items_toolbar.addWidget(QLabel("Серия (1–4):"))
        self.items_series_combo = QComboBox()
        for i in range(1, 5):
            self.items_series_combo.addItem(str(i))
        items_toolbar.addWidget(self.items_series_combo)
        items_toolbar.addWidget(QLabel("Профиль для обозначения:"))
        self.items_profile_combo = QComboBox()
        self.items_profile_combo.setMinimumWidth(220)
        self._fill_profile_combo(self.items_profile_combo)
        items_toolbar.addWidget(self.items_profile_combo, stretch=1)
        self.btn_apply_designation_rules = QPushButton("Заполнить «новое обозначение» по правилам")
        self.btn_apply_designation_rules.setToolTip(
            "По переменным сборки (длины СК/СС/Р) и выбранному профилю "
            "заполняет колонку «Новое обозначение» для деталей, где роль угадывается по наименованию."
        )
        self.btn_apply_designation_rules.clicked.connect(self._on_apply_designation_rules)
        items_toolbar.addWidget(self.btn_apply_designation_rules)
        self.btn_apply_items_meta = QPushButton("Применить новые обозн./наимен.")
        self.btn_apply_items_meta.setToolTip(
            "Записывает непустые значения из колонок «Новое обозначение» / "
            "«Новое наименование» в свойства деталей/сборки."
        )
        self.btn_apply_items_meta.clicked.connect(self._on_apply_items_meta_clicked)
        items_toolbar.addWidget(self.btn_apply_items_meta)
        self.btn_sync_assembly_components = QPushButton("Синхр. вхождения в сборке")
        # SAFE MODE: временно отключено из-за нестабильного COM-crash в окружении пользователя.
        self.btn_sync_assembly_components.setEnabled(False)
        self.btn_sync_assembly_components.setToolTip(
            "Временно отключено: синхронизация вхождений сборки может аварийно завершать КОМПАС."
        )
        items_toolbar.addWidget(self.btn_sync_assembly_components)
        items_layout.addLayout(items_toolbar)

        self.items_table = QTableWidget(0, 6, self.items_tab)
        self.items_table.setHorizontalHeaderLabels(
            [
                "Тип",
                "Текущее обозначение",
                "Текущее наименование",
                "Новое обозначение",
                "Новое наименование",
                "Файл",
            ]
        )
        self.items_table.horizontalHeader().setStretchLastSection(True)
        self.items_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        # Ячейки "новых" значений будут редактируемыми через QLineEdit в качестве виджетов

        items_layout.addWidget(self.items_table)
        items_layout.addStretch(1)

        left_tabs.addTab(self.items_tab, "Обозн./наимен.")

        # Вкладка "Штампы"
        self.stamp_tab = QWidget()
        stamp_layout = QVBoxLayout(self.stamp_tab)

        stamp_form = QFormLayout()
        stamp_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.stamp_designation_edit = QLineEdit()
        self.stamp_name_edit = QLineEdit()
        self.stamp_designation_edit.setPlaceholderText("Обозначение (ячейка 2, см. stamp_cells.py)")
        self.stamp_name_edit.setPlaceholderText("Наименование (ячейка 1)")

        self.stamp_sheet_auto_check = QCheckBox("Нумеровать листы автоматически (1…N, порядок — см. справку)")
        self.stamp_sheet_auto_check.setChecked(False)
        sheet_row = QHBoxLayout()
        sheet_row.addWidget(QLabel("Лист (вручную):"))
        self.stamp_sheet_current_spin = QSpinBox()
        self.stamp_sheet_current_spin.setRange(0, 9999)
        self.stamp_sheet_current_spin.setSpecialValueText("—")
        self.stamp_sheet_current_spin.setValue(0)
        self.stamp_sheet_total_spin = QSpinBox()
        self.stamp_sheet_total_spin.setRange(0, 9999)
        self.stamp_sheet_total_spin.setSpecialValueText("—")
        self.stamp_sheet_total_spin.setValue(0)
        sheet_row.addWidget(self.stamp_sheet_current_spin)
        sheet_row.addWidget(QLabel("из"))
        sheet_row.addWidget(self.stamp_sheet_total_spin)
        sheet_row.addStretch(1)

        self.stamp_developer_edit = self._make_person_combo(["Воробьев", "Заметалин", "Сизонов"])
        self.stamp_checker_edit = self._make_person_combo(["Заметалин", "Воробьев", "Сизонов"])
        self.stamp_org_edit = QLineEdit()
        # Материал задаётся через выпадающие списки или строку-заглушку.
        self.stamp_material_thickness_combo = QComboBox()
        self.stamp_material_type_combo = QComboBox()
        self.stamp_material_edit = QLineEdit()
        self.stamp_material_edit.setReadOnly(False)
        self.stamp_tech_ctrl_edit = self._make_person_combo([])
        self.stamp_norm_ctrl_edit = self._make_person_combo([])
        self.stamp_approved_edit = self._make_person_combo(["Сизонов", "Заметалин"])
        self.stamp_date_edit = QLineEdit()

        self.stamp_developer_edit.setPlaceholderText("Разраб. (ячейка 110)")
        self.stamp_checker_edit.setPlaceholderText("Пров. (ячейка 111)")
        self.stamp_org_edit.setPlaceholderText("Организация (ячейка 9)")
        self.stamp_material_thickness_combo.addItems(["", "6", "5", "4", "3", "2", "1"])
        self.stamp_material_type_combo.addItems([
            "",
            "Конструкции металлические деталировочные",
            "09Г2С ГОСТ 19281-2014",
            "AISI 304",
            "Ст3. ГОСТ 19281-2014",
        ])
        self.stamp_material_edit.setPlaceholderText(
            "Материал (ячейка 3). Можно ввести вручную или выбрать тип выше."
        )
        self.stamp_tech_ctrl_edit.setPlaceholderText("Т. контр. (ячейка 112)")
        self.stamp_norm_ctrl_edit.setPlaceholderText("Н. контр. (ячейка 114)")
        self.stamp_approved_edit.setPlaceholderText("Утв. (ячейка 115)")
        self.stamp_date_edit.setPlaceholderText("Дата (ячейка 130, формат 01.01.2026)")

        stamp_form.addRow("Обозначение:", self.stamp_designation_edit)
        stamp_form.addRow("Наименование:", self.stamp_name_edit)
        stamp_form.addRow("", self.stamp_sheet_auto_check)
        stamp_form.addRow("Листы:", sheet_row)
        stamp_form.addRow("Разработал:", self.stamp_developer_edit)
        stamp_form.addRow("Проверил:", self.stamp_checker_edit)
        stamp_form.addRow("Организация:", self.stamp_org_edit)
        # Материал: выпадающие списки + текстовое поле
        material_row = QHBoxLayout()
        material_row.addWidget(QLabel("Толщина:"))
        material_row.addWidget(self.stamp_material_thickness_combo)
        material_row.addWidget(QLabel("Материал:"))
        material_row.addWidget(self.stamp_material_type_combo)
        material_row.addWidget(QLabel("Строка:"))
        material_row.addWidget(self.stamp_material_edit)
        stamp_form.addRow("Материал:", material_row)
        stamp_form.addRow("Т. контр.:", self.stamp_tech_ctrl_edit)
        stamp_form.addRow("Н. контр.:", self.stamp_norm_ctrl_edit)
        stamp_form.addRow("Утв.:", self.stamp_approved_edit)
        stamp_form.addRow("Дата:", self.stamp_date_edit)
        stamp_layout.addLayout(stamp_form)

        hint_stamp = QLabel(
            "Номера ячеек штампа ГОСТ 2.104 заданы в src/core/stamp_cells.py. "
            "Для материала используется формат КОМПАС: $d ... ; ... $."
        )
        hint_stamp.setWordWrap(True)
        hint_stamp.setStyleSheet("color: palette(mid); font-size: 11px;")
        stamp_layout.addWidget(hint_stamp)

        sheet_auto_group = QGroupBox("Автонумерация листов")
        sheet_auto_layout = QHBoxLayout(sheet_auto_group)
        self.stamp_sheet_folder_edit = QLineEdit()
        self.stamp_sheet_folder_edit.setPlaceholderText("Папка чертежей (.cdw); по умолчанию Папка проекта")
        self.btn_sheet_folder_browse = QPushButton("Папка чертежей...")
        self.btn_sheet_folder_browse.clicked.connect(self._browse_sheet_folder)
        self.btn_auto_number_sheets = QPushButton("Пронумеровать листы")
        self.btn_auto_number_sheets.clicked.connect(self._on_auto_number_sheets_clicked)
        sheet_auto_layout.addWidget(self.stamp_sheet_folder_edit, stretch=1)
        sheet_auto_layout.addWidget(self.btn_sheet_folder_browse)
        sheet_auto_layout.addWidget(self.btn_auto_number_sheets)
        stamp_layout.addWidget(sheet_auto_group)

        stamp_buttons_row = QHBoxLayout()
        self.btn_update_stamps = QPushButton("Обновить штампы чертежей")
        self.btn_update_stamps.setEnabled(True)
        self.btn_update_stamps.clicked.connect(self._on_update_stamps_clicked)
        stamp_buttons_row.addStretch(1)
        stamp_buttons_row.addWidget(self.btn_update_stamps)

        stamp_layout.addLayout(stamp_buttons_row)
        stamp_layout.addStretch(1)

        left_tabs.addTab(self.stamp_tab, "Штампы")

        # Вкладка "Экспорт PDF"
        self.pdf_tab = QWidget()
        pdf_layout = QVBoxLayout(self.pdf_tab)

        pdf_group = QGroupBox("Экспорт CDW -> PDF")
        pdf_group_layout = QFormLayout(pdf_group)
        pdf_group_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        pdf_source_row = QHBoxLayout()
        self.pdf_source_folder_edit = QLineEdit()
        self.pdf_source_folder_edit.setPlaceholderText("Папка с чертежами .cdw (по умолчанию: папка проекта)")
        self.btn_pdf_source_folder_browse = QPushButton("Папка чертежей...")
        self.btn_pdf_source_folder_browse.clicked.connect(self._browse_pdf_source_folder)
        pdf_source_row.addWidget(self.pdf_source_folder_edit, stretch=1)
        pdf_source_row.addWidget(self.btn_pdf_source_folder_browse)

        pdf_output_row = QHBoxLayout()
        self.pdf_export_folder_edit = QLineEdit()
        self.pdf_export_folder_edit.setPlaceholderText("Папка для PDF (по умолчанию: <папка чертежей>/PDF)")
        self.btn_pdf_export_folder_browse = QPushButton("Папка PDF...")
        self.btn_pdf_export_folder_browse.clicked.connect(self._browse_pdf_export_folder)
        pdf_output_row.addWidget(self.pdf_export_folder_edit, stretch=1)
        pdf_output_row.addWidget(self.btn_pdf_export_folder_browse)

        self.pdf_merge_check = QCheckBox("Объединить все PDF в один файл")
        self.pdf_merge_check.setChecked(True)
        self.pdf_merged_name_edit = QLineEdit()
        self.pdf_merged_name_edit.setPlaceholderText("Имя объединенного PDF (необязательно)")

        pdf_export_btn_row = QHBoxLayout()
        self.btn_export_drawings_pdf = QPushButton("Экспорт CDW -> PDF")
        self.btn_export_drawings_pdf.clicked.connect(self._on_export_drawings_pdf_clicked)
        pdf_export_btn_row.addStretch(1)
        pdf_export_btn_row.addWidget(self.btn_export_drawings_pdf)

        pdf_group_layout.addRow("Папка чертежей:", pdf_source_row)
        pdf_group_layout.addRow("Папка вывода:", pdf_output_row)
        pdf_group_layout.addRow("", self.pdf_merge_check)
        pdf_group_layout.addRow("Объединенный файл:", self.pdf_merged_name_edit)
        pdf_group_layout.addRow("", pdf_export_btn_row)
        pdf_layout.addWidget(pdf_group)
        pdf_layout.addStretch(1)
        left_tabs.addTab(self.pdf_tab, "Экспорт PDF")

        # Вкладка "Экспорт DWG"
        self.dwg_tab = QWidget()
        dwg_layout = QVBoxLayout(self.dwg_tab)

        dwg_group = QGroupBox("Экспорт CDW -> DWG")
        dwg_group_layout = QFormLayout(dwg_group)
        dwg_group_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        dwg_source_row = QHBoxLayout()
        self.dwg_source_folder_edit = QLineEdit()
        self.dwg_source_folder_edit.setPlaceholderText("Папка с чертежами .cdw (по умолчанию: папка проекта)")
        self.btn_dwg_source_folder_browse = QPushButton("Папка чертежей...")
        self.btn_dwg_source_folder_browse.clicked.connect(self._browse_dwg_source_folder)
        dwg_source_row.addWidget(self.dwg_source_folder_edit, stretch=1)
        dwg_source_row.addWidget(self.btn_dwg_source_folder_browse)

        dwg_output_row = QHBoxLayout()
        self.dwg_export_folder_edit = QLineEdit()
        self.dwg_export_folder_edit.setPlaceholderText("Папка для DWG (по умолчанию: <папка чертежей>/DWG)")
        self.btn_dwg_export_folder_browse = QPushButton("Папка DWG...")
        self.btn_dwg_export_folder_browse.clicked.connect(self._browse_dwg_export_folder)
        dwg_output_row.addWidget(self.dwg_export_folder_edit, stretch=1)
        dwg_output_row.addWidget(self.btn_dwg_export_folder_browse)

        dwg_export_btn_row = QHBoxLayout()
        self.btn_export_drawings_dwg = QPushButton("Экспорт CDW -> DWG")
        self.btn_export_drawings_dwg.clicked.connect(self._on_export_drawings_dwg_clicked)
        dwg_export_btn_row.addStretch(1)
        dwg_export_btn_row.addWidget(self.btn_export_drawings_dwg)

        dwg_group_layout.addRow("Папка чертежей:", dwg_source_row)
        dwg_group_layout.addRow("Папка вывода:", dwg_output_row)
        dwg_group_layout.addRow("", dwg_export_btn_row)
        dwg_layout.addWidget(dwg_group)
        dwg_layout.addStretch(1)
        left_tabs.addTab(self.dwg_tab, "Экспорт DWG")

        # Вкладка "QR-коды"
        self.qr_tab = QWidget()
        qr_layout = QVBoxLayout(self.qr_tab)

        qr_form = QFormLayout()
        qr_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.qr_data_edit = QLineEdit()
        self.qr_data_edit.setPlaceholderText("NF;EL=Stoika_srednyaya;PRF=H20;L=2360;ID=000123")

        self.qr_output_edit = QLineEdit()
        self.qr_output_edit.setPlaceholderText("Папка для PNG QR (по умолчанию — папка проекта)")

        self.qr_scale_edit = QLineEdit("10")
        self.qr_border_edit = QLineEdit("4")

        qr_form.addRow("Данные QR:", self.qr_data_edit)
        qr_form.addRow("Папка вывода:", self.qr_output_edit)
        qr_form.addRow("Scale (px/модуль):", self.qr_scale_edit)
        qr_form.addRow("Quiet zone (модулей):", self.qr_border_edit)

        qr_layout.addLayout(qr_form)

        qr_buttons_row = QHBoxLayout()
        btn_qr_browse = QPushButton("Выбрать папку...")
        btn_qr_browse.clicked.connect(self._browse_qr_folder)
        btn_qr_generate = QPushButton("Создать QR PNG")
        btn_qr_generate.clicked.connect(self._on_generate_qr_clicked)
        qr_buttons_row.addWidget(btn_qr_browse)
        qr_buttons_row.addWidget(btn_qr_generate)

        qr_layout.addLayout(qr_buttons_row)
        qr_layout.addStretch(1)

        left_tabs.addTab(self.qr_tab, "QR-коды")

        splitter.addWidget(left_tabs)

        # Правая часть: логи
        log_container = QWidget()
        log_layout = QVBoxLayout(log_container)
        log_layout.setContentsMargins(0, 0, 0, 0)

        lbl_log = QLabel("Лог операций")
        lbl_log.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.log_view.setPlaceholderText("Здесь будут появляться сообщения о работе приложения...")
        self.log_view.setFont(QFont("Consolas", 9))

        log_layout.addWidget(lbl_log)
        log_layout.addWidget(self.log_view)

        log_actions_row = QHBoxLayout()
        self.btn_copy_log = QPushButton("Скопировать лог")
        self.btn_copy_log.clicked.connect(self._copy_log_to_clipboard)
        self.btn_save_log = QPushButton("Сохранить лог...")
        self.btn_save_log.clicked.connect(self._save_log_to_file)
        self.btn_open_log = QPushButton("Открыть лог...")
        self.btn_open_log.clicked.connect(self._open_log_file)
        log_actions_row.addWidget(self.btn_copy_log)
        log_actions_row.addWidget(self.btn_save_log)
        log_actions_row.addWidget(self.btn_open_log)
        log_actions_row.addStretch(1)
        log_layout.addLayout(log_actions_row)

        splitter.addWidget(log_container)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        main_layout.addWidget(splitter, stretch=1)

        # Привязываем Python-логгер к правой панели
        root_logger = logging.getLogger()
        qt_handler = QtTextLogHandler(self.log_view)
        qt_handler.setLevel(logging.INFO)
        root_logger.addHandler(qt_handler)

    def _setup_menu(self) -> None:
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Файл")

        act_open = QAction("Открыть папку проекта...", self)
        act_open.triggered.connect(self._browse_project_folder)
        file_menu.addAction(act_open)

        file_menu.addSeparator()

        act_exit = QAction("Выход", self)
        act_exit.triggered.connect(self.close)
        file_menu.addAction(act_exit)

        help_menu = menubar.addMenu("Справка")
        act_about = QAction("О программе", self)
        act_about.triggered.connect(self._show_about)
        help_menu.addAction(act_about)

        log_menu = menubar.addMenu("Логи")
        act_copy_log = QAction("Скопировать лог (Ctrl+C в окне лога)", self)
        act_copy_log.triggered.connect(self._copy_log_to_clipboard)
        log_menu.addAction(act_copy_log)
        act_save_log = QAction("Сохранить лог в файл...", self)
        act_save_log.triggered.connect(self._save_log_to_file)
        log_menu.addAction(act_save_log)
        act_open_log = QAction("Открыть сохраненный лог...", self)
        act_open_log.triggered.connect(self._open_log_file)
        log_menu.addAction(act_open_log)

    def _setup_statusbar(self) -> None:
        status = QStatusBar(self)
        self.setStatusBar(status)
        status.showMessage("Готов к работе. Выберите папку проекта.")

    @staticmethod
    def _profile_choices() -> list[str]:
        return [
            "Профиль H20.1",
            "Профиль H20",
            "Профиль H21",
            "Профиль H22",
            "Профиль H23",
            "Профиль H24",
            "Профиль DT20",
            "Профиль DT21",
            "Профиль DT22",
            "Профиль DT23",
            "Профиль DT24",
            "Профиль H Hat20",
            "Профиль H Hat21",
            "Профиль H Hat22",
            "Профиль H Hat23",
            "Профиль H Hat24",
            "Профиль T12",
            "Профиль T15",
            "Профиль T16",
            "Профиль T25",
            "Профиль T26",
            "Профиль T20",
            "Профиль T21",
            "Профиль T Hat20",
            "Профиль T11N",
            "Профиль L11N",
            "Профиль L20",
            "Профиль L15",
            "Профиль L16",
        ]

    def _fill_profile_combo(self, combo: QComboBox) -> None:
        combo.clear()
        combo.addItems(self._profile_choices())

    def _make_person_combo(self, suggestions: list[str]) -> QComboBox:
        c = QComboBox()
        c.setEditable(True)
        c.addItem("")
        for s in suggestions:
            if s:
                c.addItem(s)
        le = c.lineEdit()
        if le is not None:
            le.setPlaceholderText("Фамилия или выбор из списка")
        return c

    @staticmethod
    def _transliterate_latin_token_to_ru(token: str) -> str:
        """
        Мягкая транслитерация латиницы в кириллицу для заголовков блоков.
        Не зависит от фиксированного набора имен блоков.
        """
        s = token
        low = s.lower()
        pairs = [
            ("shch", "щ"),
            ("zh", "ж"),
            ("kh", "х"),
            ("ts", "ц"),
            ("ch", "ч"),
            ("sh", "ш"),
            ("yu", "ю"),
            ("ya", "я"),
            ("yo", "ё"),
        ]
        out = ""
        i = 0
        while i < len(low):
            matched = False
            for src, dst in pairs:
                if low.startswith(src, i):
                    out += dst
                    i += len(src)
                    matched = True
                    break
            if matched:
                continue
            ch = low[i]
            one = {
                "a": "а", "b": "б", "c": "к", "d": "д", "e": "е",
                "f": "ф", "g": "г", "h": "х", "i": "и", "j": "й",
                "k": "к", "l": "л", "m": "м", "n": "н", "o": "о",
                "p": "п", "q": "к", "r": "р", "s": "с", "t": "т",
                "u": "у", "v": "в", "w": "в", "x": "кс", "y": "ы", "z": "з",
                "_": " ",
            }.get(ch, ch)
            out += one
            i += 1
        return out.strip().capitalize()

    def _humanize_block_title(self, block_id: str) -> str:
        txt = (block_id or "").strip()
        if not txt:
            return "Прочие"
        # split camel case + underscores
        import re as _re
        txt = _re.sub(r"(?<!^)([A-ZА-Я])", r" \1", txt).replace("_", " ")
        txt = " ".join(txt.split())
        # Если токен латиницей, пробуем транслитерировать.
        parts = []
        for t in txt.split(" "):
            if t and all(("a" <= c.lower() <= "z") or c.isdigit() for c in t):
                parts.append(self._transliterate_latin_token_to_ru(t))
            else:
                parts.append(t.capitalize())
        return " ".join(parts)

    def _sync_assembly_and_stamp_fields(self) -> None:
        """Заполнить поля обозначения/наименования из текущей сборки и продублировать в штамп."""
        if not self._assembly_info:
            return
        des = self._assembly_info.designation or ""
        name = self._assembly_info.name or ""
        self.assembly_designation_edit.setText(des)
        self.assembly_name_edit.setText(name)
        self.stamp_designation_edit.setText(des)
        self.stamp_name_edit.setText(name)

    # ------------------------------------------------------------------
    # Обработчики
    # ------------------------------------------------------------------

    def _browse_project_folder(self) -> None:
        """Выбор папки проекта NordFox."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку проекта NordFox (сборка с .a3d)",
            "",
        )
        if not folder:
            return

        self._current_project_root = Path(folder)
        self.project_path_edit.setText(str(self._current_project_root))
        if not self.copy_target_edit.text().strip():
            self.copy_target_edit.setText(str(self._current_project_root.parent))
        self.statusBar().showMessage(f"Выбрана папка проекта: {self._current_project_root}")
        logger.info(f"Папка проекта: {self._current_project_root}")

        # Инициализируем JSON-лог для этой сессии/проекта
        from ..__version__ import __version__
        self._json_log = JsonLogStore(
            app_name="NordFox Module Manager",
            app_version=__version__,
            project_root=self._current_project_root,
        )

        # Сканируем проект и строим динамическую форму переменных
        try:
            assembly_info, documents, var_index = scan_project(self._current_project_root, self._kompas)
            self._assembly_info = assembly_info
            self._documents = documents
            self._var_index = var_index
            self._rebuild_variables_form()
            self._sync_assembly_and_stamp_fields()

            # Сохраняем сводное состояние в JSON-лог
            if self._json_log:
                state = {
                    "assembly_file": str(assembly_info.path),
                    "documents": [str(d.path) for d in documents],
                    "variables_index": {name: kv.value for name, kv in var_index.items()},
                }
                self._json_log.set_project_state(state)

            self.btn_update_variables.setEnabled(True)
            self.btn_rescan.setEnabled(True)
            self.btn_copy_and_update.setEnabled(True)
            QMessageBox.information(
                self,
                "Проект просканирован",
                f"Найдена сборка: {assembly_info.path.name}\n"
                f"Деталей: {len([d for d in documents if d.doc_type == 'part'])}\n"
                f"Уникальных переменных: {len(var_index)}",
            )
        except Exception as exc:
            logger.error(f"Ошибка сканирования проекта: {exc}")
            QMessageBox.critical(
                self,
                "Ошибка сканирования",
                f"Не удалось просканировать проект:\n{exc}",
            )

    def _rebuild_variables_form(self) -> None:
        """Перестроить форму переменных на вкладке 'Переменные'."""
        # Очищаем предыдущие вкладки с блоками
        while self.vars_toolbox.count() > 0:
            page = self.vars_toolbox.widget(0)
            self.vars_toolbox.removeItem(0)
            if page is not None:
                page.deleteLater()
        self._var_inputs.clear()
        self._drawing_comment_inputs.clear()

        # Группируем переменные сборки по block_id
        blocks: Dict[str, Dict[str, KompasVariable]] = {}
        ungrouped: Dict[str, KompasVariable] = {}

        logger.info(
            "Перестройка формы переменных: всего переменных в var_index=%d (должны быть только из сборки)",
            len(self._var_index),
        )

        for name, kv in self._var_index.items():
            if kv.document_type != "assembly":
                continue
            normalized_block = (kv.block_id or "").rstrip("_").strip()
            if normalized_block:
                blocks.setdefault(normalized_block, {})[name] = kv
            else:
                ungrouped[name] = kv

        # Создаем блоки
        for block_id, vars_in_block in sorted(blocks.items(), key=lambda x: x[0].lower()):
            logger.info("[UI] Блок переменных: %s", block_id)
            block_title = self._humanize_block_title(block_id)
            group = QGroupBox(block_title)
            group.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            group.setStyleSheet("QGroupBox { font-weight: bold; }")

            vbox = QVBoxLayout(group)

            # Формы переменных блока
            form = QFormLayout()
            form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

            for name, kv in sorted(vars_in_block.items(), key=lambda x: x[0].lower()):
                # Пустую переменную-заголовок блока (Stoiki____) в списке не показываем
                if kv.is_block_header:
                    continue
                logger.info(
                    "[UI]   поле переменной: name=%s, doc_type=%s, doc_path=%s, block_id=%s, value=%r",
                    name,
                    kv.document_type,
                    kv.document_path,
                    kv.block_id,
                    kv.value,
                )
                label = QLabel(name + ":")
                edit = QLineEdit()
                edit.setText(str(kv.value) if kv.value is not None else "")
                form.addRow(label, edit)
                self._var_inputs[name] = edit

            vbox.addLayout(form)
            self.vars_toolbox.addItem(group, block_title)

        # Блок "Прочие" для переменных без блока
        if ungrouped:
            logger.info("[UI] Блок переменных: Прочие")
            group = QGroupBox("Прочие")
            vbox = QVBoxLayout(group)
            form = QFormLayout()
            form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
            for name, kv in sorted(ungrouped.items(), key=lambda x: x[0].lower()):
                logger.info(
                    "[UI]   поле переменной: name=%s, doc_type=%s, doc_path=%s, block_id=%s, value=%r",
                    name,
                    kv.document_type,
                    kv.document_path,
                    kv.block_id,
                    kv.value,
                )
                label = QLabel(name + ":")
                edit = QLineEdit()
                edit.setText(str(kv.value) if kv.value is not None else "")
                form.addRow(label, edit)
                self._var_inputs[name] = edit
            vbox.addLayout(form)
            self.vars_toolbox.addItem(group, "Прочие")

        # Отдельный блок для переменных чертежей (комментарии)
        drawing_vars: Dict[str, KompasVariable] = {}
        for doc in self._documents:
            if doc.doc_type != "drawing":
                continue
            for name, kv in doc.variables.items():
                # Пустую переменную-заголовок блока (Переменные_чертежа____) не показываем
                if kv.is_block_header:
                    continue
                # Если одна и та же переменная встречается в нескольких чертежах,
                # берём первую как эталон для комментария.
                if name not in drawing_vars:
                    drawing_vars[name] = kv

        if drawing_vars:
            logger.info("[UI] Блок переменных: Переменные чертежа")
            group = QGroupBox("Переменные чертежа")
            group.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            group.setStyleSheet("QGroupBox { font-weight: bold; }")

            vbox = QVBoxLayout(group)
            form = QFormLayout()
            form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

            for name, kv in sorted(drawing_vars.items(), key=lambda x: x[0].lower()):
                label = QLabel(name + " (комментарий):")
                edit = QLineEdit()
                edit.setText(kv.comment or "")
                form.addRow(label, edit)
                self._drawing_comment_inputs[name] = edit
                logger.info(
                    "[UI]   поле комментария: name=%s, doc_type=%s, doc_path=%s, block_id=%s, comment=%r",
                    name,
                    kv.document_type,
                    kv.document_path,
                    kv.block_id,
                    kv.comment,
                )

            vbox.addLayout(form)
            self.vars_toolbox.addItem(group, "Переменные чертежа")

        # После перестройки формы переменных обновляем и таблицу деталей/подсборок
        self._rebuild_items_table()

    def _rebuild_items_table(self) -> None:
        """
        Построить таблицу деталей/подсборок:
        - используем только документы текущего уровня (assembly_info + parts из self._documents);
        - каждая строка соответствует конкретному файлу (без дедупликации по mark/name).
        """
        self._assembly_items.clear()
        self._assembly_item_new_marking.clear()
        self._assembly_item_new_name.clear()

        if not self._assembly_info:
            self.items_table.setRowCount(0)
            return

        # Собираем элементы: сборка + детали
        def add_item(doc: KompasDocumentInfo) -> None:
            if doc.doc_type not in ("assembly", "part"):
                return
            key = doc.path
            if key not in self._assembly_items:
                self._assembly_items[key] = doc

        add_item(self._assembly_info)
        for d in self._documents:
            add_item(d)

        logger.info(
            "[Items] Уникальных элементов для таблицы: %d",
            len(self._assembly_items),
        )

        self.items_table.setRowCount(len(self._assembly_items))

        for row, (doc_path, doc) in enumerate(self._assembly_items.items()):
            doc_type = doc.doc_type
            mark = doc.designation or ""
            name = doc.name or ""

            # Тип
            if doc_type == "assembly":
                type_label = "Сборка"
            elif doc_type == "part":
                type_label = "Деталь"
            else:
                type_label = doc_type
            type_item = QTableWidgetItem(type_label)
            type_item.setFlags(type_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.items_table.setItem(row, 0, type_item)

            # Текущее обозначение
            cur_mark_item = QTableWidgetItem(mark)
            cur_mark_item.setFlags(cur_mark_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.items_table.setItem(row, 1, cur_mark_item)

            # Текущее наименование
            cur_name_item = QTableWidgetItem(name)
            cur_name_item.setFlags(cur_name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.items_table.setItem(row, 2, cur_name_item)

            # Новое обозначение
            new_mark_edit = QLineEdit()
            self.items_table.setCellWidget(row, 3, new_mark_edit)
            self._assembly_item_new_marking[doc_path] = new_mark_edit

            # Новое наименование
            new_name_edit = QLineEdit()
            self.items_table.setCellWidget(row, 4, new_name_edit)
            self._assembly_item_new_name[doc_path] = new_name_edit

            # Файл
            file_item = QTableWidgetItem(str(doc.path.name))
            file_item.setFlags(file_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.items_table.setItem(row, 5, file_item)

            logger.info(
                "[Items] Строка %d: type=%s, mark=%r, name=%r, file=%s",
                row,
                type_label,
                mark,
                name,
                doc.path,
            )

    def _fill_item_metadata_by_rules(self) -> Dict[str, int]:
        """
        Заполнить колонки «Новое обозначение/наименование» по ролям
        и текущим (уже пересчитанным) значениям переменных сборки.
        """
        profile = self.items_profile_combo.currentText().strip()
        try:
            series = int(self.items_series_combo.currentText())
        except ValueError:
            series = 1
        var_values = collect_assembly_numeric_values(self._var_index)
        filled_mark = 0
        filled_name = 0
        for doc_path, mark_edit in self._assembly_item_new_marking.items():
            doc = self._assembly_items.get(doc_path)
            if doc is None:
                continue
            doc_type = doc.doc_type
            cur_mark = doc.designation or ""
            name = doc.name or ""
            doc_name = doc.path.name if doc is not None else "<unknown>"
            # КРИТИЧНО: сначала определяем роль по имени файла детали (более стабильно),
            # затем по name, и только в конце — по префиксу marking.
            if doc is not None:
                role = infer_role_from_part_name(doc.path.stem)
            else:
                role = None
            if not role:
                role = infer_role_from_part_name(name or "")
            if not role and cur_mark:
                up = cur_mark.upper()
                if up.startswith("Р-"):
                    role = "rigel"
                elif up.startswith("СС-"):
                    role = "stoika_srednyaya"
                elif up.startswith("СК-"):
                    role = "stoika_kraynaya"
            if not role:
                logger.info(
                    "[Items][Rules] skip: file=%s type=%s cur_mark=%r name=%r role=not-detected",
                    doc_name,
                    doc_type,
                    cur_mark,
                    name,
                )
                continue
            logger.info(
                "[Items][Rules] role: file=%s type=%s cur_mark=%r name=%r -> role=%s",
                doc_name,
                doc_type,
                cur_mark,
                name,
                role,
            )
            if profile:
                des = build_element_designation(role, series, profile, var_values)
                if des:
                    mark_edit.setText(des)
                    filled_mark += 1
                    logger.info(
                        "[Items][Rules] marking: file=%s role=%s -> %r",
                        doc_name,
                        role,
                        des,
                    )
                else:
                    logger.info(
                        "[Items][Rules] marking-skip: file=%s role=%s reason=no-designation",
                        doc_name,
                        role,
                    )
            name_edit = self._assembly_item_new_name.get(doc_path)
            if name_edit is not None:
                new_name = build_element_name(role, profile, var_values) if profile else None
                if new_name:
                    m = re.search(r"[НH]\d+(?:\.\d+)?", name or "", flags=re.IGNORECASE)
                    if m:
                        cur_profile = m.group(0).upper().replace("H", "Н")
                        new_name = re.sub(
                            r"Профиль\s+[НH]\d+(?:\.\d+)?",
                            f"Профиль {cur_profile}",
                            new_name,
                            flags=re.IGNORECASE,
                        )
                    name_edit.setText(new_name)
                    filled_name += 1
                    logger.info(
                        "[Items][Rules] name: file=%s role=%s -> %r",
                        doc_name,
                        role,
                        new_name,
                    )
                else:
                    logger.info(
                        "[Items][Rules] name-skip: file=%s role=%s reason=no-name",
                        doc_name,
                        role,
                    )
        # Сборка: прокидываем поля из шапки в колонку "новое", чтобы применялось в файл и UI.
        for doc_path, mark_edit in self._assembly_item_new_marking.items():
            doc = self._assembly_items.get(doc_path)
            if doc is None or doc.doc_type != "assembly":
                continue
            asm_mark = self.assembly_designation_edit.text().strip()
            asm_name = self.assembly_name_edit.text().strip()
            if asm_mark:
                mark_edit.setText(asm_mark)
                filled_mark += 1
                logger.info("[Items][Rules] assembly marking -> %r", asm_mark)
            name_edit = self._assembly_item_new_name.get(doc_path)
            if name_edit is not None and asm_name:
                name_edit.setText(asm_name)
                filled_name += 1
                logger.info("[Items][Rules] assembly name -> %r", asm_name)
            break
        logger.info(
            "[Items] Автозаполнение по правилам: обозначений=%d, наименований=%d",
            filled_mark,
            filled_name,
        )
        return {"markings": filled_mark, "names": filled_name}

    def _on_apply_designation_rules(self) -> None:
        """Заполнить колонки «Новое обозначение/наименование» по правилам."""
        if not self._assembly_info:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите и просканируйте проект.")
            return
        profile = self.items_profile_combo.currentText().strip()
        if not profile:
            QMessageBox.warning(self, "Профиль", "Выберите профиль для правил обозначения/наименования.")
            return
        filled = self._fill_item_metadata_by_rules()
        self.statusBar().showMessage(
            f"Подставлено по правилам: обозначений {filled['markings']}, наименований {filled['names']}.",
            5000,
        )
        if filled["markings"] == 0 and filled["names"] == 0:
            QMessageBox.information(
                self,
                "Нет совпадений",
                "Не удалось сопоставить детали с ролями СК/СС/Р "
                "или не хватает длины в переменных сборки.",
            )

    def _collect_item_metadata_payload(self) -> Dict[Path, Dict[str, str]]:
        """
        Собрать непустые значения из колонок «Новое обозначение/наименование».
        Возвращает: путь_документа -> {"marking": "...", "name": "..."}.
        """
        updates: Dict[Path, Dict[str, str]] = {}
        for doc_path, doc in self._assembly_items.items():
            mark_edit = self._assembly_item_new_marking.get(doc_path)
            name_edit = self._assembly_item_new_name.get(doc_path)
            if not mark_edit and not name_edit:
                continue
            new_mark = (mark_edit.text().strip() if mark_edit else "")
            new_name = (name_edit.text().strip() if name_edit else "")
            if not new_mark and not new_name:
                continue
            payload: Dict[str, str] = {}
            if new_mark:
                payload["marking"] = new_mark
            if new_name:
                payload["name"] = new_name
            updates[doc.path] = payload
        return updates

    def _apply_item_metadata_updates(
        self,
        updates: Dict[Path, Dict[str, str]],
        *,
        sync_assembly_components: bool = False,
    ) -> Dict[str, object]:
        """
        Записать обозначение/наименование в документы 3D.
        """
        def _read_part_prop(i_part: object, low_name: str, pascal_name: str) -> object:
            value = None
            try:
                value = getattr(i_part, low_name, None)
            except Exception:
                value = None
            if value in (None, ""):
                try:
                    value = getattr(i_part, pascal_name, None)
                except Exception:
                    value = None
            return value

        def _write_part_prop(i_part: object, low_name: str, pascal_name: str, value: str) -> tuple[bool, str]:
            wrote = False
            used = ""
            try:
                setattr(i_part, low_name, value)
                wrote = True
                used = low_name
            except Exception:
                pass
            try:
                setattr(i_part, pascal_name, value)
                wrote = True
                used = f"{used}+{pascal_name}" if used else pascal_name
            except Exception:
                pass
            return wrote, used

        def _sync_assembly_component_metadata(
            assembly_path: Path,
            updates_by_path: Dict[Path, Dict[str, str]],
        ) -> tuple[int, int, list[str]]:
            local_errors: list[str] = []
            logger.info("[Assembly][probe] sync func enter: assembly=%s, updates=%d", assembly_path, len(updates_by_path))
            by_old_marking: Dict[str, Dict[str, str]] = {}
            by_old_pair: Dict[tuple[str, str], Dict[str, str]] = {}
            # ISOLATED MODE: берём только изменённые позиции и не обходим всю таблицу элементов.
            # Это снижает риск сбоя в ранней фазе sync.
            for doc_path, payload in updates_by_path.items():
                doc = self._assembly_items.get(doc_path)
                if not doc or doc.doc_type != "part" or doc_path.suffix.lower() != ".m3d":
                    continue
                old_mark = str(doc.marking or "").strip()
                old_name = str(doc.name or "").strip()
                if old_mark:
                    by_old_marking.setdefault(old_mark, payload)
                if old_mark or old_name:
                    by_old_pair.setdefault((old_mark, old_name), payload)
            logger.info(
                "[Assembly][probe] maps prepared: by_old_marking=%d, by_old_pair=%d",
                len(by_old_marking),
                len(by_old_pair),
            )

            def _iter_components_probe(root_part: object, max_top: int = 12):
                # DIAG MODE: только верхний уровень и короткий лимит, чтобы локализовать crash.
                logger.info("[Assembly][probe] iterate top-level start, max_top=%d", max_top)
                for idx in range(max_top):
                    logger.info("[Assembly][probe] GetPart(%d): before", idx)
                    try:
                        child = root_part.GetPart(idx)
                    except Exception as exc:
                        logger.warning("[Assembly][probe] GetPart(%d): exception: %s", idx, exc)
                        break
                    if not child:
                        logger.info("[Assembly][probe] GetPart(%d): empty, stop", idx)
                        break
                    logger.info("[Assembly][probe] GetPart(%d): ok", idx)
                    yield child
                logger.info("[Assembly][probe] iterate top-level done")

            def _component_tag(comp: object) -> str:
                # SAFE MODE: не обращаемся к потенциально нестабильным FileName/FullName.
                cur_name = str(_read_part_prop(comp, "name", "Name") or "").strip()
                cur_mark = str(_read_part_prop(comp, "marking", "Marking") or "").strip()
                if cur_name:
                    return cur_name
                if cur_mark:
                    return cur_mark
                return "<component>"

            def _payload_for_component(comp: object) -> tuple[Dict[str, str] | None, str]:
                cur_mark = str(_read_part_prop(comp, "marking", "Marking") or "").strip()
                cur_name = str(_read_part_prop(comp, "name", "Name") or "").strip()
                by_mark = by_old_marking.get(cur_mark)
                if by_mark:
                    return by_mark, f"old-marking={cur_mark}"
                by_old = by_old_pair.get((cur_mark, cur_name))
                if by_old:
                    return by_old, f"old-pair={cur_mark}|{cur_name}"
                return None, "no-match"

            logger.info("[Assembly][probe] open assembly: %s", assembly_path)
            if not self._kompas.open_document(str(assembly_path)):
                return 0, 0, [f"Не удалось открыть сборку для синхронизации компонентов: {assembly_path}"]
            logger.info("[Assembly][probe] open assembly: ok")

            components_updated = 0
            fields_updated_local = 0
            save_on_close_local = False
            visited_components = 0
            matched_components = 0
            try:
                logger.info("[Assembly][probe] api get start")
                api5 = self._kompas.get_api5()
                api7 = self._kompas.get_api7()
                if api5 is None or api7 is None:
                    return 0, 0, ["API5/API7 недоступны для синхронизации компонентов сборки"]
                logger.info("[Assembly][probe] api get ok")

                logger.info("[Assembly][probe] ActiveDocument3D read start")
                i_doc3d = getattr(api5, "ActiveDocument3D", None)
                if not i_doc3d:
                    return 0, 0, ["ActiveDocument3D не найден для синхронизации компонентов сборки"]
                logger.info("[Assembly][probe] ActiveDocument3D read ok")
                logger.info("[Assembly][probe] GetPart(-1) start")
                i_asm = i_doc3d.GetPart(-1)
                if not i_asm:
                    return 0, 0, ["GetPart(-1) для сборки вернул пустой объект"]
                logger.info("[Assembly][probe] GetPart(-1) ok")

                for comp in _iter_components_probe(i_asm):
                    visited_components += 1
                    payload, match_reason = _payload_for_component(comp)
                    if not payload:
                        continue
                    matched_components += 1
                    comp_tag = _component_tag(comp)

                    changed_here = 0
                    if "marking" in payload:
                        old_mark = _read_part_prop(comp, "marking", "Marking")
                        wrote = False
                        used_prop = ""
                        try:
                            comp.marking = payload["marking"]
                            wrote = True
                            used_prop = "marking"
                        except Exception:
                            wrote, used_prop = _write_part_prop(comp, "marking", "Marking", payload["marking"])
                        if not wrote:
                            local_errors.append(f"[Assembly] {comp_tag}: не удалось записать marking")
                        else:
                            changed_here += 1
                            if old_mark != payload["marking"]:
                                logger.info(
                                    "[Assembly] %s [%s]: marking(%s): %r -> %r",
                                    comp_tag,
                                    match_reason,
                                    used_prop,
                                    old_mark,
                                    payload["marking"],
                                )
                            else:
                                logger.info(
                                    "[Assembly] %s [%s]: marking(%s): forced overwrite %r",
                                    comp_tag,
                                    match_reason,
                                    used_prop,
                                    payload["marking"],
                                )

                    if "name" in payload:
                        old_name = _read_part_prop(comp, "name", "Name")
                        wrote = False
                        used_prop = ""
                        try:
                            comp.name = payload["name"]
                            wrote = True
                            used_prop = "name"
                        except Exception:
                            wrote, used_prop = _write_part_prop(comp, "name", "Name", payload["name"])
                        if not wrote:
                            local_errors.append(f"[Assembly] {comp_tag}: не удалось записать name")
                        else:
                            changed_here += 1
                            if old_name != payload["name"]:
                                logger.info(
                                    "[Assembly] %s [%s]: name(%s): %r -> %r",
                                    comp_tag,
                                    match_reason,
                                    used_prop,
                                    old_name,
                                    payload["name"],
                                )
                            else:
                                logger.info(
                                    "[Assembly] %s [%s]: name(%s): forced overwrite %r",
                                    comp_tag,
                                    match_reason,
                                    used_prop,
                                    payload["name"],
                                )

                    if changed_here > 0:
                        try:
                            comp.Update()
                        except Exception:
                            pass
                        try:
                            comp.RebuildModel()
                        except Exception:
                            pass
                        fields_updated_local += changed_here
                        components_updated += 1
                        save_on_close_local = True

                if save_on_close_local:
                    try:
                        i_asm.Update()
                    except Exception:
                        pass
                    try:
                        i_doc3d.RebuildDocument()
                    except Exception:
                        pass
                    time.sleep(0.3)
                    try:
                        api7.ActiveDocument.Save()
                    except Exception as exc:
                        local_errors.append(f"Сборка: ошибка Save после синхронизации вхождений: {exc}")
                    time.sleep(0.3)
                logger.info(
                    "[Assembly] sync stats: visited=%d, matched=%d, updated=%d, fields=%d",
                    visited_components,
                    matched_components,
                    components_updated,
                    fields_updated_local,
                )
            except Exception as exc:
                local_errors.append(f"Синхронизация компонентов сборки: {exc}")
            finally:
                self._kompas.close_active_document(save=save_on_close_local)

            # SAFE MODE: пропускаем повторный глубокий reopen-check, чтобы избежать COM-crash.
            # Фактическая проверка будет выполнена стандартным пересканом проекта после операции.

            return components_updated, fields_updated_local, local_errors

        result: Dict[str, object] = {
            "success": False,
            "documents_updated": 0,
            "fields_updated": 0,
            "errors": [],
        }
        errors: list[str] = []
        docs_updated = 0
        fields_updated = 0

        # Шаг 1: сначала синхронизируем вхождения деталей внутри сборки (.a3d),
        # чтобы состояние сборки фиксировалось до прохода по файлам деталей.
        if sync_assembly_components and self._assembly_info and self._assembly_info.path.exists():
            part_updates: Dict[Path, Dict[str, str]] = {
                p: payload
                for p, payload in updates.items()
                if p.suffix.lower() == ".m3d"
            }
            if part_updates:
                logger.info(
                    "[Assembly] sync start: assembly=%s, part_updates=%d",
                    self._assembly_info.path,
                    len(part_updates),
                )
                comp_docs, comp_fields, comp_errors = _sync_assembly_component_metadata(
                    self._assembly_info.path,
                    part_updates,
                )
                docs_updated += comp_docs
                fields_updated += comp_fields
                errors.extend(comp_errors)

        # Шаг 2: обновляем сами документы деталей.
        for path, payload in updates.items():
            if not self._kompas.open_document(str(path)):
                errors.append(f"Не удалось открыть документ: {path}")
                continue
            save_on_close = False
            try:
                api5 = self._kompas.get_api5()
                api7 = self._kompas.get_api7()
                if api5 is None or api7 is None:
                    errors.append("API5/API7 недоступны для обновления обозначения/наименования")
                    continue
                i_doc3d = getattr(api5, "ActiveDocument3D", None)
                if not i_doc3d:
                    errors.append(f"{path.name}: ActiveDocument3D не найден")
                    continue
                i_part = i_doc3d.GetPart(-1)
                changed_here = 0

                if "marking" in payload:
                    old_mark = _read_part_prop(i_part, "marking", "Marking")
                    wrote = False
                    used_prop = ""
                    try:
                        i_part.marking = payload["marking"]
                        wrote = True
                        used_prop = "marking"
                    except Exception:
                        wrote, used_prop = _write_part_prop(i_part, "marking", "Marking", payload["marking"])
                    if not wrote:
                        errors.append(f"{path.name}: не удалось записать marking в COM-свойства part")
                    else:
                        changed_here += 1
                        if old_mark != payload["marking"]:
                            logger.info(
                                "%s: marking(%s): %r -> %r",
                                path.name,
                                used_prop,
                                old_mark,
                                payload["marking"],
                            )
                        else:
                            logger.info(
                                "%s: marking(%s): forced overwrite %r",
                                path.name,
                                used_prop,
                                payload["marking"],
                            )
                if "name" in payload:
                    old_name = _read_part_prop(i_part, "name", "Name")
                    wrote = False
                    used_prop = ""
                    try:
                        i_part.name = payload["name"]
                        wrote = True
                        used_prop = "name"
                    except Exception:
                        wrote, used_prop = _write_part_prop(i_part, "name", "Name", payload["name"])
                    if not wrote:
                        errors.append(f"{path.name}: не удалось записать name в COM-свойства part")
                    else:
                        changed_here += 1
                        if old_name != payload["name"]:
                            logger.info(
                                "%s: name(%s): %r -> %r",
                                path.name,
                                used_prop,
                                old_name,
                                payload["name"],
                            )
                        else:
                            logger.info(
                                "%s: name(%s): forced overwrite %r",
                                path.name,
                                used_prop,
                                payload["name"],
                            )

                if changed_here > 0:
                    try:
                        i_part.Update()
                    except Exception:
                        pass
                    try:
                        i_part.RebuildModel()
                    except Exception:
                        pass
                    try:
                        i_doc3d.RebuildDocument()
                    except Exception:
                        pass
                    time.sleep(0.2)
                    api7.ActiveDocument.Save()
                    time.sleep(0.2)
                    save_on_close = True
                    # Read-back в той же сессии документа: подтверждаем фактическую запись.
                    rb_mark = _read_part_prop(i_part, "marking", "Marking")
                    rb_name = _read_part_prop(i_part, "name", "Name")
                    if "marking" in payload and rb_mark != payload["marking"]:
                        errors.append(
                            f"{path.name}: marking read-back mismatch: {rb_mark!r} != {payload['marking']!r}"
                        )
                    if "name" in payload and rb_name != payload["name"]:
                        errors.append(
                            f"{path.name}: name read-back mismatch: {rb_name!r} != {payload['name']!r}"
                        )
                    docs_updated += 1
                    fields_updated += changed_here
            except Exception as exc:
                errors.append(f"{path.name}: ошибка обновления обозначения/наименования: {exc}")
            finally:
                self._kompas.close_active_document(save=save_on_close)

        result["documents_updated"] = docs_updated
        result["fields_updated"] = fields_updated
        result["errors"] = errors
        result["success"] = len(errors) == 0
        return result

    def _on_apply_items_meta_clicked(self) -> None:
        updates = self._collect_item_metadata_payload()
        if not updates:
            QMessageBox.information(
                self,
                "Нет данных",
                "Заполните хотя бы одно поле в колонках «Новое обозначение» / «Новое наименование».",
            )
            return
        result = self._apply_item_metadata_updates(updates, sync_assembly_components=False)
        # COM в КОМПАС может падать при плотной серии операций. Небольшая пауза
        # перед isolated subprocess заметно снижает вероятность аварии UI-процесса.
        time.sleep(0.8)
        # После записи в детали/сборку-файл дополнительно синхронизируем вхождения в .a3d
        # через изолированный subprocess (безопаснее для UI-процесса).
        sync_result = self._run_isolated_assembly_sync(updates)

        apply_errors = list(result.get("errors") or [])
        sync_errors = list(sync_result.get("errors") or [])
        all_errors = apply_errors + sync_errors

        if result.get("success") and sync_result.get("success"):
            self.statusBar().showMessage(
                f"Обозн./наимен. применены: документов={result.get('documents_updated')}, "
                f"полей={result.get('fields_updated')}; "
                f"вхождения сборки: документов={sync_result.get('documents_updated')}, "
                f"полей={sync_result.get('fields_updated')}",
                5000,
            )
        else:
            QMessageBox.warning(
                self,
                "Частичные ошибки",
                "\n".join(str(e) for e in all_errors) if all_errors else "Операция завершена с ошибками.",
            )
        self._on_rescan_clicked()

    def _on_sync_assembly_components_clicked(self) -> None:
        """Отдельно синхронизировать вхождения деталей в сборке (.a3d)."""
        updates = self._collect_item_metadata_payload()
        if not updates:
            # Если поля "новых" значений пусты, синхронизируем по текущим значениям деталей.
            # Это позволяет отдельно "протолкнуть" marking/name в вхождения сборки без ручного ввода.
            updates = {}
            for doc_path, doc in self._assembly_items.items():
                if doc.doc_type != "part" or doc_path.suffix.lower() != ".m3d":
                    continue
                cur_mark = str(doc.marking or "").strip()
                cur_name = str(doc.name or "").strip()
                payload: Dict[str, str] = {}
                if cur_mark:
                    payload["marking"] = cur_mark
                if cur_name:
                    payload["name"] = cur_name
                if payload:
                    updates[doc_path] = payload
            if not updates:
                QMessageBox.information(
                    self,
                    "Нет данных",
                    "Нет значений marking/name для синхронизации вхождений сборки.",
                )
                return
            logger.info("[Assembly][isolated] fallback updates prepared: %d", len(updates))
        result = self._run_isolated_assembly_sync(updates)
        if result.get("success"):
            self.statusBar().showMessage(
                f"Синхронизация сборки выполнена: документов={result.get('documents_updated')}, "
                f"полей={result.get('fields_updated')}",
                5000,
            )
            self._on_rescan_clicked()
        else:
            errors = result.get("errors") or []
            QMessageBox.warning(
                self,
                "Частичные ошибки синхронизации",
                "\n".join(str(e) for e in errors),
            )

    def _run_isolated_assembly_sync(self, updates: Dict[Path, Dict[str, str]]) -> Dict[str, object]:
        """
        Изолированный sync вхождений сборки в отдельном процессе.
        Даже если в subprocess произойдёт COM-crash, UI-процесс останется жив.
        """
        result: Dict[str, object] = {
            "success": False,
            "documents_updated": 0,
            "fields_updated": 0,
            "errors": [],
        }
        errors: list[str] = []
        if not self._assembly_info or not self._assembly_info.path.exists():
            errors.append("Сборка не найдена для синхронизации.")
            result["errors"] = errors
            return result

        items_payload: list[dict[str, str]] = []
        for doc_path, payload in updates.items():
            if doc_path.suffix.lower() != ".m3d":
                continue
            doc = self._assembly_items.get(doc_path)
            if not doc:
                continue
            old_mark = str(doc.designation or "").strip()
            old_name = str(doc.name or "").strip()
            new_mark = str(payload.get("marking", "") or "").strip()
            new_name = str(payload.get("name", "") or "").strip()
            if not new_mark and not new_name:
                continue
            if not old_mark and not old_name:
                continue
            items_payload.append(
                {
                    "old_marking": old_mark,
                    "old_name": old_name,
                    "new_marking": new_mark,
                    "new_name": new_name,
                    "doc_path": str(doc_path),
                    "file_stem": doc_path.stem,
                }
            )

        if not items_payload:
            result["success"] = True
            return result

        runner = Path(__file__).resolve().parents[1] / "core" / "assembly_sync_subprocess.py"
        if not runner.exists():
            errors.append(f"Не найден скрипт sync: {runner}")
            result["errors"] = errors
            return result

        payload_data = {
            "assembly_path": str(self._assembly_info.path),
            "items": items_payload,
        }
        try:
            with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as pf:
                json.dump(payload_data, pf, ensure_ascii=False)
                payload_file = Path(pf.name)
            with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as rf:
                rf.write("{}")
                result_file = Path(rf.name)

            cmd = [
                sys.executable,
                str(runner),
                "--payload",
                str(payload_file),
                "--result",
                str(result_file),
            ]
            logger.info("[Assembly][isolated] run: %s", " ".join(cmd))
            completed = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
            if completed.returncode != 0:
                stderr = (completed.stderr or "").strip()
                stdout = (completed.stdout or "").strip()
                msg = f"Isolated sync завершился с кодом {completed.returncode}."
                if stderr:
                    msg += f" stderr: {stderr}"
                elif stdout:
                    msg += f" stdout: {stdout}"
                errors.append(msg)
            try:
                result_data = json.loads(result_file.read_text(encoding="utf-8"))
            except Exception as exc:
                result_data = {"success": False, "errors": [f"Не удалось прочитать результат isolated sync: {exc}"]}

            result["success"] = bool(result_data.get("success", False)) and not errors
            result["documents_updated"] = int(result_data.get("documents_updated", 0) or 0)
            result["fields_updated"] = int(result_data.get("fields_updated", 0) or 0)
            result["errors"] = errors + list(result_data.get("errors") or [])
            visited = int(result_data.get("components_visited", 0) or 0)
            matched = int(result_data.get("components_matched", 0) or 0)
            logger.info(
                "[Assembly][isolated] stats: visited=%d, matched=%d, updated=%d, fields=%d",
                visited,
                matched,
                int(result["documents_updated"] or 0),
                int(result["fields_updated"] or 0),
            )
            if result.get("success") and int(result["documents_updated"] or 0) == 0:
                result["success"] = False
                result["errors"] = list(result["errors"]) + [
                    "Синхронизация вхождений не нашла ни одного совпадения в сборке "
                    f"(visited={visited}, matched={matched}).",
                ]
            return result
        except Exception as exc:
            result["errors"] = [f"Ошибка запуска isolated sync: {exc}"]
            return result
        finally:
            for p in ("payload_file", "result_file"):
                fp = locals().get(p)
                if isinstance(fp, Path):
                    try:
                        fp.unlink(missing_ok=True)
                    except Exception:
                        pass

    def _on_rescan_clicked(self) -> None:
        """Пересканировать текущий проект (подхватить внешние изменения)."""
        if not self._current_project_root:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите папку проекта.")
            return

        try:
            assembly_info, documents, var_index = scan_project(self._current_project_root, self._kompas)
            self._assembly_info = assembly_info
            self._documents = documents
            self._var_index = var_index
            self._rebuild_variables_form()
            self._sync_assembly_and_stamp_fields()

            if self._json_log:
                state = {
                    "assembly_file": str(assembly_info.path),
                    "documents": [str(d.path) for d in documents],
                    "variables_index": {name: kv.value for name, kv in var_index.items()},
                }
                self._json_log.set_project_state(state)

            self.statusBar().showMessage("Пересканирование проекта завершено", 5000)
        except Exception as exc:
            logger.error(f"Ошибка пересканирования проекта: {exc}")
            QMessageBox.critical(
                self,
                "Ошибка пересканирования",
                f"Не удалось пересканировать проект:\n{exc}",
            )

    def _browse_qr_folder(self) -> None:
        """Выбор папки для сохранения QR PNG."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для QR PNG",
            str(self._current_project_root or ""),
        )
        if not folder:
            return
        self.qr_output_edit.setText(folder)

    def _browse_copy_target_folder(self) -> None:
        """Выбрать папку назначения для копии проекта."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку назначения для копии",
            str(self._current_project_root.parent if self._current_project_root else ""),
        )
        if not folder:
            return
        self.copy_target_edit.setText(folder)

    def _browse_sheet_folder(self) -> None:
        """Выбрать папку с чертежами для автонумерации листов."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку с чертежами",
            str(self._current_project_root or ""),
        )
        if not folder:
            return
        self.stamp_sheet_folder_edit.setText(folder)

    def _browse_pdf_export_folder(self) -> None:
        """Выбрать папку для экспорта PDF."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для PDF",
            str(self._current_project_root or ""),
        )
        if not folder:
            return
        self.pdf_export_folder_edit.setText(folder)

    def _browse_pdf_source_folder(self) -> None:
        """Выбрать папку с чертежами для PDF-экспорта."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку с чертежами (.cdw) для PDF",
            str(self._current_project_root or ""),
        )
        if not folder:
            return
        self.pdf_source_folder_edit.setText(folder)

    def _browse_dwg_source_folder(self) -> None:
        """Выбрать папку с чертежами для DWG-экспорта."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку с чертежами (.cdw) для DWG",
            str(self._current_project_root or ""),
        )
        if not folder:
            return
        self.dwg_source_folder_edit.setText(folder)

    def _browse_dwg_export_folder(self) -> None:
        """Выбрать папку для экспорта DWG."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для DWG",
            str(self._current_project_root or ""),
        )
        if not folder:
            return
        self.dwg_export_folder_edit.setText(folder)

    def _copy_log_to_clipboard(self) -> None:
        """Скопировать весь текст лога в буфер обмена."""
        text = self.log_view.toPlainText()
        if not text.strip():
            QMessageBox.information(self, "Лог пуст", "Пока нечего копировать.")
            return
        QApplication.clipboard().setText(text)
        self.statusBar().showMessage("Лог скопирован в буфер обмена", 3000)

    def _save_log_to_file(self) -> None:
        """Сохранить текущий текст лога в файл."""
        text = self.log_view.toPlainText()
        if not text.strip():
            QMessageBox.information(self, "Лог пуст", "Пока нечего сохранять.")
            return
        default_name = f"nordfox_runtime_log_{QDateTime.currentDateTime().toString('yyyyMMdd_HHmmss')}.txt"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить лог",
            str(Path("logs") / default_name),
            "Text files (*.txt);;All files (*.*)",
        )
        if not file_path:
            return
        try:
            path = Path(file_path)
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
            self.statusBar().showMessage(f"Лог сохранен: {path}", 4000)
            logger.info("Лог сохранен в файл: %s", path)
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка сохранения", f"Не удалось сохранить лог:\n{exc}")

    def _open_log_file(self) -> None:
        """Открыть ранее сохраненный лог и показать в окне лога."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть лог",
            str(Path("logs")),
            "Log files (*.txt *.log *.json);;All files (*.*)",
        )
        if not file_path:
            return
        try:
            path = Path(file_path)
            content = path.read_text(encoding="utf-8")
            self.log_view.setPlainText(content)
            self.statusBar().showMessage(f"Лог открыт: {path}", 4000)
            logger.info("Открыт лог-файл: %s", path)
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка чтения", f"Не удалось прочитать лог:\n{exc}")

    def _collect_update_payload(self) -> tuple[Dict[str, float], Dict[str, str]] | None:
        """Собрать изменения по переменным сборки и комментариям чертежей."""
        if not self._assembly_info:
            QMessageBox.warning(self, "Нет сборки", "Проект ещё не просканирован.")
            return None

        changed_values: Dict[str, float] = {}
        for name, kv in self._var_index.items():
            if kv.document_type != "assembly":
                continue
            widget = self._var_inputs.get(name)
            if not widget:
                continue
            text = widget.text().strip()
            if text == "":
                continue
            try:
                new_val = float(text.replace(",", "."))
            except ValueError:
                QMessageBox.warning(self, "Неверное значение", f"Переменная {name}: введено не число.")
                return None
            old_val = float(kv.original_value) if kv.original_value is not None else None
            if old_val is None or abs(new_val - old_val) > 1e-6:
                changed_values[name] = new_val

        changed_comments: Dict[str, str] = {}
        for name, widget in self._drawing_comment_inputs.items():
            new_text = widget.text().strip()
            if new_text == "":
                continue
            original_comment: str | None = None
            for doc in self._documents:
                if doc.doc_type != "drawing":
                    continue
                kv = doc.variables.get(name)
                if kv is not None:
                    original_comment = kv.comment
                    break
            if original_comment is None or new_text != original_comment:
                changed_comments[name] = new_text

        return changed_values, changed_comments

    def _run_update_with_payload(
        self,
        changed_values: Dict[str, float],
        changed_comments: Dict[str, str],
    ) -> None:
        if self._updating_in_progress:
            return
        if not self._assembly_info:
            QMessageBox.warning(self, "Нет сборки", "Проект не просканирован.")
            return

        self._updating_in_progress = True
        progress = QProgressDialog(
            "Обновление переменных в проекте...",
            None,
            0,
            0,
            self,
        )
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setAutoClose(True)
        progress.setCancelButton(None)
        progress.show()
        QApplication.processEvents()

        try:
            result = update_project_variables(
                self._kompas,
                self._assembly_info,
                self._documents,
                changed_values,
                changed_comments or None,
            )
        finally:
            progress.close()
            self._updating_in_progress = False

        if self._json_log:
            self._json_log.add_action(
                type_="update_variables",
                status="success" if result.get("success") else "partial",
                input_={"changed_values": changed_values, "changed_comments": changed_comments},
                changes={
                    "documents_updated": result.get("documents_updated", 0),
                    "variables_updated": result.get("variables_updated", 0),
                },
                meta={"errors": result.get("errors", [])},
            )

        if result.get("success"):
            self.statusBar().showMessage(
                f"Переменные обновлены: документов={result.get('documents_updated')}, "
                f"переменных={result.get('variables_updated')}",
                5000,
            )
            QMessageBox.information(
                self,
                "Готово",
                f"Переменные обновлены.\n"
                f"Документов: {result.get('documents_updated')}\n"
                f"Переменных: {result.get('variables_updated')}",
            )
        else:
            errors = result.get("errors") or []
            msg = "\n".join(str(e) for e in errors) or "Неизвестная ошибка"
            QMessageBox.critical(
                self,
                "Ошибки при обновлении",
                f"Во время обновления возникли ошибки:\n{msg}",
            )

    def _on_copy_and_update_clicked(
        self,
        prepared_payload: tuple[Dict[str, float], Dict[str, str]] | bool | None = None,
    ) -> None:
        """Сценарий: создать проект-копию (и опционально применить изменения переменных)."""
        if not self._current_project_root:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите папку проекта.")
            return

        changed_values: Dict[str, float] = {}
        changed_comments: Dict[str, str] = {}
        source_root = self._current_project_root
        if isinstance(prepared_payload, tuple):
            changed_values, changed_comments = prepared_payload
        else:
            payload = self._collect_update_payload()
            if payload is not None:
                changed_values, changed_comments = payload
        # Для сценария копии сохраняем новые обозн./наимен. по ключу строки таблицы,
        # чтобы применить их уже к документам в новой папке.
        item_updates_by_key: Dict[Path, Dict[str, str]] = {}
        for doc_path in self._assembly_items.keys():
            mark_edit = self._assembly_item_new_marking.get(doc_path)
            name_edit = self._assembly_item_new_name.get(doc_path)
            new_mark = (mark_edit.text().strip() if mark_edit else "")
            new_name = (name_edit.text().strip() if name_edit else "")
            payload_meta: Dict[str, str] = {}
            if new_mark:
                payload_meta["marking"] = new_mark
            if new_name:
                payload_meta["name"] = new_name
            if payload_meta:
                item_updates_by_key[doc_path] = payload_meta

        target_text = self.copy_target_edit.text().strip()
        if not target_text:
            QMessageBox.warning(self, "Нет папки копии", "Укажите папку назначения для копии проекта.")
            return
        target_parent = Path(target_text)
        if not target_parent.exists():
            QMessageBox.warning(self, "Папка не найдена", f"Папка назначения не существует:\n{target_parent}")
            return

        progress = QProgressDialog(
            "Копирование проекта...",
            None,
            0,
            0,
            self,
        )
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setAutoClose(True)
        progress.setCancelButton(None)
        progress.show()
        QApplication.processEvents()

        try:
            folder_name = self.assembly_name_edit.text().strip()
            if not folder_name and self._assembly_info:
                folder_name = (self._assembly_info.name or "").strip()
            copy_result = copy_project_tree(
                self._current_project_root,
                target_parent,
                new_name=folder_name or None,
            )
        finally:
            progress.close()

        if not copy_result.get("success"):
            QMessageBox.critical(self, "Ошибка копирования", str(copy_result.get("error", "Unknown error")))
            return

        copied_path = Path(str(copy_result["target"]))
        self._current_project_root = copied_path
        self.project_path_edit.setText(str(copied_path))
        self.statusBar().showMessage(f"Работаем с копией проекта: {copied_path}", 5000)
        logger.info("Этап 0/3: проект скопирован в %s", copied_path)

        try:
            assembly_info, documents, var_index = scan_project(copied_path, self._kompas)
            self._assembly_info = assembly_info
            self._documents = documents
            self._var_index = var_index
            self._rebuild_variables_form()
            self._sync_assembly_and_stamp_fields()
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка сканирования", f"Не удалось просканировать копию:\n{exc}")
            return

        if changed_values or changed_comments:
            self._run_update_with_payload(changed_values, changed_comments)

        # После пересчета переменных заново читаем проект и строим обозн./наимен.
        self._on_rescan_clicked()
        auto_filled = self._fill_item_metadata_by_rules()
        auto_item_updates = self._collect_item_metadata_payload()

        # Ручные значения, которые были заполнены до копирования, должны иметь приоритет.
        item_updates_copy: Dict[Path, Dict[str, str]] = dict(auto_item_updates)
        for src_doc_path, payload_meta in item_updates_by_key.items():
            if source_root is None:
                continue
            try:
                rel = src_doc_path.relative_to(source_root)
            except ValueError:
                continue
            dst_doc_path = copied_path / rel
            existed = item_updates_copy.get(dst_doc_path, {})
            existed.update(payload_meta)
            item_updates_copy[dst_doc_path] = existed

        if item_updates_copy:
            meta_result = self._apply_item_metadata_updates(item_updates_copy, sync_assembly_components=False)
            if not meta_result.get("success"):
                errors = meta_result.get("errors") or []
                QMessageBox.warning(
                    self,
                    "Частичные ошибки",
                    "Копия создана и переменные обновлены, но часть обозначений/наименований не записалась:\n"
                    + "\n".join(str(e) for e in errors),
                )

        self._on_rescan_clicked()
        QMessageBox.information(
            self,
            "Проект создан",
            f"Копия проекта создана:\n{copied_path}\n\n"
            f"Переменных к обновлению: {len(changed_values)}\n"
            f"Комментариев чертежей: {len(changed_comments)}\n"
            f"Обозн./наимен. к применению: {len(item_updates_copy)}\n"
            f"Автоподстановка по правилам: обозн. {auto_filled['markings']}, наимен. {auto_filled['names']}",
        )

    def _on_update_variables_clicked(self) -> None:
        """
        Обработчик кнопки "Обновить переменные".

        """
        if not self._current_project_root:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите папку проекта.")
            return

        payload = self._collect_update_payload()
        if payload is None:
            return
        changed_values, changed_comments = payload
        item_updates = self._collect_item_metadata_payload()

        if not changed_values and not changed_comments and not item_updates:
            QMessageBox.information(
                self,
                "Нет изменений",
                "Не изменены ни переменные/комментарии, ни поля «Новое обозначение/наименование».",
            )
            return

        if changed_values or changed_comments:
            self._run_update_with_payload(changed_values, changed_comments)
            # Важно: читаем из Kompas уже пересчитанные формулами значения.
            self._on_rescan_clicked()
            self._fill_item_metadata_by_rules()
            # Добавляем автоматические значения, но оставляем ручные как приоритет.
            auto_updates = self._collect_item_metadata_payload()
            for p, payload in auto_updates.items():
                existed = item_updates.get(p, {})
                merged = dict(payload)
                merged.update(existed)
                item_updates[p] = merged

        if item_updates:
            meta_result = self._apply_item_metadata_updates(item_updates, sync_assembly_components=False)
            time.sleep(0.8)
            sync_result = self._run_isolated_assembly_sync(item_updates)
            if not meta_result.get("success"):
                errors = meta_result.get("errors") or []
                QMessageBox.warning(
                    self,
                    "Частичные ошибки",
                    "Переменные обновлены, но часть обозначений/наименований не записалась:\n"
                    + "\n".join(str(e) for e in errors),
                )
            if not sync_result.get("success"):
                errors = sync_result.get("errors") or []
                QMessageBox.warning(
                    self,
                    "Частичные ошибки синхронизации",
                    "Переменные обновлены, но синхронизация вхождений сборки завершилась с ошибками:\n"
                    + "\n".join(str(e) for e in errors),
                )

        # После записи в документы обновляем эталонные значения в UI.
        self._on_rescan_clicked()

    def _on_update_stamps_clicked(self) -> None:
        """Обновление штампов всех чертежей проекта."""
        if not self._current_project_root:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите папку проекта.")
            return

        designation = self.stamp_designation_edit.text().strip() or None
        name = self.stamp_name_edit.text().strip() or None

        developer = self.stamp_developer_edit.currentText().strip() or None
        checker = self.stamp_checker_edit.currentText().strip() or None
        organization = self.stamp_org_edit.text().strip() or None
        thickness_text = self.stamp_material_thickness_combo.currentText().strip()
        material_type = self.stamp_material_type_combo.currentText().strip()
        manual_material = self.stamp_material_edit.text().strip()

        material: str | None
        if manual_material:
            material = manual_material
        elif thickness_text and material_type:
            # Формат как в образцовом проекте kompas3d_project_manager.
            material = f"$d Лист {thickness_text},0х1250x2500 ГОСТ 19903-2015 ; {material_type} $"
            self.stamp_material_edit.setText(material)
        elif material_type:
            material = material_type
        else:
            material = None

        tech_control = self.stamp_tech_ctrl_edit.currentText().strip() or None
        norm_control = self.stamp_norm_ctrl_edit.currentText().strip() or None
        approved = self.stamp_approved_edit.currentText().strip() or None
        date = self.stamp_date_edit.text().strip() or None

        sheet_mode: str = "none"
        sheet_current: int | None = None
        sheet_total: int | None = None
        sc = self.stamp_sheet_current_spin.value()
        st = self.stamp_sheet_total_spin.value()
        if sc > 0 and st > 0:
            sheet_mode = "manual"
            sheet_current = sc
            sheet_total = st

        if date is None and any([developer, checker, tech_control, norm_control, approved]):
            date = QDateTime.currentDateTime().date().toString("dd.MM.yyyy")
            self.stamp_date_edit.setText(date)

        role_dates: Dict[int, str] = {}
        if date:
            if developer:
                role_dates[130] = date
            if checker:
                role_dates[131] = date
            if tech_control:
                role_dates[132] = date
            if norm_control:
                role_dates[134] = date
            if approved:
                role_dates[135] = date

        has_sheet = sheet_mode in ("batch", "manual")
        if not any(
            [
                developer,
                checker,
                organization,
                material,
                tech_control,
                norm_control,
                approved,
                date,
                designation,
                name,
                has_sheet,
            ]
        ):
            QMessageBox.information(self, "Нет данных", "Заполните хотя бы одно поле для обновления штампов.")
            return

        logger.info("Запуск обновления штампов чертежей...")

        progress = QProgressDialog(
            "Обновление штампов чертежей...",
            None,
            0,
            0,
            self,
        )
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setAutoClose(True)
        progress.setCancelButton(None)
        progress.show()
        QApplication.processEvents()

        result = update_all_drawing_stamps(
            self._kompas,
            self._current_project_root,
            developer=developer,
            checker=checker,
            organization=organization,
            material=material,
            tech_control=tech_control,
            norm_control=norm_control,
            approved=approved,
            date=date,
            role_dates=role_dates or None,
            designation=designation,
            name=name,
            sheet_mode=sheet_mode,
            sheet_current=sheet_current,
            sheet_total=sheet_total,
        )

        progress.close()

        if self._json_log:
            self._json_log.add_action(
                type_="update_stamps",
                status="success" if result.get("success") else "partial",
                input_={
                    "designation": designation,
                    "name": name,
                    "developer": developer,
                    "checker": checker,
                    "organization": organization,
                    "material": material,
                    "tech_control": tech_control,
                    "norm_control": norm_control,
                    "approved": approved,
                    "date": date,
                    "sheet_mode": sheet_mode,
                    "sheet_current": sheet_current,
                    "sheet_total": sheet_total,
                },
                changes={
                    "drawings_total": result.get("drawings_total", 0),
                    "drawings_updated": result.get("drawings_updated", 0),
                },
                meta={"errors": result.get("errors", [])},
            )

        if result.get("success"):
            self.statusBar().showMessage(
                f"Штампы обновлены: чертежей={result.get('drawings_updated', 0)} из {result.get('drawings_total', 0)}",
                5000,
            )
            QMessageBox.information(
                self,
                "Штампы обновлены",
                f"Чертежей обработано: {result.get('drawings_total', 0)}\n"
                f"Штампы обновлены у: {result.get('drawings_updated', 0)}",
            )
        else:
            errors = result.get("errors") or []
            msg = "\n".join(str(e) for e in errors) or "Неизвестная ошибка"
            QMessageBox.critical(
                self,
                "Ошибки при обновлении штампов",
                f"Во время обновления штампов возникли ошибки:\n{msg}",
            )

    def _on_export_drawings_pdf_clicked(self) -> None:
        """Экспорт всех CDW чертежей в PDF через локальный сервис."""
        if not self._current_project_root:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите папку проекта.")
            return

        source_folder_text = self.pdf_source_folder_edit.text().strip()
        source_root = Path(source_folder_text) if source_folder_text else self._current_project_root
        if not source_root.exists():
            QMessageBox.warning(self, "Папка не найдена", f"Папка с чертежами не существует:\n{source_root}")
            return

        output_folder_text = self.pdf_export_folder_edit.text().strip()
        output_folder = Path(output_folder_text) if output_folder_text else (source_root / "PDF")
        merge_into_one = self.pdf_merge_check.isChecked()
        merged_name = self.pdf_merged_name_edit.text().strip() or None

        logger.info("Запуск экспорта CDW в PDF через локальный сервис...")

        progress = QProgressDialog(
            "Экспорт CDW в PDF...",
            None,
            0,
            0,
            self,
        )
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setAutoClose(True)
        progress.setCancelButton(None)
        progress.show()
        QApplication.processEvents()

        exporter = DrawingPdfExporter()
        result = exporter.export_all_drawings_to_pdf(
            project_root=source_root,
            output_folder=output_folder,
            merge_into_one=merge_into_one,
            merged_output_name=merged_name,
        )
        progress.close()

        if self._json_log:
            self._json_log.add_action(
                type_="export_pdf",
                status="success" if result.get("success") else "partial",
                input_={
                    "source_folder": str(source_root),
                    "output_folder": str(output_folder),
                    "merge_into_one": merge_into_one,
                    "merged_name": merged_name,
                },
                changes={
                    "drawings_total": result.get("total_drawings", 0),
                    "exported_pdfs": result.get("exported_pdfs", 0),
                    "merged_pdf": result.get("merged_pdf"),
                },
                meta={"errors": result.get("errors", [])},
            )

        if result.get("success"):
            info_lines = [
                f"Чертежей обработано: {result.get('total_drawings', 0)}",
                f"Экспортировано PDF: {result.get('exported_pdfs', 0)}",
                f"Ошибок: {result.get('failed_drawings', 0)}",
            ]
            merged_pdf = result.get("merged_pdf")
            if merged_pdf:
                info_lines.append(f"Объединенный PDF:\n{merged_pdf}")

            self.statusBar().showMessage(
                f"Экспорт PDF завершен: {result.get('exported_pdfs', 0)} из {result.get('total_drawings', 0)}",
                5000,
            )
            QMessageBox.information(self, "Экспорт PDF завершен", "\n".join(info_lines))
        else:
            errors = result.get("errors") or []
            msg = "\n".join(str(e) for e in errors) or "Неизвестная ошибка"
            QMessageBox.critical(
                self,
                "Ошибки при экспорте PDF",
                f"Во время экспорта PDF возникли ошибки:\n{msg}",
            )

    def _on_export_drawings_dwg_clicked(self) -> None:
        """Прямой экспорт всех CDW чертежей в DWG через локальный сервис."""
        if not self._current_project_root:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите папку проекта.")
            return

        source_folder_text = self.dwg_source_folder_edit.text().strip()
        source_root = Path(source_folder_text) if source_folder_text else self._current_project_root
        if not source_root.exists():
            QMessageBox.warning(self, "Папка не найдена", f"Папка с чертежами не существует:\n{source_root}")
            return

        output_folder_text = self.dwg_export_folder_edit.text().strip()
        output_folder = Path(output_folder_text) if output_folder_text else (source_root / "DWG")

        logger.info("Запуск прямого экспорта CDW в DWG через локальный сервис...")

        progress = QProgressDialog(
            "Экспорт CDW в DWG...",
            None,
            0,
            0,
            self,
        )
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setAutoClose(True)
        progress.setCancelButton(None)
        progress.show()
        QApplication.processEvents()

        exporter = DrawingDwgExporter()
        result = exporter.export_all_drawings_to_dwg(
            project_root=source_root,
            output_folder=output_folder,
        )
        progress.close()

        if self._json_log:
            self._json_log.add_action(
                type_="export_dwg",
                status="success" if result.get("success") else "partial",
                input_={
                    "source_folder": str(source_root),
                    "output_folder": str(output_folder),
                },
                changes={
                    "drawings_total": result.get("total_drawings", 0),
                    "exported_dwgs": result.get("exported_dwgs", 0),
                },
                meta={"errors": result.get("errors", [])},
            )

        if result.get("success"):
            info_lines = [
                f"Чертежей обработано: {result.get('total_drawings', 0)}",
                f"Экспортировано DWG: {result.get('exported_dwgs', 0)}",
                f"Ошибок: {result.get('failed_drawings', 0)}",
                f"Папка вывода:\n{output_folder}",
            ]
            self.statusBar().showMessage(
                f"Экспорт DWG завершен: {result.get('exported_dwgs', 0)} из {result.get('total_drawings', 0)}",
                5000,
            )
            QMessageBox.information(self, "Экспорт DWG завершен", "\n".join(info_lines))
        else:
            errors = result.get("errors") or []
            msg = "\n".join(str(e) for e in errors) or "Неизвестная ошибка"
            QMessageBox.critical(
                self,
                "Ошибки при экспорте DWG",
                f"Во время экспорта DWG возникли ошибки:\n{msg}",
            )

    def _on_auto_number_sheets_clicked(self) -> None:
        """Отдельная операция автонумерации листов в выбранной папке."""
        target_folder = self.stamp_sheet_folder_edit.text().strip()
        root = Path(target_folder) if target_folder else self._current_project_root
        if not root:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите папку проекта или папку чертежей.")
            return
        if not root.exists():
            QMessageBox.warning(self, "Папка не найдена", f"Папка не существует:\n{root}")
            return

        progress = QProgressDialog(
            "Автонумерация листов...",
            None,
            0,
            0,
            self,
        )
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setAutoClose(True)
        progress.setCancelButton(None)
        progress.show()
        QApplication.processEvents()

        result = update_all_drawing_stamps(
            self._kompas,
            root,
            sheet_mode="batch",
        )
        progress.close()

        if result.get("success"):
            QMessageBox.information(
                self,
                "Готово",
                f"Листы пронумерованы.\n"
                f"Чертежей обработано: {result.get('drawings_total', 0)}\n"
                f"Обновлено: {result.get('drawings_updated', 0)}",
            )
        else:
            errors = result.get("errors") or []
            QMessageBox.critical(
                self,
                "Ошибки нумерации",
                "Во время автонумерации возникли ошибки:\n" + "\n".join(str(e) for e in errors),
            )

    def _on_generate_qr_clicked(self) -> None:
        """Сгенерировать QR PNG по данным из формы."""
        data = self.qr_data_edit.text().strip()
        if not data:
            QMessageBox.warning(self, "Пустые данные", "Введите строку данных для QR-кода.")
            return

        # Папка вывода
        output_folder = self.qr_output_edit.text().strip()
        if not output_folder:
            if self._current_project_root:
                output_folder = str(self._current_project_root / "QR")
            else:
                output_folder = str(Path.cwd() / "QR")

        try:
            scale = int(self.qr_scale_edit.text().strip() or "10")
            border = int(self.qr_border_edit.text().strip() or "4")
        except ValueError:
            QMessageBox.warning(self, "Неверные параметры", "Scale и Quiet zone должны быть целыми числами.")
            return

        # Имя файла по времени/шаблону
        from datetime import datetime

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"QR_{ts}.png"
        png_path = Path(output_folder) / filename

        ok = generate_qr_png(data=data, png_path=png_path, scale=scale, border=border)
        if ok:
            self.statusBar().showMessage(f"QR PNG создан: {png_path}", 5000)
            QMessageBox.information(self, "Успех", f"QR PNG создан:\n{png_path}")
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось создать QR PNG. См. лог для подробностей.")

    def _show_about(self) -> None:
        QMessageBox.about(
            self,
            "О программе",
            "NordFox Module Manager\n\n"
            "Приложение для работы с проектами NordFox в КОМПАС-3D:\n"
            "- обновление переменных сборки, деталей и чертежей;\n"
            "- обновление обозначений и наименований по правилам профилей;\n"
            "- генерация QR-кодов (PNG) для деталей.\n\n"
            "Эта версия содержит каркас интерфейса и базовые функции.\n"
            "Детальная логика обновления переменных будет добавлена далее.",
        )

