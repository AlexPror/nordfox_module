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
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QCheckBox,
)
from PyQt6.QtWidgets import QApplication, QProgressDialog

from ..core.qr_generator import generate_qr_png
from ..core.kompas_connector import KompasConnector
from ..core.variables_scanner import scan_project
from ..core.variables_updater import update_project_variables
from ..core.log_store import JsonLogStore
from ..core.models import KompasDocumentInfo, KompasVariable
from ..core.stamp_updater import update_all_drawing_stamps
from ..core.project_copy import copy_project_tree


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
        # Профиль и авто-обозначение/наименование по блокам: block_id -> {...}
        self._block_profile_combos: Dict[str, QComboBox] = {}
        self._block_designation_edits: Dict[str, QLineEdit] = {}
        self._block_name_edits: Dict[str, QLineEdit] = {}
        # Таблица уникальных деталей/сборок: key -> KompasDocumentInfo
        self._assembly_items: Dict[tuple[str, str, str], KompasDocumentInfo] = {}
        self._assembly_item_new_marking: Dict[tuple[str, str, str], QLineEdit] = {}
        self._assembly_item_new_name: Dict[tuple[str, str, str], QLineEdit] = {}

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
        self.copy_mode_check.setChecked(True)
        self.copy_target_edit = QLineEdit()
        self.copy_target_edit.setPlaceholderText("Папка назначения для копии проекта...")
        self.btn_copy_target_browse = QPushButton("Папка копии...")
        self.btn_copy_target_browse.clicked.connect(self._browse_copy_target_folder)
        self.btn_copy_and_update = QPushButton("Копировать и обновить")
        self.btn_copy_and_update.setEnabled(False)
        self.btn_copy_and_update.clicked.connect(self._on_copy_and_update_clicked)
        copy_row.addWidget(self.copy_mode_check)
        copy_row.addWidget(self.copy_target_edit)
        copy_row.addWidget(self.btn_copy_target_browse)
        copy_row.addWidget(self.btn_copy_and_update)
        top_layout.addLayout(copy_row)

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

        # Строка выбора типа профиля
        profile_row = QHBoxLayout()
        lbl_profile = QLabel("Тип профиля:")
        self.profile_combo = QComboBox()
        self.profile_combo.setEditable(False)
        self.profile_combo.setMinimumWidth(260)
        self._fill_profiles()

        profile_row.addWidget(lbl_profile)
        profile_row.addWidget(self.profile_combo, stretch=1)

        top_layout.addLayout(profile_row)

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

        self.stamp_developer_edit = QLineEdit()
        self.stamp_checker_edit = QLineEdit()
        self.stamp_org_edit = QLineEdit()
        # Материал теперь задается через выпадающий список толщины и типа материала.
        self.stamp_material_thickness_combo = QComboBox()
        self.stamp_material_type_combo = QComboBox()
        self.stamp_material_edit = QLineEdit()
        self.stamp_material_edit.setReadOnly(False)
        self.stamp_tech_ctrl_edit = QLineEdit()
        self.stamp_norm_ctrl_edit = QLineEdit()
        self.stamp_approved_edit = QLineEdit()
        self.stamp_date_edit = QLineEdit()
        self.stamp_order_edit = QLineEdit()

        # Предзаполненные значения ФИО (можно изменить вручную)
        self.stamp_developer_edit.setText("Воробьев")
        self.stamp_checker_edit.setText("Заметалин")
        self.stamp_approved_edit.setText("Сизонов")

        self.stamp_developer_edit.setPlaceholderText("Разраб. (ячейка 110)")
        self.stamp_checker_edit.setPlaceholderText("Пров. (ячейка 111)")
        self.stamp_org_edit.setPlaceholderText("Организация (ячейка 9)")
        self.stamp_material_thickness_combo.addItems(["", "6", "5", "4", "3", "2", "1"])
        self.stamp_material_type_combo.addItems([
            "",
            "09Г2С ГОСТ 19281-2014",
            "AISI 304",
            "Ст3. ГОСТ 19281-2014",
        ])
        self.stamp_material_edit.setPlaceholderText(
            "Материал (ячейка 3, для несборочных). Формируется из толщины и типа."
        )
        self.stamp_tech_ctrl_edit.setPlaceholderText("Т. контр. (ячейка 112)")
        self.stamp_norm_ctrl_edit.setPlaceholderText("Н. контр. (ячейка 114)")
        self.stamp_approved_edit.setPlaceholderText("Утв. (ячейка 115)")
        self.stamp_date_edit.setPlaceholderText("Дата разработки (ячейка 130, формат 01.01.2026)")
        self.stamp_order_edit.setPlaceholderText("Номер заказа (опционально, пока без логики)")

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
        stamp_form.addRow("Заказ:", self.stamp_order_edit)

        stamp_layout.addLayout(stamp_form)

        stamp_buttons_row = QHBoxLayout()
        self.btn_update_stamps = QPushButton("Обновить штампы чертежей")
        self.btn_update_stamps.setEnabled(True)
        self.btn_update_stamps.clicked.connect(self._on_update_stamps_clicked)
        stamp_buttons_row.addStretch(1)
        stamp_buttons_row.addWidget(self.btn_update_stamps)

        stamp_layout.addLayout(stamp_buttons_row)
        stamp_layout.addStretch(1)

        left_tabs.addTab(self.stamp_tab, "Штампы")

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

    def _fill_profiles(self) -> None:
        """
        Заполнить список профилей теми типами, которые вы перечислили.
        При расширении логики обозначений достаточно будет обновить этот список.
        """
        profiles = [
            # H-серия
            "Профиль H20.1",
            "Профиль H20",
            "Профиль H21",
            "Профиль H22",
            "Профиль H23",
            "Профиль H24",
            # DT
            "Профиль DT20",
            "Профиль DT21",
            "Профиль DT22",
            "Профиль DT23",
            "Профиль DT24",
            # H Hat
            "Профиль H Hat20",
            "Профиль H Hat21",
            "Профиль H Hat22",
            "Профиль H Hat23",
            "Профиль H Hat24",
            # T-серия
            "Профиль T12",
            "Профиль T15",
            "Профиль T16",
            "Профиль T25",
            "Профиль T26",
            "Профиль T20",
            "Профиль T21",
            "Профиль T Hat20",
            "Профиль T11N",
            # L-серия
            "Профиль L11N",
            "Профиль L20",
            "Профиль L15",
            "Профиль L16",
        ]
        self.profile_combo.clear()
        self.profile_combo.addItems(profiles)

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
        # Очищаем предыдущие блоки
        while self.vars_blocks_layout.count():
            item = self.vars_blocks_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._var_inputs.clear()
        self._block_profile_combos.clear()
        self._block_designation_edits.clear()
        self._block_name_edits.clear()
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
            if kv.block_id:
                blocks.setdefault(kv.block_id, {})[name] = kv
            else:
                ungrouped[name] = kv

        # Создаем блоки
        for block_id, vars_in_block in sorted(blocks.items(), key=lambda x: x[0].lower()):
            logger.info("[UI] Блок переменных: %s", block_id)
            group = QGroupBox(block_id)
            group.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            group.setStyleSheet("QGroupBox { font-weight: bold; }")

            vbox = QVBoxLayout(group)

            # Верхняя строка: выбор профиля + автообозначение/наименование
            header_row = QHBoxLayout()
            lbl_prof = QLabel("Профиль:")
            combo = QComboBox()
            combo.addItems(self.profile_combo.itemText(i) for i in range(self.profile_combo.count()))

            lbl_des = QLabel("Обозначение:")
            des_edit = QLineEdit()
            lbl_name = QLabel("Наименование:")
            name_edit = QLineEdit()

            header_row.addWidget(lbl_prof)
            header_row.addWidget(combo)
            header_row.addWidget(lbl_des)
            header_row.addWidget(des_edit)
            header_row.addWidget(lbl_name)
            header_row.addWidget(name_edit)

            vbox.addLayout(header_row)

            self._block_profile_combos[block_id] = combo
            self._block_designation_edits[block_id] = des_edit
            self._block_name_edits[block_id] = name_edit

            # Автозаполнение обозначения/наименования из сборки/деталей
            default_des, default_name = self._get_default_designation_name_for_block(block_id)
            if default_des:
                des_edit.setText(default_des)
            if default_name:
                name_edit.setText(default_name)

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

            self.vars_blocks_layout.addWidget(group)

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
            self.vars_blocks_layout.addWidget(group)

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
            self.vars_blocks_layout.addWidget(group)

        self.vars_blocks_layout.addStretch(1)

        # После перестройки формы переменных обновляем и таблицу деталей/подсборок
        self._rebuild_items_table()

    def _rebuild_items_table(self) -> None:
        """
        Построить таблицу уникальных деталей/подсборок:
        - используем только документы текущего уровня (assembly_info + parts из self._documents);
        - объединяем по ключу (тип, обозначение, наименование), чтобы метизы не дублировались.
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
            doc_type = doc.doc_type
            mark = doc.designation or ""
            name = doc.name or ""
            key = (doc_type, mark, name)
            # Один элемент на ключ (тип + обозначение + наименование)
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

        for row, (key, doc) in enumerate(self._assembly_items.items()):
            doc_type, mark, name = key

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
            self._assembly_item_new_marking[key] = new_mark_edit

            # Новое наименование
            new_name_edit = QLineEdit()
            self.items_table.setCellWidget(row, 4, new_name_edit)
            self._assembly_item_new_name[key] = new_name_edit

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

    def _get_default_designation_name_for_block(self, block_id: str) -> tuple[str | None, str | None]:
        """
        Попробовать подобрать разумное обозначение/наименование для блока,
        исходя из данных сборки и деталей.

        Логика первого приближения:
        - Obolochka: деталь с "оболочка" в имени;
        - Rigel: деталь с "ригель" в имени;
        - Stoiki: деталь с "стойка" в имени;
        - Kronshtein_MacFox: деталь с "macfox" в имени;
        - иначе — обозначение/имя сборки.
        """
        block_lower = block_id.lower()

        # 1. Пытаемся найти подходящую деталь
        if self._documents:
            for doc in self._documents:
                if doc.doc_type != "part":
                    continue
                name = (doc.name or "").lower()
                path_name = doc.path.name.lower()

                if "obolochka" in block_lower or "оболоч" in block_lower:
                    if "оболочка" in name or "оболочка" in path_name:
                        return doc.designation, doc.name

                if "rigel" in block_lower or "ригель" in block_lower:
                    if "ригель" in name or "ригель" in path_name:
                        return doc.designation, doc.name

                if "stoik" in block_lower or "стойк" in block_lower:
                    if "стойка" in name or "стойка" in path_name:
                        return doc.designation, doc.name

                if "kronshtein_macfox" in block_lower or "macfox" in block_lower:
                    if "macfox" in name or "macfox" in path_name:
                        return doc.designation, doc.name

        # 2. Фолбэк — обозначение и наименование сборки (если есть)
        if self._assembly_info:
            return self._assembly_info.designation, self._assembly_info.name

        return None, None

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
        prepared_payload: tuple[Dict[str, float], Dict[str, str]] | None = None,
    ) -> None:
        """Сценарий: копирование папки проекта и каскадное обновление в копии."""
        if not self._current_project_root:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите папку проекта.")
            return

        payload = prepared_payload or self._collect_update_payload()
        if payload is None:
            return
        changed_values, changed_comments = payload
        if not changed_values and not changed_comments:
            QMessageBox.information(self, "Нет изменений", "Вы не изменили ни одной переменной.")
            return

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
            copy_result = copy_project_tree(self._current_project_root, target_parent)
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
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка сканирования", f"Не удалось просканировать копию:\n{exc}")
            return

        self._run_update_with_payload(changed_values, changed_comments)

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
        if not changed_values and not changed_comments:
            QMessageBox.information(self, "Нет изменений", "Вы не изменили ни одной переменной.")
            return

        if self.copy_mode_check.isChecked():
            self._on_copy_and_update_clicked((changed_values, changed_comments))
            return

        self._run_update_with_payload(changed_values, changed_comments)

    def _on_update_stamps_clicked(self) -> None:
        """Обновление штампов всех чертежей проекта."""
        if not self._current_project_root:
            QMessageBox.warning(self, "Нет проекта", "Сначала выберите папку проекта.")
            return

        # Собираем значения полей; пустые строки означают "не изменять"
        developer = self.stamp_developer_edit.text().strip() or None
        checker = self.stamp_checker_edit.text().strip() or None
        organization = self.stamp_org_edit.text().strip() or None
        # Формируем материал из толщины и типа, если заданы
        thickness_text = self.stamp_material_thickness_combo.currentText().strip()
        material_type = self.stamp_material_type_combo.currentText().strip()
        manual_material = self.stamp_material_edit.text().strip()

        material: str | None
        if manual_material:
            # Если пользователь явно задал строку материала, используем её как есть
            material = manual_material
        elif thickness_text and material_type:
            # Автоматическая строка: "Лист {t} ГОСТ 19903-2015/ {material_type}"
            material = f"Лист {thickness_text} ГОСТ 19903-2015/ {material_type}"
            self.stamp_material_edit.setText(material)
        else:
            material = None

        tech_control = self.stamp_tech_ctrl_edit.text().strip() or None
        norm_control = self.stamp_norm_ctrl_edit.text().strip() or None
        approved = self.stamp_approved_edit.text().strip() or None
        date = self.stamp_date_edit.text().strip() or None
        order_number = self.stamp_order_edit.text().strip() or None

        # Автоматическая дата: если есть хотя бы одна фамилия, а дата пустая — ставим текущую
        if date is None and any([developer, checker, tech_control, norm_control, approved]):
            date = QDateTime.currentDateTime().date().toString("dd.MM.yyyy")
            self.stamp_date_edit.setText(date)

        # Даты напротив каждой фамилии (ячейки могут зависеть от вашего шаблона штампа).
        # Используем дефолтное соответствие 130–135, которое легко поменять позже.
        role_dates: Dict[int, str] = {}
        if date:
            # 130: Разраб., 131: Пров., 132: Т. контр., 134: Н. контр., 135: Утв.
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

        if not any(
            [developer, checker, organization, material, tech_control, norm_control, approved, date, order_number]
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
            order_number=order_number,
        )

        progress.close()

        if self._json_log:
            self._json_log.add_action(
                type_="update_stamps",
                status="success" if result.get("success") else "partial",
                input_={
                    "developer": developer,
                    "checker": checker,
                    "organization": organization,
                    "material": material,
                    "tech_control": tech_control,
                    "norm_control": norm_control,
                    "approved": approved,
                    "date": date,
                    "order_number": order_number,
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

