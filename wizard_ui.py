"""
competency_wizard/wizard_ui.py
職能說明書精靈 — PyQt6 UI
流程：初始化 → Step1(5W2H 輸入) → Step2(分析結果) → Step3(缺口詳情) → 輸出 Excel
"""

import sys
from pathlib import Path
from typing import Optional

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextEdit, QLineEdit, QPushButton, QProgressBar,
    QStackedWidget, QGroupBox, QFormLayout, QSplitter,
    QListWidget, QListWidgetItem, QFileDialog, QMessageBox,
    QScrollArea, QFrame, QComboBox, QCheckBox, QTabWidget,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QColor

from wizard_rag import WizardRAG
from gap_analyzer import GapAnalyzer, UserInput5W2H, GapReport
from excel_exporter import export_to_excel


# ─────────────────────────────────────────
# 全域樣式
# ─────────────────────────────────────────

APP_STYLE = """
/* ── 字型與文字顏色（不設全域背景，避免染色） ── */
QWidget {
    font-family: "Microsoft JhengHei", "PingFang TC", sans-serif;
    font-size: 10pt;
    color: #212529;
}

/* 僅對主視窗與 StackedWidget 頁面設底色 */
QMainWindow, QStackedWidget, QScrollArea > QWidget#page_bg {
    background-color: #EEF2F8;
}

/* ── 輸入元件 ── */
QLineEdit, QTextEdit {
    background: #ffffff;
    border: 1px solid #C8D3E0;
    border-radius: 5px;
    padding: 4px 8px;
    selection-background-color: #4472C4;
    selection-color: #ffffff;
}
QLineEdit:focus, QTextEdit:focus {
    border: 1.5px solid #4472C4;
    background: #FAFCFF;
}
QLineEdit:read-only, QTextEdit[readOnly="true"] {
    background: #F5F7FB;
    border-color: #D8DEE6;
    color: #3A4A5C;
}

/* ── 下拉選單 ── */
QComboBox {
    background: #ffffff;
    border: 1px solid #C8D3E0;
    border-radius: 5px;
    padding: 4px 8px;
    min-height: 26px;
    color: #212529;
}
QComboBox:focus { border: 1.5px solid #4472C4; }
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 22px;
    border-left: 1px solid #D0D7E2;
    border-top-right-radius: 5px;
    border-bottom-right-radius: 5px;
    background: #F0F4FA;
}
QComboBox QAbstractItemView {
    background: #ffffff;
    border: 1px solid #C8D3E0;
    selection-background-color: #D6E4F7;
    selection-color: #1a3a6e;
    outline: none;
    padding: 2px;
}

/* ── 按鈕（白底，清晰邊框） ── */
QPushButton {
    background: #ffffff;
    color: #2F5496;
    border: 1.5px solid #8AAAC8;
    border-radius: 5px;
    padding: 5px 18px;
    font-weight: bold;
    min-height: 28px;
}
QPushButton:hover  { background: #EBF2FB; border-color: #4472C4; color: #1a3a6e; }
QPushButton:pressed { background: #D4E4F5; border-color: #2F5496; }
QPushButton:disabled { background: #F0F3F7; color: #A0AABB; border-color: #C8D3DE; }

QPushButton#primary {
    background: #2F5496;
    color: #ffffff;
    border: none;
    min-height: 28px;
}
QPushButton#primary:hover   { background: #3A64B0; }
QPushButton#primary:pressed { background: #243F74; }
QPushButton#primary:disabled { background: #8DA4C4; color: #D8E4F0; }

QPushButton#success {
    background: #1A6E3C;
    color: #ffffff;
    border: none;
}
QPushButton#success:hover   { background: #208046; }
QPushButton#success:pressed { background: #14562F; }
QPushButton#success:disabled { background: #7AAD90; color: #D0E8DA; }

QPushButton#danger {
    background: #B83227;
    color: #ffffff;
    border: none;
}
QPushButton#danger:hover { background: #CC3B2E; }

/* ── GroupBox ── */
QGroupBox {
    background: #ffffff;
    border: 1px solid #D0DAE8;
    border-radius: 7px;
    margin-top: 14px;
    padding: 6px 10px 8px 10px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 6px;
    color: #2F5496;
    font-weight: bold;
    font-size: 10pt;
    background: #ffffff;
}

/* ── 捲動區 ── */
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical {
    background: #EEF1F5;
    width: 8px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #B0BECF;
    border-radius: 4px;
    min-height: 24px;
}
QScrollBar::handle:vertical:hover { background: #8A9CB0; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }

/* ── 分頁標籤 ── */
QTabWidget::pane {
    border: 1px solid #D0DAE8;
    border-radius: 0 6px 6px 6px;
    background: #ffffff;
}
QTabBar::tab {
    background: #E3EBF6;
    color: #4A5568;
    border: 1px solid #C8D3E0;
    border-bottom: none;
    border-top-left-radius: 5px;
    border-top-right-radius: 5px;
    padding: 5px 14px;
    margin-right: 2px;
    font-weight: bold;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #2F5496;
    border-color: #D0DAE8;
}
QTabBar::tab:hover:!selected { background: #D4E0F0; }

/* ── 進度條 ── */
QProgressBar {
    border: 1px solid #C8D3E0;
    border-radius: 5px;
    background: #E8EEF5;
    text-align: center;
    height: 16px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #4472C4, stop:1 #6A96D8);
    border-radius: 4px;
}

/* ── 核取方塊 ── */
QCheckBox { spacing: 8px; }
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 1.5px solid #7A8FA6;
    border-radius: 3px;
    background: white;
}
QCheckBox::indicator:checked {
    background: #2F5496;
    border-color: #2F5496;
    image: none;
}
QCheckBox::indicator:hover { border-color: #4472C4; }

/* ── 分隔線 ── */
QFrame[frameShape="4"] { color: #D0DAE8; }

/* ── Splitter ── */
QSplitter::handle {
    background: #D0DAE8;
    width: 3px;
}
QSplitter::handle:hover { background: #4472C4; }
"""


# ─────────────────────────────────────────
# 背景工作執行緒
# ─────────────────────────────────────────

class InitThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)   # success, error_msg

    def __init__(self, rag: WizardRAG, force_rebuild: bool = False):
        super().__init__()
        self.rag = rag
        self.force_rebuild = force_rebuild

    def run(self):
        try:
            if self.force_rebuild:
                self.rag.invalidate_cache()
            self.rag.initialize(progress_cb=lambda msg: self.progress.emit(msg))
            self.finished.emit(True, "")
        except Exception as e:
            self.finished.emit(False, str(e))


class AnalyzeThread(QThread):
    finished = pyqtSignal(object)   # GapReport or None
    error = pyqtSignal(str)

    def __init__(self, analyzer: GapAnalyzer, user_input: UserInput5W2H):
        super().__init__()
        self.analyzer = analyzer
        self.user_input = user_input

    def run(self):
        try:
            report = self.analyzer.analyze(self.user_input)
            self.finished.emit(report)
        except Exception as e:
            self.error.emit(str(e))


# ─────────────────────────────────────────
# 主視窗
# ─────────────────────────────────────────

class WizardMainWindow(QMainWindow):
    def __init__(self, engine=None):
        super().__init__()
        self.setWindowTitle("職能說明書精靈")
        self.setMinimumSize(960, 700)

        self.rag = WizardRAG(engine=engine)
        self.analyzer: Optional[GapAnalyzer] = None
        self.report: Optional[GapReport] = None
        self._init_thread: Optional[InitThread] = None
        self._analyze_thread: Optional[AnalyzeThread] = None

        self._build_ui()
        self._start_init()

    # ─── UI 建立 ─────────────────────────────

    def _build_ui(self):
        central = QWidget()
        central.setStyleSheet("background:#EEF2F8;")
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._top_bar = self._make_top_bar()
        layout.addWidget(self._top_bar)

        self.stack = QStackedWidget()
        layout.addWidget(self.stack, 1)

        self._page_loading = self._make_loading_page()
        self._page_form    = self._make_form_page()
        self._page_result  = self._make_result_page()

        self.stack.addWidget(self._page_loading)  # 0
        self.stack.addWidget(self._page_form)      # 1
        self.stack.addWidget(self._page_result)    # 2

    def _make_top_bar(self) -> QWidget:
        bar = QFrame()
        bar.setFixedHeight(52)
        bar.setStyleSheet(
            "QFrame { background: qlineargradient(x1:0,y1:0,x2:1,y2:0,"
            "stop:0 #1E3A6E, stop:1 #2F5496); }"
        )
        h = QHBoxLayout(bar)
        h.setContentsMargins(20, 0, 20, 0)

        dot = QLabel("●")
        dot.setStyleSheet("color:#7EC8E3; font-size:10pt; margin-right:4px;")
        h.addWidget(dot)

        title = QLabel("職能說明書精靈")
        title.setFont(QFont("Microsoft JhengHei", 14, QFont.Weight.Bold))
        title.setStyleSheet("color:white; letter-spacing:1px;")
        h.addWidget(title)
        h.addStretch()

        self._status_label = QLabel("初始化中...")
        self._status_label.setStyleSheet(
            "color:#A8C8F0; font-size:9pt; "
            "background:rgba(255,255,255,0.08); "
            "border-radius:4px; padding:2px 10px;"
        )
        h.addWidget(self._status_label)

        return bar

    def _make_loading_page(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background:#EEF2F8;")
        v = QVBoxLayout(w)
        v.setAlignment(Qt.AlignmentFlag.AlignCenter)
        v.setSpacing(18)

        icon = QLabel("⚙")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setStyleSheet("font-size:40pt; color:#4472C4;")
        v.addWidget(icon)

        self._loading_label = QLabel("正在載入 Embedding 模型，請稍候...")
        self._loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._loading_label.setFont(QFont("Microsoft JhengHei", 11))
        self._loading_label.setStyleSheet("color:#4A5568;")
        v.addWidget(self._loading_label)

        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 0)
        self._progress_bar.setFixedWidth(440)
        self._progress_bar.setFixedHeight(14)
        v.addWidget(self._progress_bar, alignment=Qt.AlignmentFlag.AlignCenter)

        btn_rebuild = QPushButton("強制重建索引")
        btn_rebuild.setFixedWidth(160)
        btn_rebuild.setObjectName("danger")
        btn_rebuild.clicked.connect(self._on_force_rebuild)
        v.addWidget(btn_rebuild, alignment=Qt.AlignmentFlag.AlignCenter)

        return w

    def _make_form_page(self) -> QWidget:
        """Step 1: 5W2H 輸入表單"""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { background:#EEF2F8; }")
        inner = QWidget()
        inner.setStyleSheet("background:#EEF2F8;")
        scroll.setWidget(inner)
        v = QVBoxLayout(inner)
        v.setContentsMargins(28, 20, 28, 20)
        v.setSpacing(14)

        def section(title):
            gb = QGroupBox(title)
            return gb

        # What
        gb_what = section("What — 做什麼")
        f = QFormLayout(gb_what)
        f.setSpacing(10)
        f.setContentsMargins(10, 8, 10, 10)
        self._what_tasks = QTextEdit()
        self._what_tasks.setPlaceholderText("描述主要工作任務（例：撰寫行銷企劃案、管理社群媒體帳號）")
        self._what_tasks.setFixedHeight(72)
        self._what_outputs = QLineEdit()
        self._what_outputs.setPlaceholderText("工作產出/交付物（例：企劃書、月報、產品說明頁）")
        f.addRow("工作任務：", self._what_tasks)
        f.addRow("工作產出：", self._what_outputs)
        v.addWidget(gb_what)

        # Why
        gb_why = section("Why — 為何做")
        f2 = QFormLayout(gb_why)
        f2.setSpacing(10)
        f2.setContentsMargins(10, 8, 10, 10)
        self._why_purpose = QLineEdit()
        self._why_purpose.setPlaceholderText("工作目的（例：提升品牌知名度、達成業績目標）")
        f2.addRow("工作目的：", self._why_purpose)
        v.addWidget(gb_why)

        # Who
        gb_who = section("Who — 誰做 / 與誰協作")
        f3 = QFormLayout(gb_who)
        f3.setSpacing(10)
        f3.setContentsMargins(10, 8, 10, 10)
        self._who_role = QLineEdit()
        self._who_role.setPlaceholderText("自身職稱（例：行銷專員、資深工程師）")
        self._who_collaborate = QLineEdit()
        self._who_collaborate.setPlaceholderText("主要協作對象（例：業務部、設計師、客戶）")
        f3.addRow("自身角色：", self._who_role)
        f3.addRow("協作對象：", self._who_collaborate)
        v.addWidget(gb_who)

        # When
        gb_when = section("When — 何時做")
        f4 = QFormLayout(gb_when)
        f4.setSpacing(10)
        f4.setContentsMargins(10, 8, 10, 10)
        self._when_frequency = QComboBox()
        self._when_frequency.addItems(["每日", "每週", "每月", "每季", "專案型（不固定）", "其他"])
        self._when_frequency.setEditable(True)
        f4.addRow("執行頻率：", self._when_frequency)
        v.addWidget(gb_when)

        # Where
        gb_where = section("Where — 在哪做")
        f5 = QFormLayout(gb_where)
        f5.setSpacing(10)
        f5.setContentsMargins(10, 8, 10, 10)
        self._where_env = QLineEdit()
        self._where_env.setPlaceholderText("工作環境/地點（例：辦公室、工廠現場、遠端居家）")
        f5.addRow("工作環境：", self._where_env)
        v.addWidget(gb_where)

        # How
        gb_how = section("How — 如何做")
        f6 = QFormLayout(gb_how)
        f6.setSpacing(10)
        f6.setContentsMargins(10, 8, 10, 10)
        self._how_skills = QTextEdit()
        self._how_skills.setPlaceholderText("使用的技能/工具/方法（例：Excel 資料分析、Python 自動化、溝通協商）")
        self._how_skills.setFixedHeight(72)
        f6.addRow("技能/工具：", self._how_skills)
        v.addWidget(gb_how)

        # How Much
        gb_howmuch = section("How Much — 做到什麼程度")
        f7 = QFormLayout(gb_howmuch)
        f7.setSpacing(10)
        f7.setContentsMargins(10, 8, 10, 10)
        self._how_much = QLineEdit()
        self._how_much.setPlaceholderText("績效指標（例：按時完成率 95%、客戶滿意度 4.5/5、錯誤率 <2%）")
        f7.addRow("績效指標：", self._how_much)
        v.addWidget(gb_howmuch)

        # 按鈕列
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_clear = QPushButton("清除")
        btn_clear.setFixedHeight(36)
        btn_clear.setFixedWidth(90)
        btn_clear.clicked.connect(self._clear_form)

        self._btn_analyze = QPushButton("開始分析 →")
        self._btn_analyze.setObjectName("primary")
        self._btn_analyze.setFixedHeight(38)
        self._btn_analyze.setFixedWidth(160)
        self._btn_analyze.setFont(QFont("Microsoft JhengHei", 11, QFont.Weight.Bold))
        self._btn_analyze.clicked.connect(self._on_analyze)

        btn_row.addWidget(btn_clear)
        btn_row.addStretch()
        btn_row.addWidget(self._btn_analyze)
        v.addLayout(btn_row)

        return scroll

    def _make_result_page(self) -> QWidget:
        """Step 2+3: 結果與缺口"""
        w = QWidget()
        w.setStyleSheet("background:#EEF2F8;")
        v = QVBoxLayout(w)
        v.setContentsMargins(14, 10, 14, 10)
        v.setSpacing(8)

        # ── 頂部狀態列 ─────────────────────────────
        status_bar = QFrame()
        status_bar.setStyleSheet(
            "QFrame { background:#ffffff; border:1px solid #D0DAE8; "
            "border-radius:6px; }"
        )
        status_bar.setFixedHeight(40)
        sh = QHBoxLayout(status_bar)
        sh.setContentsMargins(12, 0, 12, 0)

        self._result_status = QLabel("")
        self._result_status.setFont(QFont("Microsoft JhengHei", 10))
        self._result_status.setStyleSheet(
            "color:#2F5496; font-weight:bold; background:transparent; border:none;"
        )
        self._btn_export = QPushButton("匯出 Excel")
        self._btn_export.setObjectName("success")
        self._btn_export.setFixedHeight(28)
        self._btn_export.setEnabled(False)
        self._btn_export.clicked.connect(self._on_export)

        sh.addWidget(self._result_status, 1)
        sh.addWidget(self._btn_export)
        v.addWidget(status_bar)

        # ── 水平分割 ───────────────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        # ── 左側：職能基準選擇 + 可編輯 5W2H ──────────
        left = QFrame()
        left.setStyleSheet(
            "QFrame { background:#ffffff; border:1px solid #D0DAE8; "
            "border-radius:7px; }"
        )
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 10, 12, 10)
        lv.setSpacing(8)

        lbl_match = QLabel("相似職能基準")
        lbl_match.setStyleSheet(
            "font-weight:bold; color:#2F5496; font-size:9pt; "
            "background:transparent; border:none;"
        )
        lv.addWidget(lbl_match)

        self._match_combo = QComboBox()
        self._match_combo.setFont(QFont("Microsoft JhengHei", 9))
        self._match_combo.currentIndexChanged.connect(self._on_match_selected)
        lv.addWidget(self._match_combo)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        lv.addWidget(sep)

        lbl_edit = QLabel("工作內容（可直接修改後重新分析）")
        lbl_edit.setStyleSheet(
            "font-weight:bold; color:#2F5496; font-size:9pt; "
            "background:transparent; border:none;"
        )
        lv.addWidget(lbl_edit)

        edit_scroll = QScrollArea()
        edit_scroll.setWidgetResizable(True)
        edit_scroll.setFrameShape(QFrame.Shape.NoFrame)
        edit_scroll.setStyleSheet("background:transparent;")
        edit_inner = QWidget()
        edit_inner.setStyleSheet("background:transparent;")
        edit_scroll.setWidget(edit_inner)
        form = QFormLayout(edit_inner)
        form.setContentsMargins(0, 2, 4, 2)
        form.setSpacing(7)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        label_style = "color:#4A5568; font-size:9pt; background:transparent; border:none;"

        def _lbl(text):
            l = QLabel(text)
            l.setStyleSheet(label_style)
            return l

        self._r_who_role = QLineEdit()
        self._r_who_role.setPlaceholderText("自身職稱")
        form.addRow(_lbl("角色："), self._r_who_role)

        self._r_what_tasks = QTextEdit()
        self._r_what_tasks.setPlaceholderText("主要工作任務")
        self._r_what_tasks.setFixedHeight(54)
        form.addRow(_lbl("工作任務："), self._r_what_tasks)

        self._r_what_outputs = QLineEdit()
        self._r_what_outputs.setPlaceholderText("工作產出/交付物")
        form.addRow(_lbl("工作產出："), self._r_what_outputs)

        self._r_why_purpose = QLineEdit()
        self._r_why_purpose.setPlaceholderText("工作目的")
        form.addRow(_lbl("工作目的："), self._r_why_purpose)

        self._r_how_skills = QTextEdit()
        self._r_how_skills.setPlaceholderText("技能/工具/方法")
        self._r_how_skills.setFixedHeight(54)
        form.addRow(_lbl("技能/工具："), self._r_how_skills)

        self._r_how_much = QLineEdit()
        self._r_how_much.setPlaceholderText("績效指標")
        form.addRow(_lbl("績效指標："), self._r_how_much)

        self._r_when_frequency = QLineEdit()
        self._r_when_frequency.setPlaceholderText("執行頻率")
        form.addRow(_lbl("執行頻率："), self._r_when_frequency)

        self._r_where_env = QLineEdit()
        self._r_where_env.setPlaceholderText("工作環境")
        form.addRow(_lbl("工作環境："), self._r_where_env)

        self._r_who_collaborate = QLineEdit()
        self._r_who_collaborate.setPlaceholderText("協作對象")
        form.addRow(_lbl("協作對象："), self._r_who_collaborate)

        lv.addWidget(edit_scroll, 1)

        self._btn_reanalyze = QPushButton("重新分析 →")
        self._btn_reanalyze.setObjectName("primary")
        self._btn_reanalyze.setFixedHeight(34)
        self._btn_reanalyze.clicked.connect(self._on_reanalyze)
        lv.addWidget(self._btn_reanalyze)

        splitter.addWidget(left)

        # ── 右側：分頁職能基準資料 ─────────────────────
        self._detail_tabs = QTabWidget()

        def _make_tab() -> QTextEdit:
            t = QTextEdit()
            t.setReadOnly(True)
            t.setFont(QFont("Microsoft JhengHei", 10))
            t.setStyleSheet(
                "QTextEdit { background:#ffffff; border:none; padding:6px; }"
            )
            return t

        self._tab_basic = _make_tab()
        self._tab_gap   = _make_tab()

        _task_tab_w = QWidget()
        _task_tab_w.setStyleSheet("background:#ffffff;")
        _task_tab_v = QVBoxLayout(_task_tab_w)
        _task_tab_v.setContentsMargins(8, 8, 8, 8)
        _task_tab_v.setSpacing(6)

        self._task_combo = QComboBox()
        self._task_combo.setFont(QFont("Microsoft JhengHei", 9))
        self._task_combo.currentIndexChanged.connect(self._on_task_selected)
        _task_tab_v.addWidget(self._task_combo)

        self._tab_task_detail = _make_tab()
        _task_tab_v.addWidget(self._tab_task_detail, 1)

        self._detail_tabs.addTab(self._tab_basic, "基本資訊")
        self._detail_tabs.addTab(_task_tab_w,     "工作職能")
        self._detail_tabs.addTab(self._tab_gap,   "缺口分析")

        splitter.addWidget(self._detail_tabs)
        splitter.setSizes([330, 560])
        v.addWidget(splitter, 1)

        # ── 底部確認列 ────────────────────────────────
        confirm_bar = QFrame()
        confirm_bar.setStyleSheet(
            "QFrame { background:#ffffff; border:1px solid #D0DAE8; "
            "border-radius:6px; }"
        )
        confirm_bar.setFixedHeight(42)
        ch = QHBoxLayout(confirm_bar)
        ch.setContentsMargins(14, 0, 14, 0)

        self._confirm_check = QCheckBox(
            "我已確認以上缺口分析結果正確無誤，同意匯出職能說明書"
        )
        self._confirm_check.setFont(QFont("Microsoft JhengHei", 10))
        self._confirm_check.setStyleSheet(
            "QCheckBox { color:#2F5496; font-weight:bold; "
            "background:transparent; border:none; }"
        )
        self._confirm_check.toggled.connect(self._btn_export.setEnabled)
        ch.addWidget(self._confirm_check)
        ch.addStretch()
        v.addWidget(confirm_bar)

        return w

    # ─── 初始化流程 ───────────────────────────

    def _start_init(self, force_rebuild: bool = False):
        self.stack.setCurrentIndex(0)
        self._init_thread = InitThread(self.rag, force_rebuild)
        self._init_thread.progress.connect(self._loading_label.setText)
        self._init_thread.finished.connect(self._on_init_finished)
        self._init_thread.start()

    def _on_init_finished(self, success: bool, error_msg: str):
        if success:
            self.analyzer = GapAnalyzer(self.rag)
            mode = "共用索引" if self.rag.using_shared_engine else "獨立索引"
            self._status_label.setText(f"已就緒（{self.rag.chunk_count} chunks，{mode}）")
            self.stack.setCurrentIndex(1)
        else:
            self._loading_label.setText(f"初始化失敗：{error_msg}")
            self._status_label.setText("初始化失敗")

    def _on_force_rebuild(self):
        self._loading_label.setText("強制重建索引中...")
        self._start_init(force_rebuild=True)

    # ─── 表單操作 ─────────────────────────────

    def _clear_form(self):
        self._what_tasks.clear()
        self._what_outputs.clear()
        self._why_purpose.clear()
        self._who_role.clear()
        self._who_collaborate.clear()
        self._when_frequency.setCurrentIndex(0)
        self._where_env.clear()
        self._how_skills.clear()
        self._how_much.clear()

    def _collect_input(self) -> UserInput5W2H:
        return UserInput5W2H(
            what_tasks=self._what_tasks.toPlainText().strip(),
            what_outputs=self._what_outputs.text().strip(),
            why_purpose=self._why_purpose.text().strip(),
            who_role=self._who_role.text().strip(),
            who_collaborate=self._who_collaborate.text().strip(),
            when_frequency=self._when_frequency.currentText().strip(),
            where_environment=self._where_env.text().strip(),
            how_skills=self._how_skills.toPlainText().strip(),
            how_much_kpi=self._how_much.text().strip(),
        )

    def _on_analyze(self):
        if not self.analyzer:
            QMessageBox.warning(self, "未就緒", "RAG 尚未初始化，請稍候")
            return
        ui = self._collect_input()
        if not ui.to_search_query().strip():
            QMessageBox.warning(self, "輸入不足", "請至少填寫工作任務或技能欄位")
            return
        self._status_label.setText("分析中...")
        self._btn_analyze.setEnabled(False)
        self._analyze_thread = AnalyzeThread(self.analyzer, ui)
        self._analyze_thread.finished.connect(self._on_analyze_done)
        self._analyze_thread.error.connect(self._on_analyze_error)
        self._analyze_thread.start()

    def _on_analyze_done(self, report: GapReport):
        self.report = report
        self._btn_analyze.setEnabled(True)
        self._btn_reanalyze.setEnabled(True)
        self._status_label.setText("分析完成")
        self._populate_results(report)
        self.stack.setCurrentIndex(2)

    def _on_analyze_error(self, msg: str):
        self._btn_analyze.setEnabled(True)
        self._btn_reanalyze.setEnabled(True)
        self._status_label.setText("分析錯誤")
        QMessageBox.critical(self, "分析失敗", msg)

    # ─── 結果顯示 ─────────────────────────────

    def _populate_results(self, report: GapReport):
        self._confirm_check.setChecked(False)
        self._btn_export.setEnabled(False)

        ui = report.user_input
        self._r_who_role.setText(ui.who_role)
        self._r_what_tasks.setPlainText(ui.what_tasks)
        self._r_what_outputs.setText(ui.what_outputs)
        self._r_why_purpose.setText(ui.why_purpose)
        self._r_how_skills.setPlainText(ui.how_skills)
        self._r_how_much.setText(ui.how_much_kpi)
        self._r_when_frequency.setText(ui.when_frequency)
        self._r_where_env.setText(ui.where_environment)
        self._r_who_collaborate.setText(ui.who_collaborate)

        self._match_combo.blockSignals(True)
        self._match_combo.clear()
        for r in report.matched_standards:
            self._match_combo.addItem(
                f"[{r['score']:.2f}] {r['standard_name']}（{r['standard_code']}）"
            )
        self._match_combo.blockSignals(False)

        score = report.completeness_score
        self._result_status.setText(
            f"最佳匹配：{report.best_standard_name}  ｜  完整度：{score}%"
        )

        if self.analyzer:
            self._tab_gap.setPlainText(self.analyzer.get_summary_text(report))

        if self._match_combo.count() > 0:
            self._match_combo.setCurrentIndex(0)
            self._on_match_selected(0)

    def _on_match_selected(self, index: int):
        if self.report is None or index < 0 or index >= len(self.report.matched_standards):
            return
        std_code = self.report.matched_standards[index]["standard_code"]
        std_data = self.rag.get_standard(std_code)
        if not std_data:
            return

        self._current_std_data = std_data

        knowledge_list = std_data.get("competency_knowledge") or std_data.get("knowledge") or []
        skills_list    = std_data.get("competency_skills")    or std_data.get("skills")    or []
        self._k_lookup = {k.get("code", ""): k for k in knowledge_list if isinstance(k, dict)}
        self._s_lookup = {s.get("code", ""): s for s in skills_list    if isinstance(s, dict)}

        meta  = std_data.get("metadata")   or {}
        basic = std_data.get("basic_info") or {}
        b = ["═══ metadata ═══"]
        for k, v in meta.items():
            b.append(f"  {k}：{v}")
        b.append("")
        b.append("═══ basic_info ═══")
        for k, v in basic.items():
            b.append(f"  {k}：{v}")
        self._tab_basic.setPlainText("\n".join(b))

        tasks = std_data.get("competency_tasks") or []
        self._task_combo.blockSignals(True)
        self._task_combo.clear()
        for task in tasks:
            tid  = task.get("task_id", "")
            name = task.get("task_name", "")
            self._task_combo.addItem(f"[{tid}] {name}")
        self._task_combo.blockSignals(False)
        if self._task_combo.count() > 0:
            self._task_combo.setCurrentIndex(0)
            self._on_task_selected(0)

    def _on_task_selected(self, index: int):
        std_data = getattr(self, "_current_std_data", None)
        if not std_data:
            return
        tasks = std_data.get("competency_tasks") or []
        if index < 0 or index >= len(tasks):
            return

        task     = tasks[index]
        k_lookup = getattr(self, "_k_lookup", {})
        s_lookup = getattr(self, "_s_lookup", {})

        lines = []
        lines.append(f"▌ [{task.get('task_id','')}] {task.get('task_name','')}")
        if task.get("main_responsibility"):
            lines.append(f"  主責：{task['main_responsibility']}")
        lines.append(f"  層級：{task.get('level', '')}")

        output = task.get("output") or ""
        if isinstance(output, str) and output:
            lines.append(f"\n【工作產出】\n  {output}")
        elif isinstance(output, list):
            outs = [o for o in output if isinstance(o, dict)]
            if outs:
                lines.append("\n【工作產出】")
                for o in outs:
                    lines.append(f"  [{o.get('code','')}] {o.get('name','')}")

        behaviors = task.get("behaviors") or []
        if behaviors:
            lines.append(f"\n【行為指標】（{len(behaviors)} 項）")
            for bv in behaviors:
                if isinstance(bv, dict):
                    lines.append(f"  [{bv.get('code','')}] {bv.get('description','')}")
                elif isinstance(bv, str):
                    lines.append(f"  • {bv}")

        k_codes = task.get("knowledge") or []
        if k_codes:
            lines.append(f"\n【對應知識項目】（{len(k_codes)} 項）")
            for code in k_codes:
                info = k_lookup.get(code)
                if info:
                    lines.append(f"  [{code}] {info.get('name','')}")
                    if info.get("description"):
                        lines.append(f"      {info['description']}")
                else:
                    lines.append(f"  [{code}]")

        s_codes = task.get("skills") or []
        if s_codes:
            lines.append(f"\n【對應技能項目】（{len(s_codes)} 項）")
            for code in s_codes:
                info = s_lookup.get(code)
                if info:
                    lines.append(f"  [{code}] {info.get('name','')}")
                    if info.get("description"):
                        lines.append(f"      {info['description']}")
                else:
                    lines.append(f"  [{code}]")

        self._tab_task_detail.setPlainText("\n".join(lines))

    def _collect_result_input(self) -> UserInput5W2H:
        return UserInput5W2H(
            who_role=self._r_who_role.text().strip(),
            what_tasks=self._r_what_tasks.toPlainText().strip(),
            what_outputs=self._r_what_outputs.text().strip(),
            why_purpose=self._r_why_purpose.text().strip(),
            how_skills=self._r_how_skills.toPlainText().strip(),
            how_much_kpi=self._r_how_much.text().strip(),
            when_frequency=self._r_when_frequency.text().strip(),
            where_environment=self._r_where_env.text().strip(),
            who_collaborate=self._r_who_collaborate.text().strip(),
        )

    def _on_reanalyze(self):
        if not self.analyzer:
            QMessageBox.warning(self, "未就緒", "RAG 尚未初始化，請稍候")
            return
        ui = self._collect_result_input()
        if not ui.to_search_query().strip():
            QMessageBox.warning(self, "輸入不足", "請至少填寫工作任務或技能欄位")
            return
        self._status_label.setText("重新分析中...")
        self._btn_reanalyze.setEnabled(False)
        self._analyze_thread = AnalyzeThread(self.analyzer, ui)
        self._analyze_thread.finished.connect(self._on_analyze_done)
        self._analyze_thread.error.connect(self._on_analyze_error)
        self._analyze_thread.start()

    # ─── Excel 輸出 ───────────────────────────

    def _on_export(self):
        if not self.report:
            QMessageBox.warning(self, "無資料", "請先執行分析")
            return
        role_name = self.report.user_input.who_role or "職能說明書"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "儲存 Excel",
            str(Path.home() / f"{role_name}_職能說明書.xlsx"),
            "Excel 檔案 (*.xlsx)",
        )
        if not path:
            return
        try:
            out = export_to_excel(self.report, Path(path), role_name=role_name)
            QMessageBox.information(self, "完成", f"已儲存至：\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "匯出失敗", str(e))
