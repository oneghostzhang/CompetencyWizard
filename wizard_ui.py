"""
competency_wizard/wizard_ui.py
職能說明書精靈 — PyQt6 UI
流程：初始化 → Step1(5W2H 輸入) → Step2(分析結果) → Step3(缺口詳情) → 輸出 Excel
"""

import shutil
import sys
from pathlib import Path
from typing import Optional

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextEdit, QLineEdit, QPushButton, QProgressBar,
    QStackedWidget, QGroupBox, QFormLayout, QSplitter,
    QListWidget, QListWidgetItem, QFileDialog, QMessageBox,
    QScrollArea, QFrame, QComboBox, QCheckBox, QTabWidget,
    QDialog,
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
/* ══ 配色借鑒自 Graph_RAG_test ══
   主色  #3498db  天藍
   深色  #2c3e50  石板
   綠色  #27ae60  知識節點綠
   紅色  #e74c3c  警示
   背景  #f8f9fa  淨白
*/

/* ── 字型 ── */
QWidget {
    font-family: "Microsoft JhengHei", "Segoe UI", sans-serif;
    font-size: 10pt;
    color: #2c3e50;
}

/* ── 輸入元件 ── */
QLineEdit, QTextEdit {
    background: #ffffff;
    border: 1px solid #ced4da;
    border-radius: 4px;
    padding: 4px 8px;
    selection-background-color: #3498db;
    selection-color: #ffffff;
    color: #2c3e50;
}
QLineEdit:focus, QTextEdit:focus {
    border: 1.5px solid #3498db;
    background: #fdfdff;
}
QLineEdit:read-only, QTextEdit[readOnly="true"] {
    background: #f4f6f8;
    border-color: #dee2e6;
    color: #4a5568;
}

/* ── 下拉選單 ── */
QComboBox {
    background: #ffffff;
    border: 1px solid #ced4da;
    border-radius: 4px;
    padding: 4px 8px;
    min-height: 26px;
    color: #2c3e50;
}
QComboBox:focus { border: 1.5px solid #3498db; }
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 22px;
    border-left: 1px solid #dee2e6;
    border-top-right-radius: 4px;
    border-bottom-right-radius: 4px;
    background: #f4f6f8;
}
QComboBox QAbstractItemView {
    background: #ffffff;
    border: 1px solid #ced4da;
    selection-background-color: #d6eaf8;
    selection-color: #1a5276;
    outline: none;
    padding: 2px;
}

/* ── 按鈕（白底 + 石板色邊框） ── */
QPushButton {
    background: #ffffff;
    color: #2c3e50;
    border: 1.5px solid #aab4be;
    border-radius: 4px;
    padding: 5px 18px;
    font-weight: bold;
    min-height: 28px;
}
QPushButton:hover  { background: #eaf4fb; border-color: #3498db; color: #1a5276; }
QPushButton:pressed { background: #d6eaf8; border-color: #2980b9; }
QPushButton:disabled { background: #f4f6f8; color: #aab4be; border-color: #dee2e6; }

/* primary — 天藍（借自 Graph_RAG_test btn-primary） */
QPushButton#primary {
    background: #3498db;
    color: #ffffff;
    border: none;
    min-height: 28px;
}
QPushButton#primary:hover   { background: #2980b9; }
QPushButton#primary:pressed { background: #1f618d; }
QPushButton#primary:disabled { background: #85c1e9; color: #eaf4fb; }

/* success — 知識節點綠 */
QPushButton#success {
    background: #27ae60;
    color: #ffffff;
    border: none;
}
QPushButton#success:hover   { background: #219a52; }
QPushButton#success:pressed { background: #1a7a41; }
QPushButton#success:disabled { background: #82c09a; color: #e8f5e9; }

/* danger — 警示紅 */
QPushButton#danger {
    background: #e74c3c;
    color: #ffffff;
    border: none;
}
QPushButton#danger:hover { background: #c0392b; }

/* ── GroupBox ── */
QGroupBox {
    background: #ffffff;
    border: 1px solid #dee2e6;
    border-radius: 6px;
    margin-top: 14px;
    padding: 6px 10px 8px 10px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 6px;
    color: #2980b9;
    font-weight: bold;
    font-size: 10pt;
    background: #ffffff;
}

/* ── 捲動區 ── */
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical {
    background: #ecf0f1;
    width: 8px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #aab4be;
    border-radius: 4px;
    min-height: 24px;
}
QScrollBar::handle:vertical:hover { background: #7f8c8d; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }

/* ── 分頁標籤 ── */
QTabWidget::pane {
    border: 1px solid #dee2e6;
    border-radius: 0 5px 5px 5px;
    background: #ffffff;
}
QTabBar::tab {
    background: #ecf0f1;
    color: #4a5568;
    border: 1px solid #ced4da;
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    padding: 5px 14px;
    margin-right: 2px;
    font-weight: bold;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #2980b9;
    border-color: #dee2e6;
}
QTabBar::tab:hover:!selected { background: #d6eaf8; color: #2980b9; }

/* ── 進度條 ── */
QProgressBar {
    border: 1px solid #ced4da;
    border-radius: 4px;
    background: #ecf0f1;
    text-align: center;
    height: 14px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #3498db, stop:1 #5dade2);
    border-radius: 3px;
}

/* ── 核取方塊 ── */
QCheckBox { spacing: 8px; }
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 1.5px solid #aab4be;
    border-radius: 3px;
    background: white;
}
QCheckBox::indicator:checked {
    background: #27ae60;
    border-color: #219a52;
}
QCheckBox::indicator:hover { border-color: #3498db; }

/* ── 分隔線 ── */
QFrame[frameShape="4"] { color: #dee2e6; }

/* ── Splitter ── */
QSplitter::handle { background: #dee2e6; width: 3px; }
QSplitter::handle:hover { background: #3498db; }

/* ════ 版面容器（ID 選擇器不會向下傳遞） ════ */

/* 主視窗底色 */
#central   { background: #f8f9fa; }

/* 頂部導覽列 */
#topBar {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
        stop:0 #2c3e50, stop:1 #34495e);
    border: none;
}

/* 載入 / 表單頁面底色 */
#pageLoading { background: #f8f9fa; }
#formInner   { background: #f8f9fa; }
#pageResult  { background: #f8f9fa; }

/* 結果頁容器卡片 */
#statusBar, #confirmBar {
    background: #ffffff;
    border: 1px solid #dee2e6;
    border-radius: 6px;
}
#leftPanel {
    background: #ffffff;
    border: 1px solid #dee2e6;
    border-radius: 7px;
}

/* 工作職能分頁內容區 */
#taskTabWidget { background: #ffffff; }
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
# PDF 解析執行緒
# ─────────────────────────────────────────

class ParseThread(QThread):
    progress = pyqtSignal(str)
    done     = pyqtSignal(int, int)   # ok_count, err_count

    def __init__(self, pdf_paths: list, json_dir: Path):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.json_dir  = json_dir

    def run(self):
        try:
            from pdf_parser_v2 import parse_pdf_to_json
        except ImportError:
            self.progress.emit("✗ pdfplumber 未安裝，請執行：pip install pdfplumber")
            self.done.emit(0, len(self.pdf_paths))
            return

        ok = err = 0
        for path_str in self.pdf_paths:
            p   = Path(path_str)
            out = self.json_dir / (p.stem + ".json")
            try:
                self.progress.emit(f"解析中：{p.name} ...")
                parse_pdf_to_json(str(p), str(out))
                self.progress.emit(f"  ✓ {p.name}")
                ok += 1
            except Exception as e:
                self.progress.emit(f"  ✗ {p.name} 失敗：{e}")
                err += 1
        self.done.emit(ok, err)


# ─────────────────────────────────────────
# 資料管理對話框
# ─────────────────────────────────────────

class DataManagerDialog(QDialog):
    """管理 raw_pdf / parsed_json_v2 資料，並觸發重建索引。"""
    rebuild_requested = pyqtSignal()

    def __init__(self, rag: WizardRAG, parent=None):
        super().__init__(parent)
        self.rag  = rag
        self._raw_dir  = rag.json_dir.parent / "raw_pdf"
        self._json_dir = rag.json_dir
        self._parse_thread: Optional[ParseThread] = None

        self.setWindowTitle("資料管理")
        self.setMinimumSize(660, 500)
        self._build_ui()
        self._refresh_list()

    # ── UI ────────────────────────────────

    def _build_ui(self):
        v = QVBoxLayout(self)
        v.setSpacing(10)
        v.setContentsMargins(14, 12, 14, 12)

        # ── PDF 清單 ─────────────────────────
        v.addWidget(QLabel("raw_pdf 資料夾中的 PDF（勾選要操作的項目）："))

        self._search = QLineEdit()
        self._search.setPlaceholderText("搜尋 PDF 名稱...")
        self._search.setClearButtonEnabled(True)
        self._search.textChanged.connect(self._on_search)
        v.addWidget(self._search)

        self._list = QListWidget()
        v.addWidget(self._list, 1)

        # ── 清單操作列 ───────────────────────
        row1 = QHBoxLayout()
        btn_add = QPushButton("新增 PDF")
        btn_add.clicked.connect(self._on_add)

        btn_del = QPushButton("刪除選取")
        btn_del.setObjectName("danger")
        btn_del.clicked.connect(self._on_delete)

        btn_all  = QPushButton("全選")
        btn_all.clicked.connect(self._check_all)
        btn_none = QPushButton("全不選")
        btn_none.clicked.connect(self._check_none)

        row1.addWidget(btn_add)
        row1.addWidget(btn_del)
        row1.addStretch()
        row1.addWidget(btn_all)
        row1.addWidget(btn_none)
        v.addLayout(row1)

        # ── 分隔 ─────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        v.addWidget(sep)

        # ── 記錄區 ───────────────────────────
        v.addWidget(QLabel("操作記錄："))
        self._log = QTextEdit()
        self._log.setReadOnly(True)
        self._log.setFixedHeight(130)
        self._log.setFont(QFont("Consolas", 9))
        v.addWidget(self._log)

        # ── 底部動作列 ───────────────────────
        row2 = QHBoxLayout()
        self._btn_parse = QPushButton("解析勾選的 PDF → JSON")
        self._btn_parse.setObjectName("primary")
        self._btn_parse.clicked.connect(self._on_parse)

        self._btn_rebuild = QPushButton("重建向量索引")
        self._btn_rebuild.setObjectName("success")
        self._btn_rebuild.clicked.connect(self._on_rebuild)

        btn_close = QPushButton("關閉")
        btn_close.clicked.connect(self.close)

        row2.addWidget(self._btn_parse)
        row2.addWidget(self._btn_rebuild)
        row2.addStretch()
        row2.addWidget(btn_close)
        v.addLayout(row2)

    # ── 清單管理 ──────────────────────────

    def _refresh_list(self):
        self._list.clear()
        self._raw_dir.mkdir(parents=True, exist_ok=True)
        pdfs = sorted(self._raw_dir.glob("*.pdf"))
        if not pdfs:
            item = QListWidgetItem("（資料夾中目前沒有 PDF）")
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEnabled)
            self._list.addItem(item)
            return
        for pdf in pdfs:
            parsed = (self._json_dir / (pdf.stem + ".json")).exists()
            label = f"{'✓' if parsed else '✗'}  {pdf.name}"
            item = QListWidgetItem(label)
            item.setCheckState(Qt.CheckState.Unchecked)
            item.setData(Qt.ItemDataRole.UserRole, str(pdf))
            if not parsed:
                item.setForeground(QColor("#e74c3c"))
            self._list.addItem(item)

    def _checked_paths(self) -> list:
        result = []
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                p = item.data(Qt.ItemDataRole.UserRole)
                if p:
                    result.append(p)
        return result

    def _check_all(self):
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item.data(Qt.ItemDataRole.UserRole):
                item.setCheckState(Qt.CheckState.Checked)

    def _check_none(self):
        for i in range(self._list.count()):
            item = self._list.item(i)
            item.setCheckState(Qt.CheckState.Unchecked)

    def _on_search(self, text: str):
        kw = text.strip().lower()
        for i in range(self._list.count()):
            item = self._list.item(i)
            item.setHidden(bool(kw) and kw not in item.text().lower())

    # ── 操作 ──────────────────────────────

    def _on_add(self):
        self._raw_dir.mkdir(parents=True, exist_ok=True)
        paths, _ = QFileDialog.getOpenFileNames(
            self, "選擇 PDF 檔案", str(Path.home()), "PDF 檔案 (*.pdf)"
        )
        if not paths:
            return
        copied = 0
        for src in paths:
            dst = self._raw_dir / Path(src).name
            if dst.exists():
                self._log.append(f"⚠ 已存在，略過：{Path(src).name}")
            else:
                shutil.copy2(src, dst)
                self._log.append(f"✓ 已複製：{Path(src).name}")
                copied += 1
        if copied:
            self._refresh_list()

    def _on_delete(self):
        paths = self._checked_paths()
        if not paths:
            QMessageBox.information(self, "提示", "請先勾選要刪除的 PDF")
            return
        names = "\n".join(Path(p).name for p in paths)
        reply = QMessageBox.question(
            self, "確認刪除",
            f"確定要刪除以下 {len(paths)} 個 PDF 及其對應 JSON？\n\n{names}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        for p in paths:
            pdf = Path(p)
            pdf.unlink(missing_ok=True)
            self._log.append(f"🗑 已刪除 PDF：{pdf.name}")
            json_f = self._json_dir / (pdf.stem + ".json")
            if json_f.exists():
                json_f.unlink()
                self._log.append(f"🗑 已刪除 JSON：{json_f.name}")
        self._refresh_list()

    def _on_parse(self):
        paths = self._checked_paths()
        if not paths:
            QMessageBox.information(self, "提示", "請先勾選要解析的 PDF")
            return
        self._btn_parse.setEnabled(False)
        self._btn_rebuild.setEnabled(False)
        self._log.append(f"\n▶ 開始解析 {len(paths)} 個 PDF...")
        self._parse_thread = ParseThread(paths, self._json_dir)
        self._parse_thread.progress.connect(self._log.append)
        self._parse_thread.done.connect(self._on_parse_done)
        self._parse_thread.start()

    def _on_parse_done(self, ok: int, err: int):
        self._log.append(f"── 完成：{ok} 成功，{err} 失敗 ──")
        self._btn_parse.setEnabled(True)
        self._btn_rebuild.setEnabled(True)
        self._refresh_list()

    def _on_rebuild(self):
        reply = QMessageBox.question(
            self, "重建向量索引",
            "確定要重建向量索引？\n（需要數分鐘，完成後程式將回到載入畫面）",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._log.append("\n▶ 送出重建請求...")
        self.rebuild_requested.emit()
        self.close()


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
        central.setObjectName("central")
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
        bar.setObjectName("topBar")
        bar.setFixedHeight(52)
        h = QHBoxLayout(bar)
        h.setContentsMargins(20, 0, 20, 0)

        dot = QLabel("●")
        dot.setStyleSheet("color:#5dade2; font-size:10pt; margin-right:4px;")
        h.addWidget(dot)

        title = QLabel("職能說明書精靈")
        title.setFont(QFont("Microsoft JhengHei", 14, QFont.Weight.Bold))
        title.setStyleSheet("color:white; letter-spacing:1px;")
        h.addWidget(title)
        h.addStretch()

        btn_data = QPushButton("資料管理")
        btn_data.setFixedHeight(28)
        btn_data.setStyleSheet(
            "QPushButton { background:rgba(255,255,255,0.12); color:white; "
            "border:1px solid rgba(255,255,255,0.28); border-radius:4px; "
            "padding:2px 12px; font-size:9pt; font-weight:bold; }"
            "QPushButton:hover { background:rgba(255,255,255,0.22); }"
            "QPushButton:pressed { background:rgba(255,255,255,0.32); }"
        )
        btn_data.clicked.connect(self._open_data_manager)
        h.addWidget(btn_data)

        self._status_label = QLabel("初始化中...")
        self._status_label.setStyleSheet(
            "color:#aed6f1; font-size:9pt; "
            "background:rgba(255,255,255,0.10); "
            "border-radius:4px; padding:2px 10px;"
        )
        h.addWidget(self._status_label)

        return bar

    def _make_loading_page(self) -> QWidget:
        w = QWidget()
        w.setObjectName("pageLoading")
        v = QVBoxLayout(w)
        v.setAlignment(Qt.AlignmentFlag.AlignCenter)
        v.setSpacing(18)

        icon = QLabel("⚙")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setStyleSheet("font-size:40pt; color:#3498db;")
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
        inner = QWidget()
        inner.setObjectName("formInner")
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
        w.setObjectName("pageResult")
        v = QVBoxLayout(w)
        v.setContentsMargins(14, 10, 14, 10)
        v.setSpacing(8)

        # ── 頂部狀態列 ─────────────────────────────
        status_bar = QFrame()
        status_bar.setObjectName("statusBar")
        status_bar.setFixedHeight(40)
        sh = QHBoxLayout(status_bar)
        sh.setContentsMargins(12, 0, 12, 0)

        self._result_status = QLabel("")
        self._result_status.setFont(QFont("Microsoft JhengHei", 10))
        self._result_status.setStyleSheet(
            "color:#2980b9; font-weight:bold; background:transparent; border:none;"
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
        left.setObjectName("leftPanel")
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 10, 12, 10)
        lv.setSpacing(8)

        lbl_match = QLabel("相似職能基準")
        lbl_match.setStyleSheet(
            "font-weight:bold; color:#2980b9; font-size:9pt; "
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
            "font-weight:bold; color:#2980b9; font-size:9pt; "
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
        _task_tab_w.setObjectName("taskTabWidget")
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
        confirm_bar.setObjectName("confirmBar")
        confirm_bar.setFixedHeight(42)
        ch = QHBoxLayout(confirm_bar)
        ch.setContentsMargins(14, 0, 14, 0)

        self._confirm_check = QCheckBox(
            "我已確認以上缺口分析結果正確無誤，同意匯出職能說明書"
        )
        self._confirm_check.setFont(QFont("Microsoft JhengHei", 10))
        self._confirm_check.setStyleSheet(
            "QCheckBox { color:#2980b9; font-weight:bold; "
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

    def _open_data_manager(self):
        dlg = DataManagerDialog(self.rag, self)
        dlg.rebuild_requested.connect(self._on_force_rebuild)
        dlg.exec()

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
