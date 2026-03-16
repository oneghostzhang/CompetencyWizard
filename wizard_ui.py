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
    QDialog, QRadioButton,
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
# 職能基準選擇對話框
# ─────────────────────────────────────────

class StandardSelectorDialog(QDialog):
    """從已載入的職能基準中選擇範本，供表單預填使用。"""

    def __init__(self, rag: WizardRAG, parent=None):
        super().__init__(parent)
        self.rag = rag
        self.selected_data: Optional[dict] = None
        self.setWindowTitle("選擇職能基準範本")
        self.setMinimumSize(600, 520)
        self._build_ui()
        self._populate_list()

    def _build_ui(self):
        v = QVBoxLayout(self)
        v.setSpacing(10)
        v.setContentsMargins(14, 12, 14, 12)

        hint = QLabel("選擇一個職能基準作為起始範本，系統將自動預填 5W2H 欄位，您可隨時修改。")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#555; font-size:9pt; padding:4px 0;")
        v.addWidget(hint)

        self._search = QLineEdit()
        self._search.setPlaceholderText("搜尋職能名稱 / 代碼...")
        self._search.setClearButtonEnabled(True)
        self._search.textChanged.connect(self._on_search)
        v.addWidget(self._search)

        self._list = QListWidget()
        self._list.setAlternatingRowColors(True)
        self._list.currentItemChanged.connect(self._on_selection_changed)
        self._list.itemDoubleClicked.connect(self._on_confirm)
        v.addWidget(self._list, 1)

        # 預覽區
        preview_lbl = QLabel("工作說明：")
        preview_lbl.setStyleSheet("font-weight:bold; font-size:9pt;")
        v.addWidget(preview_lbl)
        self._preview = QTextEdit()
        self._preview.setReadOnly(True)
        self._preview.setFixedHeight(90)
        self._preview.setPlaceholderText("選取職能基準後顯示工作說明預覽...")
        v.addWidget(self._preview)

        # 按鈕列
        btn_row = QHBoxLayout()
        btn_cancel = QPushButton("取消")
        btn_cancel.setFixedHeight(34)
        btn_cancel.clicked.connect(self.reject)
        self._btn_load = QPushButton("載入此職能基準 →")
        self._btn_load.setObjectName("primary")
        self._btn_load.setFixedHeight(34)
        self._btn_load.setEnabled(False)
        self._btn_load.clicked.connect(self._on_confirm)
        btn_row.addWidget(btn_cancel)
        btn_row.addStretch()
        btn_row.addWidget(self._btn_load)
        v.addLayout(btn_row)

    def _populate_list(self):
        self._list.clear()
        standards = getattr(self.rag, "_standards", {})
        items = sorted(
            standards.items(),
            key=lambda kv: kv[1].get("metadata", {}).get("name", "")
        )
        for code, data in items:
            name = data.get("metadata", {}).get("name", "") or code
            level = data.get("basic_info", {}).get("level", "")
            level_str = f"  Lv.{level}" if level else ""
            item = QListWidgetItem(f"{name}{level_str}  （{code}）")
            item.setData(Qt.ItemDataRole.UserRole, code)
            self._list.addItem(item)

    def _on_search(self, text: str):
        kw = text.strip().lower()
        for i in range(self._list.count()):
            item = self._list.item(i)
            item.setHidden(bool(kw) and kw not in item.text().lower())

    def _on_selection_changed(self, current, _previous):
        if current is None:
            self._preview.clear()
            self._btn_load.setEnabled(False)
            return
        code = current.data(Qt.ItemDataRole.UserRole)
        data = self.rag.get_standard(code)
        if data:
            desc = data.get("basic_info", {}).get("job_description", "（無說明）")
            self._preview.setPlainText(desc[:400])
        self._btn_load.setEnabled(True)

    def _on_confirm(self):
        item = self._list.currentItem()
        if not item:
            return
        code = item.data(Qt.ItemDataRole.UserRole)
        self.selected_data = self.rag.get_standard(code)
        self.accept()


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
# 職能基準候選選擇對話框
# ─────────────────────────────────────────

class StandardPickerDialog(QDialog):
    """分析完成後讓使用者從 Top-K 候選中確認或選擇正確的職能基準"""

    def __init__(self, matched_standards: list, rag, parent=None):
        super().__init__(parent)
        self.setWindowTitle("請確認職能基準")
        self.setMinimumWidth(580)
        self.setModal(True)
        self.selected_index = 0
        self._rag = rag
        self._standards = matched_standards
        self._radio_btns = []
        self._build_ui()

    def _build_ui(self):
        lv = QVBoxLayout(self)
        lv.setSpacing(10)

        lbl_hint = QLabel(
            "系統找到以下最相似的職能基準，請確認或選擇最符合您職務的項目，\n"
            "再進行逐項確認。若職稱或級別不對，請選擇其他候選。"
        )
        lbl_hint.setWordWrap(True)
        lbl_hint.setStyleSheet("color:#555; padding:4px 0;")
        lv.addWidget(lbl_hint)

        for i, r in enumerate(self._standards):
            std_data  = self._rag.get_standard(r["standard_code"]) or {}
            basic     = std_data.get("basic_info") or {}
            level     = basic.get("level", "—")
            tasks     = std_data.get("competency_tasks") or []
            task_cnt  = len(tasks)

            frame = QFrame()
            frame.setFrameShape(QFrame.Shape.StyledPanel)
            frame.setStyleSheet(
                "QFrame { border:1px solid #ced4da; border-radius:6px; "
                "background:#ffffff; padding:6px; }"
            )
            hl = QHBoxLayout(frame)
            hl.setSpacing(10)

            rb = QRadioButton()
            rb.setChecked(i == 0)
            hl.addWidget(rb)
            self._radio_btns.append(rb)

            # 資訊文字
            warn_html = ""
            if task_cnt <= 2:
                warn_html = (
                    f"　<span style='color:#e67e22; font-weight:bold;'>"
                    f"⚠️ 任務僅 {task_cnt} 項，可能為助理/初階職位</span>"
                )
            info = QLabel(
                f"<b>{r['standard_name']}</b>（{r['standard_code']}）{warn_html}<br>"
                f"<span style='color:#555;'>級別：Level {level}　"
                f"工作任務：{task_cnt} 項　"
                f"相似度：{r['score']:.3f}</span>"
            )
            info.setTextFormat(Qt.TextFormat.RichText)
            info.setWordWrap(True)
            hl.addWidget(info, 1)

            # 點擊整個 frame 也能選中
            frame.mousePressEvent = (lambda e, btn=rb: btn.setChecked(True))

            lv.addWidget(frame)

        lv.addSpacing(6)

        bb = QHBoxLayout()
        bb.addStretch()
        btn_ok = QPushButton("使用此基準，開始確認 →")
        btn_ok.setObjectName("primaryBtn")
        btn_ok.setFixedHeight(32)
        btn_ok.clicked.connect(self._on_confirm)
        bb.addWidget(btn_ok)
        lv.addLayout(bb)

    def _on_confirm(self):
        for i, rb in enumerate(self._radio_btns):
            if rb.isChecked():
                self.selected_index = i
                break
        self.accept()


# ─────────────────────────────────────────
# 職能基準逐項確認精靈
# ─────────────────────────────────────────

class StandardAdoptionWizard(QDialog):
    """
    分析完成後開啟，讓員工逐項確認工作任務、知識、技能，
    系統依確認結果重新計算缺口與完整度。
    """

    def __init__(self, report: GapReport, parent=None):
        super().__init__(parent)
        self.report   = report
        self.std_data = report.best_standard_data or {}
        self._task_checks:  list = []
        self._know_checks:  list = []
        self._skill_checks: list = []
        self.confirmed_tasks:     list = []
        self.confirmed_knowledge: list = []
        self.confirmed_skills:    list = []
        self.setWindowTitle(f"職能基準確認 — {report.best_standard_name}")
        self.setMinimumSize(780, 600)
        self._build_ui()

    # ── UI 建構 ──────────────────────────────────

    def _build_ui(self):
        v = QVBoxLayout(self)
        v.setSpacing(8)
        v.setContentsMargins(14, 12, 14, 12)

        # 頂部資訊列
        info = QFrame()
        info.setStyleSheet(
            "QFrame { background:#e8f4fd; border:1px solid #bee3f8; border-radius:6px; }"
        )
        ih = QHBoxLayout(info)
        ih.setContentsMargins(12, 8, 12, 8)
        bi  = self.std_data.get("basic_info", {})
        lbl_std = QLabel(
            f"<b>最佳匹配職能基準：</b>{self.report.best_standard_name}"
            f"&nbsp;&nbsp;（{self.report.best_standard_code}）"
        )
        lbl_std.setStyleSheet("background:transparent; border:none; color:#1a202c;")
        lbl_score = QLabel(f"基準級別：<b>Level {bi.get('level', '—')}</b>")
        lbl_score.setStyleSheet("background:transparent; border:none; color:#555;")
        ih.addWidget(lbl_std, 1)
        ih.addWidget(lbl_score)
        v.addWidget(info)

        # 職務說明
        jd  = bi.get("job_description", "")
        if jd:
            lbl_jd = QLabel(jd[:220] + ("…" if len(jd) > 220 else ""))
            lbl_jd.setWordWrap(True)
            lbl_jd.setStyleSheet(
                "color:#4a5568; font-size:9pt; background:transparent; padding:2px 4px;"
            )
            v.addWidget(lbl_jd)

        # 三個分頁
        self._tabs = QTabWidget()
        self._tabs.addTab(self._build_task_tab(),          "📋 工作任務")
        self._tabs.addTab(self._build_item_tab("knowledge"), "📖 知識")
        self._tabs.addTab(self._build_item_tab("skills"),    "🔧 技能")
        v.addWidget(self._tabs, 1)

        # 說明文字
        lbl_hint = QLabel(
            "💡 系統已預先勾選所有項目（綠色為自動偵測確認）。"
            "請取消勾選您實際上不執行或不具備的項目，完成後按「確認採用 ✓」。"
        )
        lbl_hint.setWordWrap(True)
        lbl_hint.setStyleSheet(
            "color:#555; font-size:9pt; background:transparent; padding:4px 2px;"
        )
        v.addWidget(lbl_hint)

        # 按鈕列
        btn_row = QHBoxLayout()
        btn_all  = QPushButton("全選目前頁")
        btn_none = QPushButton("全不選目前頁")
        btn_all.clicked.connect(self._select_all)
        btn_none.clicked.connect(self._select_none)
        btn_skip = QPushButton("略過")
        btn_skip.clicked.connect(self.reject)
        btn_confirm = QPushButton("確認採用 ✓")
        btn_confirm.setObjectName("primary")
        btn_confirm.setFixedHeight(34)
        btn_confirm.clicked.connect(self._on_confirm)
        btn_row.addWidget(btn_all)
        btn_row.addWidget(btn_none)
        btn_row.addStretch()
        btn_row.addWidget(btn_skip)
        btn_row.addWidget(btn_confirm)
        v.addLayout(btn_row)

    def _make_scroll_widget(self) -> tuple:
        """建立帶 ScrollArea 的容器，回傳 (container, inner_vlay)"""
        container = QWidget()
        scroll    = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        vlay  = QVBoxLayout(inner)
        vlay.setContentsMargins(8, 8, 8, 8)
        vlay.setSpacing(4)
        scroll.setWidget(inner)
        outer = QVBoxLayout(container)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)
        return container, vlay

    def _build_task_tab(self) -> QWidget:
        """工作任務 tab — 依 main_responsibility 分組"""
        container, vlay = self._make_scroll_widget()
        tasks         = self.std_data.get("competency_tasks", [])
        covered_names = set(self.report.covered_tasks)
        current_resp  = None

        for task in tasks:
            resp = task.get("main_responsibility", "")
            if resp != current_resp:
                current_resp = resp
                grp = QLabel(f"▌ {resp}")
                grp.setStyleSheet(
                    "font-weight:bold; color:#2c3e50; font-size:9pt; "
                    "background:#f0f4f8; border-radius:3px; padding:3px 6px;"
                )
                grp.setContentsMargins(0, 6, 0, 2)
                vlay.addWidget(grp)

            tid    = task.get("task_id", "")
            name   = task.get("task_name", "")
            output = task.get("output", "") or ""
            label  = f"[{tid}]  {name}"
            if output:
                label += f"　→　{output[:50]}{'…' if len(output) > 50 else ''}"

            cb = QCheckBox(label)
            cb.setFont(QFont("Microsoft JhengHei", 9))
            is_covered = name in covered_names
            cb.setChecked(True)   # 預設全選，讓員工取消不具備的項目
            cb.setProperty("item_name", name)
            cb.setStyleSheet(
                "QCheckBox { color:#27ae60; background:transparent; }" if is_covered
                else "QCheckBox { color:#2980b9; background:transparent; }"
            )
            self._task_checks.append(cb)
            vlay.addWidget(cb)

        vlay.addStretch()
        return container

    def _build_item_tab(self, item_type: str) -> QWidget:
        """知識 / 技能 tab"""
        container, vlay = self._make_scroll_widget()

        if item_type == "knowledge":
            items         = self.std_data.get("competency_knowledge", [])
            covered_names = set(self.report.covered_knowledge)
            target        = self._know_checks
        else:
            items         = self.std_data.get("competency_skills", [])
            covered_names = set(self.report.covered_skills)
            target        = self._skill_checks

        for item in items:
            code = item.get("code", "")
            name = item.get("name", "")
            if not name:
                continue
            cb = QCheckBox(f"[{code}]  {name}")
            cb.setFont(QFont("Microsoft JhengHei", 9))
            is_covered = name in covered_names
            cb.setChecked(True)   # 預設全選，讓員工取消不具備的項目
            cb.setProperty("item_name", name)
            cb.setStyleSheet(
                "QCheckBox { color:#27ae60; background:transparent; }" if is_covered
                else "QCheckBox { color:#2980b9; background:transparent; }"
            )
            target.append(cb)
            vlay.addWidget(cb)

        vlay.addStretch()
        return container

    # ── 全選 / 全不選 ─────────────────────────────

    def _current_checks(self) -> list:
        return [self._task_checks, self._know_checks, self._skill_checks][
            self._tabs.currentIndex()
        ]

    def _select_all(self):
        for cb in self._current_checks():
            cb.setChecked(True)

    def _select_none(self):
        for cb in self._current_checks():
            cb.setChecked(False)

    # ── 確認採用 ──────────────────────────────────

    def _on_confirm(self):
        self.confirmed_tasks      = [
            cb.property("item_name") for cb in self._task_checks  if cb.isChecked()
        ]
        self.confirmed_knowledge  = [
            cb.property("item_name") for cb in self._know_checks  if cb.isChecked()
        ]
        self.confirmed_skills     = [
            cb.property("item_name") for cb in self._skill_checks if cb.isChecked()
        ]
        self.accept()


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

        # ── 快速載入提示區塊 ─────────────────────
        hint_bar = QFrame()
        hint_bar.setStyleSheet(
            "QFrame { background:#eaf4fb; border:1px solid #aed6f1; border-radius:6px; }"
        )
        hint_h = QHBoxLayout(hint_bar)
        hint_h.setContentsMargins(12, 8, 12, 8)
        hint_h.setSpacing(12)
        hint_icon = QLabel("💡")
        hint_icon.setStyleSheet("font-size:16pt; background:transparent; border:none;")
        hint_h.addWidget(hint_icon)
        hint_text = QLabel(
            "<b>從職能基準快速填寫</b><br>"
            "<span style='color:#555; font-size:9pt;'>"
            "選擇 ICAP 職能基準作為起始範本，系統自動預填工作任務、技能、產出等欄位，"
            "您只需微調個人情境即可完成 80–90% 的填寫。"
            "</span>"
        )
        hint_text.setTextFormat(Qt.TextFormat.RichText)
        hint_text.setWordWrap(True)
        hint_text.setStyleSheet("background:transparent; border:none;")
        hint_h.addWidget(hint_text, 1)
        self._btn_load_template = QPushButton("選擇範本 →")
        self._btn_load_template.setFixedHeight(32)
        self._btn_load_template.setFixedWidth(110)
        self._btn_load_template.setStyleSheet(
            "QPushButton { background:#3498db; color:white; border:none; "
            "border-radius:4px; font-weight:bold; font-size:9pt; padding:0 8px; }"
            "QPushButton:hover { background:#2980b9; }"
            "QPushButton:pressed { background:#2471a3; }"
            "QPushButton:disabled { background:#aaa; }"
        )
        self._btn_load_template.setEnabled(False)   # 初始化完成前禁用
        self._btn_load_template.clicked.connect(self._on_load_from_standard)
        hint_h.addWidget(self._btn_load_template)
        v.addWidget(hint_bar)

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
        freq_widget = QWidget()
        freq_layout = QHBoxLayout(freq_widget)
        freq_layout.setContentsMargins(0, 0, 0, 0)
        freq_layout.setSpacing(16)
        self._when_checkboxes: dict = {}
        for opt in ["每日", "每週", "每月", "每季", "專案型（不固定）", "其他"]:
            cb = QCheckBox(opt)
            self._when_checkboxes[opt] = cb
            freq_layout.addWidget(cb)
        freq_layout.addStretch()
        f4.addRow("執行頻率：", freq_widget)
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
        self._btn_adoption = QPushButton("📋 重新確認職能")
        self._btn_adoption.setFixedHeight(28)
        self._btn_adoption.setEnabled(False)
        self._btn_adoption.clicked.connect(self._open_adoption_wizard)

        self._btn_export = QPushButton("匯出 Excel")
        self._btn_export.setObjectName("success")
        self._btn_export.setFixedHeight(28)
        self._btn_export.setEnabled(False)
        self._btn_export.clicked.connect(self._on_export)

        sh.addWidget(self._result_status, 1)
        sh.addWidget(self._btn_adoption)
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
            self._btn_load_template.setEnabled(True)
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

    def _on_load_from_standard(self):
        dlg = StandardSelectorDialog(self.rag, self)
        if dlg.exec() == QDialog.DialogCode.Accepted and dlg.selected_data:
            self._fill_form_from_standard(dlg.selected_data)

    def _fill_form_from_standard(self, data: dict):
        """將職能基準 JSON 資料對應預填到 5W2H 表單欄位。"""
        meta = data.get("metadata", {})
        bi   = data.get("basic_info", {})
        tasks    = data.get("competency_tasks", [])
        skills   = data.get("competency_skills", [])

        # What — 工作任務（數字編號列表）
        task_lines = [
            f"{i+1}. {t.get('task_name', '')}"
            for i, t in enumerate(tasks) if t.get("task_name")
        ]
        self._what_tasks.setPlainText("\n".join(task_lines))

        # What — 工作產出（前 3 筆任務產出）
        outputs = [t.get("output", "") for t in tasks if t.get("output")]
        self._what_outputs.setText("、".join(outputs[:3]))

        # Why — 工作目的（job_description 前 120 字）
        desc = bi.get("job_description", "")
        self._why_purpose.setText(desc[:120])

        # Who — 自身角色（職能基準名稱）
        self._who_role.setText(meta.get("name", ""))

        # Who — 協作對象（無對應欄位，保留空白讓使用者自填）

        # Where — 工作環境（第一個行業別）
        industry = bi.get("industry", [])
        if isinstance(industry, list) and industry:
            self._where_env.setText(industry[0])
        elif isinstance(industry, str):
            self._where_env.setText(industry)

        # How — 技能/工具（所有技能名稱）
        skill_names = [s.get("name", "") for s in skills if s.get("name")]
        self._how_skills.setPlainText("、".join(skill_names))

        # How Much — 績效指標（第一項行為指標）
        for t in tasks:
            behaviors = t.get("behaviors", [])
            if behaviors:
                self._how_much.setText(behaviors[0][:100])
                break

    # ─── 表單操作 ─────────────────────────────

    def _clear_form(self):
        self._what_tasks.clear()
        self._what_outputs.clear()
        self._why_purpose.clear()
        self._who_role.clear()
        self._who_collaborate.clear()
        for cb in self._when_checkboxes.values():
            cb.setChecked(False)
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
            when_frequency="、".join(
                opt for opt, cb in self._when_checkboxes.items() if cb.isChecked()
            ),
            where_environment=self._where_env.text().strip(),
            how_skills=self._how_skills.toPlainText().strip(),
            how_much_kpi=self._how_much.text().strip(),
        )

    def _validate_input(self, ui: "UserInput5W2H") -> bool:
        """驗證 5W2H 輸入；回傳 True 表示可繼續分析。"""
        missing_required = []
        missing_suggested = []

        if not ui.what_tasks:
            missing_required.append("• 工作任務（What）")
        if not ui.why_purpose:
            missing_suggested.append("• 工作目的（Why）")
        if not ui.who_role:
            missing_suggested.append("• 自身角色（Who）")
        if not ui.how_skills:
            missing_suggested.append("• 技能 / 工具（How）")

        if missing_required:
            msg = "以下必填欄位尚未填寫，請補充後再分析：\n\n" + "\n".join(missing_required)
            if missing_suggested:
                msg += "\n\n以下欄位也建議補充以提高分析準確度：\n" + "\n".join(missing_suggested)
            QMessageBox.warning(self, "欄位未填寫", msg)
            return False

        if missing_suggested:
            msg = "以下欄位尚未填寫，補充後分析結果會更準確：\n\n" + "\n".join(missing_suggested)
            msg += "\n\n是否仍要繼續分析？"
            reply = QMessageBox.question(
                self, "欄位未填寫", msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            return reply == QMessageBox.StandardButton.Yes

        return True

    def _on_analyze(self):
        if not self.analyzer:
            QMessageBox.warning(self, "未就緒", "RAG 尚未初始化，請稍候")
            return
        ui = self._collect_input()
        if not self._validate_input(ui):
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
        # 先讓使用者確認/選擇職能基準，再開啟精靈
        if report.matched_standards:
            self._open_standard_picker()

    def _open_standard_picker(self):
        """顯示 Top-K 候選基準讓使用者確認，再開啟逐項確認精靈"""
        if not self.report or not self.report.matched_standards:
            return
        dlg = StandardPickerDialog(self.report.matched_standards, self.rag, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        idx = dlg.selected_index
        if idx != 0:
            # 使用者選了非預設基準，更新 report 與 UI 下拉
            chosen = self.report.matched_standards[idx]
            self.report.best_standard_code = chosen["standard_code"]
            self.report.best_standard_name = chosen["standard_name"]
            self.report.best_standard_data = self.rag.get_standard(chosen["standard_code"])
            self._match_combo.blockSignals(True)
            self._match_combo.setCurrentIndex(idx)
            self._match_combo.blockSignals(False)
            self._on_match_selected(idx)
        self._open_adoption_wizard()

    def _on_analyze_error(self, msg: str):
        self._btn_analyze.setEnabled(True)
        self._btn_reanalyze.setEnabled(True)
        self._status_label.setText("分析錯誤")
        QMessageBox.critical(self, "分析失敗", msg)

    # ─── 職能確認精靈 ─────────────────────────────

    def _open_adoption_wizard(self):
        """開啟 StandardAdoptionWizard，確認後重建缺口報告"""
        if not self.report or not self.report.best_standard_data:
            QMessageBox.information(self, "無法開啟", "請先執行分析後再確認職能項目。")
            return

        # 任務數過少警告
        std_data  = self.report.best_standard_data
        task_cnt  = len(std_data.get("competency_tasks") or [])
        std_name  = self.report.best_standard_name
        if task_cnt <= 2:
            reply = QMessageBox.warning(
                self, "工作任務項目偏少",
                f"職能基準「{std_name}」僅包含 {task_cnt} 個工作任務，\n"
                "可能為助理或初階職位，與您實際工作內容可能不符。\n\n"
                "建議關閉後，在上方下拉選單切換至其他候選基準。\n\n"
                "是否仍要繼續確認此基準？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                return

        dlg = StandardAdoptionWizard(self.report, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._rebuild_from_confirmation(
                dlg.confirmed_tasks,
                dlg.confirmed_knowledge,
                dlg.confirmed_skills,
            )

    def _rebuild_from_confirmation(
        self,
        confirmed_task_names: list,
        confirmed_k_names: list,
        confirmed_s_names: list,
    ):
        """依員工確認結果重建 covered/gap 清單與完整度"""
        from gap_analyzer import GapItem

        std_data      = self.report.best_standard_data or {}
        all_tasks     = std_data.get("competency_tasks", [])
        all_knowledge = std_data.get("competency_knowledge", [])
        all_skills    = std_data.get("competency_skills", [])

        confirmed_t = set(confirmed_task_names)
        confirmed_k = set(confirmed_k_names)
        confirmed_s = set(confirmed_s_names)

        # 更新 covered
        self.report.covered_tasks     = confirmed_task_names[:]
        self.report.covered_knowledge = confirmed_k_names[:]
        self.report.covered_skills    = confirmed_s_names[:]

        # 重建 gap（標準中有但員工未勾選的項目）
        self.report.gap_tasks = [
            GapItem(
                category="task",
                code=t.get("task_id", ""),
                name=t.get("task_name", ""),
                description=t.get("output", "") or "",
            )
            for t in all_tasks if t.get("task_name", "") not in confirmed_t
        ]
        self.report.gap_knowledge = [
            GapItem(
                category="knowledge",
                code=k.get("code", ""),
                name=k.get("name", ""),
                description=k.get("description", "") or "",
            )
            for k in all_knowledge if k.get("name", "") not in confirmed_k
        ]
        self.report.gap_skills = [
            GapItem(
                category="skill",
                code=s.get("code", ""),
                name=s.get("name", ""),
                description=s.get("description", "") or "",
            )
            for s in all_skills if s.get("name", "") not in confirmed_s
        ]

        # 重算完整度
        total = len(all_tasks) + len(all_knowledge) + len(all_skills)
        confirmed_count = (
            len(confirmed_task_names) + len(confirmed_k_names) + len(confirmed_s_names)
        )
        self.report.completeness_score = (
            round(confirmed_count / total * 100, 1) if total > 0 else 0.0
        )

        # 刷新結果頁顯示
        score = self.report.completeness_score
        self._result_status.setText(
            f"最佳匹配：{self.report.best_standard_name}"
            f"  ｜  完整度：{score}%（員工確認）"
        )
        if self.analyzer:
            self._tab_gap.setPlainText(self.analyzer.get_summary_text(self.report))

    # ─── 結果顯示 ─────────────────────────────

    def _populate_results(self, report: GapReport):
        self._confirm_check.setChecked(False)
        self._btn_export.setEnabled(False)
        self._btn_adoption.setEnabled(bool(report.best_standard_data))

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

        self._result_status.setText(
            f"找到最相似職能基準：{report.best_standard_name}"
            f"  ｜  請確認符合您工作的項目 👇"
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
        if not self._validate_input(ui):
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
