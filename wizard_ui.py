"""
competency_wizard/wizard_ui.py
職能說明書精靈 — PyQt6 UI  v2.0
流程：初始化 → 搜索職業 → 編輯職能基準書 → 填寫工作詳情 → LLM建議確認 → 補充匯出
"""

import hashlib
import shutil
import sys
from enum import Enum, auto
from pathlib import Path
from typing import Optional, List, Dict, TypedDict

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextEdit, QLineEdit, QPushButton, QProgressBar,
    QStackedWidget, QGroupBox, QFileDialog, QMessageBox,
    QScrollArea, QFrame, QComboBox, QCheckBox, QTabWidget,
    QDialog, QTextBrowser, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QListWidget, QListWidgetItem,
    QSpinBox, QSplitter,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QColor

from wizard_rag import WizardRAG
from ai_chat import create_persistent_worker


class SuggestEntry(TypedDict):
    checks: list[tuple[QCheckBox, QLineEdit]]
    extra: QTextEdit


# ─────────────────────────────────────────
# 全域樣式
# ─────────────────────────────────────────

APP_STYLE = """
QWidget {
    font-family: "Microsoft JhengHei", "Segoe UI", sans-serif;
    font-size: 10pt;
    color: #2c3e50;
}
QLineEdit, QTextEdit {
    background: #ffffff;
    border: 1px solid #ced4da;
    border-radius: 4px;
    padding: 4px 8px;
    selection-background-color: #3498db;
    selection-color: #ffffff;
    color: #2c3e50;
}
QLineEdit:focus, QTextEdit:focus { border: 1.5px solid #3498db; background: #fdfdff; }
QLineEdit:read-only, QTextEdit[readOnly="true"] { background: #f4f6f8; border-color: #dee2e6; color: #4a5568; }
QComboBox {
    background: #ffffff; border: 1px solid #ced4da; border-radius: 4px;
    padding: 4px 8px; min-height: 26px; color: #2c3e50;
}
QComboBox:focus { border: 1.5px solid #3498db; }
QComboBox::drop-down { subcontrol-origin: padding; subcontrol-position: top right;
    width: 22px; border-left: 1px solid #dee2e6; border-top-right-radius: 4px;
    border-bottom-right-radius: 4px; background: #f4f6f8; }
QComboBox QAbstractItemView { background: #ffffff; border: 1px solid #ced4da;
    selection-background-color: #d6eaf8; selection-color: #1a5276; outline: none; padding: 2px; }
QPushButton {
    background: #ffffff; color: #2c3e50; border: 1.5px solid #aab4be;
    border-radius: 4px; padding: 5px 18px; font-weight: bold; min-height: 28px;
}
QPushButton:hover  { background: #eaf4fb; border-color: #3498db; color: #1a5276; }
QPushButton:pressed { background: #d6eaf8; border-color: #2980b9; }
QPushButton:disabled { background: #f4f6f8; color: #aab4be; border-color: #dee2e6; }
QPushButton#primary { background: #3498db; color: #ffffff; border: none; min-height: 28px; }
QPushButton#primary:hover   { background: #2980b9; }
QPushButton#primary:pressed { background: #1f618d; }
QPushButton#primary:disabled { background: #85c1e9; color: #eaf4fb; }
QPushButton#success { background: #27ae60; color: #ffffff; border: none; }
QPushButton#success:hover   { background: #219a52; }
QPushButton#success:pressed { background: #1a7a41; }
QPushButton#success:disabled { background: #82c09a; color: #e8f5e9; }
QPushButton#danger  { background: #e74c3c; color: #ffffff; border: none; }
QPushButton#danger:hover { background: #c0392b; }
QGroupBox {
    background: #ffffff; border: 1px solid #dee2e6; border-radius: 6px;
    margin-top: 14px; padding: 6px 10px 8px 10px;
}
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left;
    left: 12px; padding: 0 6px; color: #2980b9; font-weight: bold;
    font-size: 10pt; background: #ffffff; }
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical { background: #ecf0f1; width: 8px; border-radius: 4px; }
QScrollBar::handle:vertical { background: #aab4be; border-radius: 4px; min-height: 24px; }
QScrollBar::handle:vertical:hover { background: #7f8c8d; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QProgressBar { border: 1px solid #ced4da; border-radius: 4px; background: #ecf0f1;
    text-align: center; height: 14px; }
QProgressBar::chunk { background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
    stop:0 #3498db, stop:1 #5dade2); border-radius: 3px; }
QCheckBox { spacing: 8px; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1.5px solid #aab4be;
    border-radius: 3px; background: white; }
QCheckBox::indicator:checked { background: #27ae60; border-color: #219a52; }
QCheckBox::indicator:hover { border-color: #3498db; }
QFrame[frameShape="4"] { color: #dee2e6; }
QSplitter::handle { background: #dee2e6; width: 3px; }
QSplitter::handle:hover { background: #3498db; }
QTableWidget { border: 1px solid #dee2e6; background: #ffffff;
    gridline-color: #dee2e6; alternate-background-color: #f8f9fa; }
QTableWidget::item { padding: 4px; }
QTableWidget::item:selected { background: #d6eaf8; color: #1a5276; }
QHeaderView::section { background: #2F5496; color: white; font-weight: bold;
    padding: 6px 4px; border: none; border-right: 1px solid #dee2e6; }
#central   { background: #f8f9fa; }
#topBar { background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
    stop:0 #2c3e50, stop:1 #34495e); border: none; }
#pageLoading { background: #f8f9fa; }
#pageSearch  { background: #f8f9fa; }
#pageEditor  { background: #f8f9fa; }
#pageDetail  { background: #f8f9fa; }
#pageSuggest { background: #f8f9fa; }
#pageSupplement { background: #f8f9fa; }
"""


# ─────────────────────────────────────────
# 背景執行緒
# ─────────────────────────────────────────

class InitThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, rag: WizardRAG, force_rebuild: bool = False):
        super().__init__()
        self.rag = rag
        self.force_rebuild = force_rebuild

    def cancel(self) -> None:
        self.rag.stop()

    def run(self):
        try:
            if self.force_rebuild:
                self.rag.invalidate_cache()
            self.rag.initialize(progress_cb=lambda msg: self.progress.emit(msg))
            self._check_model()
            self.finished.emit(True, "")
        except Exception as e:
            self.finished.emit(False, str(e))

    def _check_model(self):
        try:
            from ai_chat import TAIDE_MODEL_PATH
            if not Path(TAIDE_MODEL_PATH).exists():
                self.progress.emit(
                    "⚠️ 找不到 TAIDE 模型，AI 行為指標生成將無法使用。"
                    f"請確認路徑：{TAIDE_MODEL_PATH}"
                )
        except Exception:
            pass


class SearchThread(QThread):
    """在背景執行緒執行 RAG 搜尋職能基準。"""
    finished = pyqtSignal(list)
    error    = pyqtSignal(str)

    def __init__(self, rag: WizardRAG, query: str):
        super().__init__()
        self.rag   = rag
        self.query = query

    def run(self):
        try:
            results = self.rag.search(self.query, top_k=3)
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))



# ─────────────────────────────────────────
# 逐任務 LLM 狀態
# ─────────────────────────────────────────

class TaskLLMState(Enum):
    IDLE    = auto()   # 未提交（描述為空或從未提交）
    PENDING = auto()   # 已提交，等待結果
    DONE    = auto()   # 結果已回來
    STALE   = auto()   # 描述已修改，需重新提交


def _task_hash(row: dict) -> str:
    """計算影響 LLM 輸出的欄位 MD5，用於偵測描述是否變更。"""
    key = (
        row.get("task_name", "") +
        row.get("user_description", "") +
        row.get("user_output", "") +
        row.get("template", "AUTO") +
        str(row.get("level", 3))
    )
    return hashlib.md5(key.encode()).hexdigest()


class PersistentLLMWorker(QThread):
    """
    包裝長駐 LLM 子 process（create_persistent_worker）。
    透過 submit() 送任務，run() 持續輪詢 result_q 並 emit task_done。
    """
    task_done    = pyqtSignal(int, list, str, str)   # idx, indicators, template_used, task_hash
    worker_ready = pyqtSignal()
    error        = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._proc     = None
        self._input_q  = None
        self._result_q = None
        self._stop     = False

    def start_worker(self):
        self._proc, self._input_q, self._result_q = create_persistent_worker()
        self.start()

    def submit(self, idx: int, task_args: dict, task_hash: str):
        if self._input_q is not None:
            self._input_q.put((idx, task_args, task_hash))

    def stop_worker(self):
        self._stop = True
        if self._input_q is not None:
            try:
                self._input_q.put(None)
            except Exception:
                pass
        if self._proc is not None:
            try:
                self._proc.kill()
            except Exception:
                pass

    def run(self):
        while not self._stop:
            if self._proc is None or not self._proc.is_alive():
                if not self._stop:
                    self.error.emit("AI 子程序意外中止")
                break
            try:
                item = self._result_q.get(timeout=1)
            except Exception:
                continue
            if item is None:
                break
            msg_type = item.get("type", "")
            if msg_type == "ready":
                self.worker_ready.emit()
            elif msg_type == "result":
                self.task_done.emit(
                    item.get("idx", -1),
                    item.get("indicators", []),
                    item.get("template_used", ""),
                    item.get("task_hash", ""),
                )


class ParseThread(QThread):
    progress = pyqtSignal(str)
    done     = pyqtSignal(int, int)

    def __init__(self, pdf_paths: list, json_dir: Path):
        super().__init__()
        self.pdf_paths  = pdf_paths
        self.json_dir   = json_dir
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def run(self):
        try:
            from pdf_parser_v2 import parse_pdf_to_json
        except ImportError:
            self.progress.emit("✗ pdfplumber 未安裝，請執行：pip install pdfplumber")
            self.done.emit(0, len(self.pdf_paths))
            return
        ok = err = 0
        for path_str in self.pdf_paths:
            if self._cancelled:
                self.progress.emit("⚠ 使用者已取消解析")
                break
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
# 資料管理對話框（保留原有功能）
# ─────────────────────────────────────────

class DataManagerDialog(QDialog):
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

    def _build_ui(self):
        v = QVBoxLayout(self)
        v.setSpacing(10)
        v.setContentsMargins(14, 12, 14, 12)
        v.addWidget(QLabel("raw_pdf 資料夾中的 PDF（勾選要操作的項目）："))
        self._search = QLineEdit()
        self._search.setPlaceholderText("搜尋 PDF 名稱...")
        self._search.setClearButtonEnabled(True)
        self._search.textChanged.connect(self._on_search)
        v.addWidget(self._search)
        self._list = QListWidget()
        v.addWidget(self._list, 1)
        row1 = QHBoxLayout()
        btn_add  = QPushButton("新增 PDF")
        btn_add.clicked.connect(self._on_add)
        btn_del  = QPushButton("刪除選取")
        btn_del.setObjectName("danger")
        btn_del.clicked.connect(self._on_delete)
        btn_all  = QPushButton("全選")
        btn_all.clicked.connect(self._check_all)
        btn_none = QPushButton("全不選")
        btn_none.clicked.connect(self._check_none)
        row1.addWidget(btn_add); row1.addWidget(btn_del)
        row1.addStretch()
        row1.addWidget(btn_all); row1.addWidget(btn_none)
        v.addLayout(row1)
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine); v.addWidget(sep)
        v.addWidget(QLabel("操作記錄："))
        self._log = QTextEdit()
        self._log.setReadOnly(True)
        self._log.setFixedHeight(130)
        self._log.setFont(QFont("Consolas", 9))
        v.addWidget(self._log)
        row2 = QHBoxLayout()
        self._btn_parse = QPushButton("解析勾選的 PDF → JSON")
        self._btn_parse.setObjectName("primary")
        self._btn_parse.clicked.connect(self._on_parse)
        self._btn_cancel_parse = QPushButton("取消解析")
        self._btn_cancel_parse.setObjectName("danger")
        self._btn_cancel_parse.setVisible(False)
        self._btn_cancel_parse.clicked.connect(self._on_parse_cancel)
        self._btn_rebuild = QPushButton("重建向量索引")
        self._btn_rebuild.setObjectName("success")
        self._btn_rebuild.clicked.connect(self._on_rebuild)
        btn_close = QPushButton("關閉")
        btn_close.clicked.connect(self.close)
        row2.addWidget(self._btn_parse); row2.addWidget(self._btn_cancel_parse)
        row2.addWidget(self._btn_rebuild); row2.addStretch(); row2.addWidget(btn_close)
        v.addLayout(row2)

    def _refresh_list(self):
        self._list.clear()
        self._raw_dir.mkdir(parents=True, exist_ok=True)
        pdfs = sorted(self._raw_dir.glob("*.pdf"))
        if not pdfs:
            item = QListWidgetItem("（資料夾中目前沒有 PDF）")
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEnabled)
            self._list.addItem(item); return
        for pdf in pdfs:
            parsed = (self._json_dir / (pdf.stem + ".json")).exists()
            label  = f"{'✓' if parsed else '✗'}  {pdf.name}"
            item   = QListWidgetItem(label)
            item.setCheckState(Qt.CheckState.Unchecked)
            item.setData(Qt.ItemDataRole.UserRole, str(pdf))
            if not parsed:
                item.setForeground(QColor("#e74c3c"))
            self._list.addItem(item)

    def _checked_paths(self) -> list:
        result = []
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item is None:
                continue
            if item.checkState() == Qt.CheckState.Checked:
                p = item.data(Qt.ItemDataRole.UserRole)
                if p:
                    result.append(p)
        return result

    def _check_all(self):
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item is None:
                continue
            if item.data(Qt.ItemDataRole.UserRole):
                item.setCheckState(Qt.CheckState.Checked)

    def _check_none(self):
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item is not None:
                item.setCheckState(Qt.CheckState.Unchecked)

    def _on_search(self, text: str):
        kw = text.strip().lower()
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item is None:
                continue
            item.setHidden(bool(kw) and kw not in item.text().lower())

    def _on_add(self):
        self._raw_dir.mkdir(parents=True, exist_ok=True)
        paths, _ = QFileDialog.getOpenFileNames(
            self, "選擇 PDF 檔案", str(Path.home()), "PDF 檔案 (*.pdf)")
        if not paths: return
        copied = 0
        for src in paths:
            dst = self._raw_dir / Path(src).name
            if dst.exists():
                self._log.append(f"⚠ 已存在，略過：{Path(src).name}")
            else:
                shutil.copy2(src, dst)
                self._log.append(f"✓ 已複製：{Path(src).name}")
                copied += 1
        if copied: self._refresh_list()

    def _on_delete(self):
        paths = self._checked_paths()
        if not paths:
            QMessageBox.information(self, "提示", "請先勾選要刪除的 PDF"); return
        names = "\n".join(Path(p).name for p in paths)
        reply = QMessageBox.question(
            self, "確認刪除",
            f"確定要刪除以下 {len(paths)} 個 PDF 及其對應 JSON？\n\n{names}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes: return
        for p in paths:
            pdf = Path(p); pdf.unlink(missing_ok=True)
            self._log.append(f"🗑 已刪除 PDF：{pdf.name}")
            json_f = self._json_dir / (pdf.stem + ".json")
            if json_f.exists():
                json_f.unlink()
                self._log.append(f"🗑 已刪除 JSON：{json_f.name}")
        self._refresh_list()

    def _on_parse(self):
        paths = self._checked_paths()
        if not paths:
            QMessageBox.information(self, "提示", "請先勾選要解析的 PDF"); return
        self._btn_parse.setEnabled(False); self._btn_rebuild.setEnabled(False)
        self._btn_cancel_parse.setVisible(True)
        self._log.append(f"\n▶ 開始解析 {len(paths)} 個 PDF...")
        self._parse_thread = ParseThread(paths, self._json_dir)
        self._parse_thread.progress.connect(self._log.append)
        self._parse_thread.done.connect(self._on_parse_done)
        self._parse_thread.start()

    def _on_parse_done(self, ok: int, err: int):
        self._log.append(f"── 完成：{ok} 成功，{err} 失敗 ──")
        self._btn_parse.setEnabled(True); self._btn_rebuild.setEnabled(True)
        self._btn_cancel_parse.setVisible(False)
        self._refresh_list()
        # 有失敗時主動彈出警告，避免使用者遺漏 log 訊息（P10）
        if err > 0:
            QMessageBox.warning(
                self, "PDF 解析部分失敗",
                f"共 {ok + err} 個檔案，{err} 個解析失敗。\n"
                "請查看下方日誌確認失敗原因，\n"
                "問題排除後可重新選取該 PDF 再次解析。"
            )

    def _on_parse_cancel(self):
        if self._parse_thread and self._parse_thread.isRunning():
            self._parse_thread.cancel()
        self._btn_cancel_parse.setVisible(False)

    def _on_rebuild(self):
        reply = QMessageBox.question(
            self, "重建向量索引",
            "確定要重建向量索引？\n（需要數分鐘，完成後程式將回到載入畫面）",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes: return
        self._log.append("\n▶ 送出重建請求...")
        self.rebuild_requested.emit()
        self.close()


# ─────────────────────────────────────────
# 輔助函式：從職能基準資料轉換為 table row list
# ─────────────────────────────────────────

def _rows_from_standard(std_data: dict) -> List[Dict]:
    """將職能基準 JSON 展開為每任務一列的 row list。"""
    import re as _re
    rows = []
    tasks = std_data.get("competency_tasks") or []

    # 建立 code→name 對照表
    k_map = {k["code"]: k["name"] for k in std_data.get("competency_knowledge", []) if "code" in k}
    s_map = {s["code"]: s["name"] for s in std_data.get("competency_skills", []) if "code" in s}

    for task in tasks:
        task_id   = task.get("task_id", "")
        resp_code = task_id.split(".")[0] if "." in task_id else task_id

        # 主責名稱去除開頭的 T-code 前綴（如 "T1製作與..." → "製作與..."）
        resp_raw  = task.get("main_responsibility", "")
        resp_name = _re.sub(r'^T\d+', '', resp_raw).strip()

        outputs = task.get("output") or []
        if isinstance(outputs, list) and outputs:
            out_str = "；".join(
                o.get("name", "") if isinstance(o, dict) else str(o)
                for o in outputs[:3]
            )
        elif isinstance(outputs, str):
            out_str = outputs
        else:
            out_str = ""

        # 知識/技能展開為 {code, name} dict，讓 exporter 能輸出名稱
        knowledge = [
            {"code": code, "name": k_map.get(code, "")}
            for code in (task.get("knowledge") or [])
        ]
        skills = [
            {"code": code, "name": s_map.get(code, "")}
            for code in (task.get("skills") or [])
        ]

        rows.append({
            "resp_code":        resp_code,
            "resp_name":        resp_name,
            "task_code":        task_id,
            "task_name":        task.get("task_name", ""),
            "output":           out_str,
            "level":            task.get("level", 3),
            # 隱藏欄：供 LLM 使用
            "_behaviors":       task.get("behaviors") or [],
            "_knowledge":       knowledge,
            "_skills":          skills,
            # Step 3 填入
            "user_description": "",
            "user_output":      "",
            # Step 4 LLM 生成
            "behavior_accepted": [],
        })
    return rows


# ─────────────────────────────────────────
# 主視窗
# ─────────────────────────────────────────

class WizardMainWindow(QMainWindow):
    """
    職能說明書精靈主視窗。

    Stack pages:
      0 — 載入頁
      1 — 搜索頁（填職業名稱）
      2 — 編輯器頁（職能基準書 Table）
      3 — 詳細填寫頁（逐任務）
      4 — LLM 建議確認頁
      5 — 補充說明 & 匯出頁
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("職能說明書精靈 v2.0")
        self.setMinimumSize(900, 640)
        self.resize(1100, 740)

        self._rag: WizardRAG = WizardRAG()
        self._init_thread: Optional[InitThread] = None
        self._search_thread: Optional[SearchThread] = None
        self._llm_worker: Optional[PersistentLLMWorker] = None
        self._init_timer: Optional[QTimer] = None    # T1：初始化逾時計時器
        self._search_timer: Optional[QTimer] = None  # T3：搜尋逾時計時器
        self._analysis_template: str = "AUTO"

        # 跨頁資料
        self._position: str = ""
        self._level:    int = 3
        self._search_results: List[Dict] = []   # RAG 候選清單
        self._matched_std:    Optional[Dict] = None
        self._last_std_code:  str = ""          # 上次載入編輯器使用的基準代碼（P2）
        self._std_radio_group: List[QCheckBox] = []  # 搜尋結果選擇器（P1）
        self._competency_rows: List[Dict] = []  # 主要資料
        self._current_task_idx: int = 0
        self._suggest_checks: list[SuggestEntry | None] = []

        # 逐任務 LLM 狀態追蹤
        self._task_states: List[TaskLLMState] = []
        self._task_hashes: List[str] = []
        self._suggest_boxes: List = []  # QGroupBox per task

        self._build_ui()
        self._start_init()

    # ─────────────────────────────────────
    # UI 建立
    # ─────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self._make_top_bar())
        self.stack = QStackedWidget()
        layout.addWidget(self.stack, 1)
        self.stack.addWidget(self._make_loading_page())   # 0
        self.stack.addWidget(self._make_search_page())    # 1
        self.stack.addWidget(self._make_editor_page())    # 2
        self.stack.addWidget(self._make_detail_page())    # 3
        self.stack.addWidget(self._make_suggest_page())   # 4
        self.stack.addWidget(self._make_supplement_page())# 5

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
        )
        btn_data.clicked.connect(self._open_data_manager)
        h.addWidget(btn_data)
        self._status_label = QLabel("初始化中...")
        self._status_label.setStyleSheet(
            "color:#aed6f1; font-size:9pt; background:rgba(255,255,255,0.10); "
            "border-radius:4px; padding:2px 10px;")
        h.addWidget(self._status_label)
        return bar

    # ── Page 0: 載入頁 ──────────────────────────────────────────────────────

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
        self._loading_bar = QProgressBar()
        self._loading_bar.setRange(0, 0)
        self._loading_bar.setFixedWidth(320)
        v.addWidget(self._loading_bar, 0, Qt.AlignmentFlag.AlignCenter)
        self._btn_force_rebuild = QPushButton("強制重建索引")
        self._btn_force_rebuild.setFixedWidth(160)
        self._btn_force_rebuild.clicked.connect(lambda: self._start_init(force=True))
        v.addWidget(self._btn_force_rebuild, 0, Qt.AlignmentFlag.AlignCenter)

        self._btn_cancel_init = QPushButton("取消初始化")
        self._btn_cancel_init.setFixedWidth(160)
        self._btn_cancel_init.setObjectName("danger")
        self._btn_cancel_init.setVisible(False)
        self._btn_cancel_init.clicked.connect(self._on_cancel_init)
        v.addWidget(self._btn_cancel_init, 0, Qt.AlignmentFlag.AlignCenter)
        return w

    # ── Page 1: 搜索頁 ──────────────────────────────────────────────────────

    def _make_search_page(self) -> QWidget:
        w = QWidget()
        w.setObjectName("pageSearch")
        outer = QVBoxLayout(w)
        outer.setContentsMargins(40, 30, 40, 30)
        outer.setSpacing(20)

        # 標題
        title = QLabel("填寫職業名稱")
        title.setFont(QFont("Microsoft JhengHei", 15, QFont.Weight.Bold))
        title.setStyleSheet("color:#2c3e50;")
        outer.addWidget(title)

        subtitle = QLabel("請輸入您的職業名稱，系統將自動搜尋最相近的 ICAP 職能基準。")
        subtitle.setStyleSheet("color:#7f8c8d;")
        outer.addWidget(subtitle)

        # 輸入列
        input_row = QHBoxLayout()
        self._search_input = QLineEdit()
        self._search_input.setPlaceholderText("例：糕點師傅、IT 維運工程師、護理人員...")
        self._search_input.setFont(QFont("Microsoft JhengHei", 12))
        self._search_input.setFixedHeight(40)
        self._search_input.returnPressed.connect(self._on_search)
        input_row.addWidget(self._search_input, 1)
        self._btn_search = QPushButton("搜尋職能基準")
        self._btn_search.setObjectName("primary")
        self._btn_search.setFixedHeight(40)
        self._btn_search.setMinimumWidth(130)
        self._btn_search.clicked.connect(self._on_search)
        input_row.addWidget(self._btn_search)
        outer.addLayout(input_row)

        # 搜尋結果區
        result_box = QGroupBox("搜尋結果（請選擇最符合的職能基準）")
        result_v = QVBoxLayout(result_box)

        self._search_result_label = QLabel("尚未搜尋")
        self._search_result_label.setStyleSheet("color:#7f8c8d; font-style:italic;")
        result_v.addWidget(self._search_result_label)

        self._result_scroll = QScrollArea()
        self._result_scroll.setWidgetResizable(True)
        self._result_scroll.setFixedHeight(240)
        self._result_content = QWidget()
        self._result_layout  = QVBoxLayout(self._result_content)
        self._result_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self._result_scroll.setWidget(self._result_content)
        result_v.addWidget(self._result_scroll)
        outer.addWidget(result_box, 1)

        # 底部按鈕
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self._btn_goto_editor = QPushButton("下一步：確認職能基準書  →")
        self._btn_goto_editor.setObjectName("success")
        self._btn_goto_editor.setMinimumWidth(220)
        self._btn_goto_editor.setEnabled(False)
        self._btn_goto_editor.clicked.connect(self._goto_editor)
        btn_row.addWidget(self._btn_goto_editor)
        outer.addLayout(btn_row)
        return w

    # ── Page 2: 編輯器頁 ─────────────────────────────────────────────────────

    def _make_editor_page(self) -> QWidget:
        w = QWidget()
        w.setObjectName("pageEditor")
        v = QVBoxLayout(w)
        v.setContentsMargins(24, 16, 24, 16)
        v.setSpacing(10)

        # 標題列
        title_row = QHBoxLayout()
        title = QLabel("職能基準書編輯")
        title.setFont(QFont("Microsoft JhengHei", 13, QFont.Weight.Bold))
        title.setStyleSheet("color:#2c3e50;")
        title_row.addWidget(title)
        title_row.addStretch()

        level_label = QLabel("職能等級：")
        level_label.setStyleSheet("color:#4a5568;")
        title_row.addWidget(level_label)
        self._level_spin = QSpinBox()
        self._level_spin.setRange(1, 5)
        self._level_spin.setValue(3)
        self._level_spin.setFixedWidth(60)
        title_row.addWidget(self._level_spin)
        v.addLayout(title_row)

        subtitle = QLabel("可新增、刪除或直接點選格子修改內容。完成後按「下一步」填寫工作詳情。")
        subtitle.setStyleSheet("color:#7f8c8d; font-size:9pt;")
        v.addWidget(subtitle)

        # 主要 Table
        self._editor_table = QTableWidget()
        self._editor_table.setColumnCount(6)
        self._editor_table.setHorizontalHeaderLabels(
            ["主責代碼", "主責名稱", "任務代碼", "任務名稱", "工作產出", "等級"])
        hh = self._editor_table.horizontalHeader()
        if hh is not None:
            hh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            hh.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
            hh.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            hh.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
            hh.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
            hh.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self._editor_table.setAlternatingRowColors(True)
        self._editor_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        v.addWidget(self._editor_table, 1)

        # Table 操作列
        tbl_row = QHBoxLayout()
        btn_add_row = QPushButton("＋ 新增任務列")
        btn_add_row.clicked.connect(self._table_add_row)
        btn_del_row = QPushButton("− 刪除選取列")
        btn_del_row.setObjectName("danger")
        btn_del_row.clicked.connect(self._table_del_row)
        tbl_row.addWidget(btn_add_row)
        tbl_row.addWidget(btn_del_row)
        tbl_row.addStretch()
        v.addLayout(tbl_row)

        # 導航列
        nav_row = QHBoxLayout()
        btn_back = QPushButton("← 返回搜尋")
        btn_back.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        nav_row.addWidget(btn_back)
        nav_row.addStretch()
        self._btn_goto_detail = QPushButton("下一步：填寫工作詳情  →")
        self._btn_goto_detail.setObjectName("primary")
        self._btn_goto_detail.setMinimumWidth(200)
        self._btn_goto_detail.clicked.connect(self._goto_detail)
        nav_row.addWidget(self._btn_goto_detail)
        v.addLayout(nav_row)
        return w

    # ── Page 3: 詳細填寫頁 ──────────────────────────────────────────────────

    def _make_detail_page(self) -> QWidget:
        w = QWidget()
        w.setObjectName("pageDetail")
        v = QVBoxLayout(w)
        v.setContentsMargins(40, 20, 40, 20)
        v.setSpacing(14)

        # 進度標題
        progress_row = QHBoxLayout()
        title = QLabel("填寫工作詳情")
        title.setFont(QFont("Microsoft JhengHei", 13, QFont.Weight.Bold))
        title.setStyleSheet("color:#2c3e50;")
        progress_row.addWidget(title)
        progress_row.addStretch()
        self._detail_llm_badge = QLabel()
        self._detail_llm_badge.setStyleSheet("font-size:9pt; font-weight:bold;")
        progress_row.addWidget(self._detail_llm_badge)
        self._detail_progress_label = QLabel("任務 1 / 1")
        self._detail_progress_label.setStyleSheet(
            "color:#ffffff; background:#3498db; border-radius:4px; padding:3px 12px; font-weight:bold;")
        progress_row.addWidget(self._detail_progress_label)
        v.addLayout(progress_row)

        self._detail_progress_bar = QProgressBar()
        self._detail_progress_bar.setFixedHeight(8)
        v.addWidget(self._detail_progress_bar)

        # 後台分析進度提示
        self._detail_bg_status = QLabel()
        self._detail_bg_status.setStyleSheet(
            "color:#7f8c8d; font-size:8.5pt; padding:1px 0;")
        v.addWidget(self._detail_bg_status)

        # 任務資訊卡
        self._detail_task_card = QGroupBox("當前任務")
        card_v = QVBoxLayout(self._detail_task_card)
        self._detail_task_code_lbl = QLabel("")
        self._detail_task_code_lbl.setStyleSheet("color:#2980b9; font-size:9pt;")
        self._detail_task_name_lbl = QLabel("")
        self._detail_task_name_lbl.setFont(QFont("Microsoft JhengHei", 11, QFont.Weight.Bold))
        card_v.addWidget(self._detail_task_code_lbl)
        card_v.addWidget(self._detail_task_name_lbl)
        v.addWidget(self._detail_task_card)

        # 描述欄位
        lbl1 = QLabel("請描述您實際如何執行此工作任務：")
        lbl1.setStyleSheet("font-weight:bold; color:#2c3e50;")
        v.addWidget(lbl1)
        self._detail_desc = QTextEdit()
        self._detail_desc.setPlaceholderText(
            "例：我負責每週一次清點倉庫庫存，使用 ERP 系統登記盤點結果，並在出入量異常時通知主管...")
        self._detail_desc.setFixedHeight(110)
        v.addWidget(self._detail_desc)

        _DESC_LIMIT = 800
        _OUTPUT_LIMIT = 400
        self._detail_desc_count = QLabel(f"0 / {_DESC_LIMIT}")
        self._detail_desc_count.setAlignment(Qt.AlignmentFlag.AlignRight)
        self._detail_desc_count.setStyleSheet("color:#888; font-size:8pt;")
        v.addWidget(self._detail_desc_count)

        def _enforce_desc_limit():
            txt = self._detail_desc.toPlainText()
            if len(txt) > _DESC_LIMIT:
                cur = self._detail_desc.textCursor()
                pos = min(cur.position(), _DESC_LIMIT)
                self._detail_desc.blockSignals(True)
                self._detail_desc.setPlainText(txt[:_DESC_LIMIT])
                self._detail_desc.blockSignals(False)
                cur.setPosition(pos)
                self._detail_desc.setTextCursor(cur)
            self._detail_desc_count.setText(f"{len(self._detail_desc.toPlainText())} / {_DESC_LIMIT}")

        self._detail_desc.textChanged.connect(_enforce_desc_limit)

        lbl2 = QLabel("此任務的主要工作成果或產出：")
        lbl2.setStyleSheet("font-weight:bold; color:#2c3e50;")
        v.addWidget(lbl2)
        self._detail_output = QTextEdit()
        self._detail_output.setPlaceholderText("例：每週庫存盤點報告、異常差異通報紀錄...")
        self._detail_output.setFixedHeight(80)
        v.addWidget(self._detail_output)

        self._detail_output_count = QLabel(f"0 / {_OUTPUT_LIMIT}")
        self._detail_output_count.setAlignment(Qt.AlignmentFlag.AlignRight)
        self._detail_output_count.setStyleSheet("color:#888; font-size:8pt;")
        v.addWidget(self._detail_output_count)

        def _enforce_output_limit():
            txt = self._detail_output.toPlainText()
            if len(txt) > _OUTPUT_LIMIT:
                cur = self._detail_output.textCursor()
                pos = min(cur.position(), _OUTPUT_LIMIT)
                self._detail_output.blockSignals(True)
                self._detail_output.setPlainText(txt[:_OUTPUT_LIMIT])
                self._detail_output.blockSignals(False)
                cur.setPosition(pos)
                self._detail_output.setTextCursor(cur)
            self._detail_output_count.setText(f"{len(self._detail_output.toPlainText())} / {_OUTPUT_LIMIT}")

        self._detail_output.textChanged.connect(_enforce_output_limit)

        # AI 分析框架選擇
        tpl_row = QHBoxLayout()
        tpl_lbl = QLabel("AI 分析框架：")
        tpl_lbl.setStyleSheet("font-weight:bold; color:#2c3e50;")
        tpl_row.addWidget(tpl_lbl)
        self._detail_tpl_combo = QComboBox()
        for _key, _label in [
            ("AUTO", "自動｜AI 依任務性質選擇（推薦）"),
            ("ABCD", "ABCD｜條件＋行動＋標準"),
            ("5W2H", "5W2H｜操作步驟＋頻率＋標準"),
            ("STAR", "STAR｜情境＋行動＋成果"),
        ]:
            self._detail_tpl_combo.addItem(_label, _key)
        self._detail_tpl_combo.setFixedHeight(30)
        tpl_row.addWidget(self._detail_tpl_combo, 1)
        v.addLayout(tpl_row)

        v.addStretch()

        # 導航列
        nav_row = QHBoxLayout()
        self._btn_detail_prev = QPushButton("← 上一個任務")
        self._btn_detail_prev.clicked.connect(self._detail_prev)
        nav_row.addWidget(self._btn_detail_prev)
        nav_row.addStretch()
        self._btn_detail_next = QPushButton("下一個任務  →")
        self._btn_detail_next.setObjectName("primary")
        self._btn_detail_next.setMinimumWidth(180)
        self._btn_detail_next.clicked.connect(self._detail_next)
        nav_row.addWidget(self._btn_detail_next)
        v.addLayout(nav_row)
        return w

    # ── Page 4: LLM 建議確認頁 ──────────────────────────────────────────────

    def _make_suggest_page(self) -> QWidget:
        w = QWidget()
        w.setObjectName("pageSuggest")
        v = QVBoxLayout(w)
        v.setContentsMargins(24, 16, 24, 16)
        v.setSpacing(10)

        title_row = QHBoxLayout()
        title = QLabel("AI 行為指標建議")
        title.setFont(QFont("Microsoft JhengHei", 13, QFont.Weight.Bold))
        title.setStyleSheet("color:#2c3e50;")
        title_row.addWidget(title)
        title_row.addStretch()
        self._suggest_status_lbl = QLabel("準備中...")
        self._suggest_status_lbl.setStyleSheet(
            "color:#ffffff; background:#27ae60; border-radius:4px; padding:3px 10px; font-size:9pt;")
        title_row.addWidget(self._suggest_status_lbl)
        v.addLayout(title_row)

        subtitle = QLabel("請勾選要採用的行為指標，也可以直接在文字方塊中修改。")
        subtitle.setStyleSheet("color:#7f8c8d; font-size:9pt;")
        v.addWidget(subtitle)

        self._tpl_label = QLabel()
        self._tpl_label.setStyleSheet(
            "color:#1a3a5c; background:#d9e1f2; border-radius:4px;"
            " padding:2px 10px; font-size:9pt;")
        v.addWidget(self._tpl_label)

        self._suggest_progress = QProgressBar()
        self._suggest_progress.setFixedHeight(8)
        v.addWidget(self._suggest_progress)

        # 建議內容捲動區
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self._suggest_content = QWidget()
        self._suggest_layout  = QVBoxLayout(self._suggest_content)
        self._suggest_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self._suggest_layout.setSpacing(8)
        scroll.setWidget(self._suggest_content)
        v.addWidget(scroll, 1)

        # 導航列
        nav_row = QHBoxLayout()
        btn_back = QPushButton("← 修改工作詳情")
        btn_back.clicked.connect(lambda: self.stack.setCurrentIndex(3))
        nav_row.addWidget(btn_back)
        btn_rerun = QPushButton("重新 AI 分析")
        btn_rerun.clicked.connect(self._rerun_llm)
        btn_rerun.setToolTip("使用「填寫工作詳情」頁儲存的描述重新呼叫 LLM\n若需修改描述請先按「← 修改工作詳情」")
        nav_row.addWidget(btn_rerun)
        self._rerun_hint = QLabel("（依工作詳情頁的描述重新生成）")
        self._rerun_hint.setStyleSheet("color:#718096; font-size:8pt;")
        nav_row.addWidget(self._rerun_hint)
        nav_row.addStretch()
        self._btn_confirm_suggest = QPushButton("確認採用  →")
        self._btn_confirm_suggest.setObjectName("success")
        self._btn_confirm_suggest.setMinimumWidth(160)
        self._btn_confirm_suggest.setEnabled(False)
        self._btn_confirm_suggest.clicked.connect(self._goto_supplement)
        nav_row.addWidget(self._btn_confirm_suggest)
        v.addLayout(nav_row)
        return w

    # ── Page 5: 補充說明 & 匯出頁 ───────────────────────────────────────────

    def _make_supplement_page(self) -> QWidget:
        w = QWidget()
        w.setObjectName("pageSupplement")
        v = QVBoxLayout(w)
        v.setContentsMargins(40, 30, 40, 30)
        v.setSpacing(16)

        title = QLabel("填寫說明與補充事項")
        title.setFont(QFont("Microsoft JhengHei", 13, QFont.Weight.Bold))
        title.setStyleSheet("color:#2c3e50;")
        v.addWidget(title)

        # 補充說明
        v.addWidget(QLabel("說明與補充事項（選填）："))
        self._supplement_text = QTextEdit()
        self._supplement_text.setPlaceholderText(
            "可填寫特殊工作情境、資格說明、部門背景或其他備注...")
        self._supplement_text.setFixedHeight(160)
        v.addWidget(self._supplement_text)

        # 摘要預覽
        summary_box = QGroupBox("職能說明書摘要")
        summary_v = QVBoxLayout(summary_box)
        self._summary_label = QLabel("（完成前頁步驟後將顯示摘要）")
        self._summary_label.setWordWrap(True)
        self._summary_label.setStyleSheet("color:#4a5568; font-size:9pt; line-height:1.6;")
        summary_v.addWidget(self._summary_label)
        v.addWidget(summary_box, 1)

        # 導航列
        nav_row = QHBoxLayout()
        btn_back = QPushButton("← 返回")
        btn_back.clicked.connect(lambda: self.stack.setCurrentIndex(4))
        nav_row.addWidget(btn_back)
        nav_row.addStretch()
        self._btn_export = QPushButton("匯出 Excel 職能說明書")
        self._btn_export.setObjectName("success")
        self._btn_export.setMinimumWidth(220)
        self._btn_export.clicked.connect(self._on_export)
        nav_row.addWidget(self._btn_export)
        v.addLayout(nav_row)
        return w

    # ─────────────────────────────────────
    # 初始化（Page 0）
    # ─────────────────────────────────────

    _INIT_TIMEOUT_MS = 5 * 60 * 1000   # 5 分鐘

    def _start_init(self, force: bool = False):
        self.stack.setCurrentIndex(0)
        self._loading_label.setText("正在載入 Embedding 模型，請稍候...")
        self._loading_bar.setRange(0, 0)
        self._btn_force_rebuild.setEnabled(False)
        self._btn_cancel_init.setVisible(True)
        self._btn_cancel_init.setEnabled(True)

        self._init_thread = InitThread(self._rag, force_rebuild=force)
        self._init_thread.progress.connect(self._loading_label.setText)
        self._init_thread.finished.connect(self._on_init_done)
        self._init_thread.start()

        # T1：啟動逾時計時器，超過 5 分鐘自動取消
        if self._init_timer is not None:
            self._init_timer.stop()
        self._init_timer = QTimer(self)
        self._init_timer.setSingleShot(True)
        self._init_timer.timeout.connect(self._on_init_timeout)
        self._init_timer.start(self._INIT_TIMEOUT_MS)

    def _on_cancel_init(self):
        """使用者手動取消初始化。"""
        if self._init_timer is not None:
            self._init_timer.stop()
        if self._init_thread is not None:
            self._init_thread.cancel()
        self._btn_cancel_init.setEnabled(False)
        self._loading_label.setText("已取消初始化，請點擊「強制重建索引」重試。")
        self._loading_bar.setRange(0, 1)
        self._loading_bar.setValue(0)
        self._btn_force_rebuild.setEnabled(True)
        self._btn_cancel_init.setVisible(False)

    def _on_init_timeout(self):
        """初始化超過 5 分鐘，自動取消並提示使用者。"""
        if self._init_thread is not None:
            self._init_thread.cancel()
        self._loading_label.setText(
            "初始化逾時（超過 5 分鐘），可能是模型載入或索引建立過慢。\n"
            "請確認磁碟空間充足後，點擊「強制重建索引」重試。"
        )
        self._loading_bar.setRange(0, 1)
        self._loading_bar.setValue(0)
        self._btn_force_rebuild.setEnabled(True)
        self._btn_cancel_init.setVisible(False)

    def _on_init_done(self, ok: bool, err: str):
        if self._init_timer is not None:
            self._init_timer.stop()
        self._btn_cancel_init.setVisible(False)
        self._loading_bar.setRange(0, 1)
        self._loading_bar.setValue(1)
        self._btn_force_rebuild.setEnabled(True)
        if ok:
            self._status_label.setText(f"就緒 — {self._rag.chunk_count} 個向量")
            self.stack.setCurrentIndex(1)
        else:
            self._loading_label.setText(f"初始化失敗：{err}")

    # ─────────────────────────────────────
    # 搜尋（Page 1）
    # ─────────────────────────────────────

    def _on_search(self):
        query = self._search_input.text().strip()
        if not query:
            QMessageBox.information(self, "提示", "請輸入職業名稱後再搜尋")
            return
        self._position = query
        self._btn_search.setEnabled(False)
        self._btn_goto_editor.setEnabled(False)
        self._search_result_label.setText("搜尋中...")
        # 清空舊結果
        while self._result_layout.count():
            item = self._result_layout.takeAt(0)
            if item is None:
                continue
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self._search_results = []
        self._search_thread = SearchThread(self._rag, query)
        self._search_thread.finished.connect(self._on_search_done)
        self._search_thread.error.connect(self._on_search_error)
        self._search_thread.start()

        # T3：30 秒搜尋逾時，恢復按鈕並顯示錯誤（不強制終止執行緒）
        if self._search_timer is not None:
            self._search_timer.stop()
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._on_search_timeout)
        self._search_timer.start(30_000)

    def _on_search_timeout(self):
        """搜尋逾時：斷開舊執行緒信號，恢復 UI。"""
        if self._search_thread is not None:
            try:
                self._search_thread.finished.disconnect(self._on_search_done)
                self._search_thread.error.disconnect(self._on_search_error)
            except Exception:
                pass
        self._btn_search.setEnabled(True)
        self._search_result_label.setText(
            "搜尋逾時（超過 30 秒），請確認 Embedding 模型是否正常載入後重試。"
        )

    def _on_search_done(self, results: list):
        if self._search_timer is not None:
            self._search_timer.stop()
        self._btn_search.setEnabled(True)
        self._search_results = results

        if not results:
            self._search_result_label.setText("找不到相符的職能基準，請嘗試其他關鍵字。")
            return

        self._search_result_label.setText(f"找到 {len(results)} 個候選職能基準：")
        self._std_radio_group: List[QCheckBox] = []

        # 「不使用基準」選項
        no_std_rb = QCheckBox("不使用職能基準，從空白開始填寫")
        no_std_rb.setChecked(False)
        no_std_rb.stateChanged.connect(lambda: self._on_result_selected(-1, no_std_rb))
        self._result_layout.addWidget(no_std_rb)
        self._std_radio_group.append(no_std_rb)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        self._result_layout.addWidget(sep)

        for i, r in enumerate(results):
            score_pct = int(r.get("score", 0) * 100)
            name      = r.get("standard_name", "（未知）")
            code      = r.get("standard_code", "")
            category  = r.get("standard_category", "")
            preview   = r.get("matched_text", "")[:120].replace("\n", " ")

            cb = QCheckBox(f"[{code}] {name}  （相似度 {score_pct}%）")
            cb.setChecked(i == 0)
            cb.setStyleSheet("font-weight:bold; color:#1a5276;")
            cb.stateChanged.connect(lambda state, idx=i, _cb=cb: self._on_result_selected(idx, _cb))
            self._result_layout.addWidget(cb)

            if category:
                cat_lbl = QLabel(f"  🏷 {category}")
                cat_lbl.setStyleSheet(
                    "color:#fff; background:#2980b9; border-radius:3px;"
                    "font-size:8pt; padding:1px 6px; margin-left:24px;"
                )
                cat_lbl.setFixedHeight(18)
                self._result_layout.addWidget(cat_lbl)

            lbl = QLabel(f"  {preview}...")
            lbl.setStyleSheet("color:#666; font-size:9pt; padding-left:24px;")
            lbl.setWordWrap(True)
            self._result_layout.addWidget(lbl)

            self._std_radio_group.append(cb)

        # 預設選第一個
        self._on_result_selected(0, self._std_radio_group[1])
        self._btn_goto_editor.setEnabled(True)

    def _on_result_selected(self, idx: int, source_cb: QCheckBox):
        """Radio-group 行為：只保留一個 checked。"""
        for cb in self._std_radio_group:
            if cb is not source_cb:
                cb.setChecked(False)
        if idx >= 0 and idx < len(self._search_results):
            code = self._search_results[idx].get("standard_code", "")
            self._matched_std = self._rag.get_standard(code)
        else:
            self._matched_std = None

    def _on_search_error(self, msg: str):
        if self._search_timer is not None:
            self._search_timer.stop()
        self._btn_search.setEnabled(True)
        self._search_result_label.setText(f"搜尋失敗：{msg}")

    def _goto_editor(self):
        """從搜尋頁進入編輯器頁，預填或清空 Table。"""
        self._level = self._level_spin.value()
        new_code = (self._matched_std or {}).get("metadata", {}).get("code", "")
        # 只有基準改變或目前無任何任務時才重新載入，避免覆蓋使用者的編輯（P2）
        if not self._competency_rows or new_code != self._last_std_code:
            self._competency_rows = _rows_from_standard(self._matched_std) if self._matched_std else []
            self._last_std_code = new_code
        self._refresh_editor_table()
        self.stack.setCurrentIndex(2)

    # ─────────────────────────────────────
    # 編輯器（Page 2）
    # ─────────────────────────────────────

    def _refresh_editor_table(self):
        """將 self._competency_rows 寫入 QTableWidget。"""
        t = self._editor_table
        t.setRowCount(0)
        for row in self._competency_rows:
            r = t.rowCount()
            t.insertRow(r)
            t.setItem(r, 0, QTableWidgetItem(row.get("resp_code", "")))
            t.setItem(r, 1, QTableWidgetItem(row.get("resp_name", "")))
            t.setItem(r, 2, QTableWidgetItem(row.get("task_code", "")))
            t.setItem(r, 3, QTableWidgetItem(row.get("task_name", "")))
            t.setItem(r, 4, QTableWidgetItem(row.get("output", "")))
            t.setItem(r, 5, QTableWidgetItem(str(row.get("level", 3))))

    def _table_add_row(self):
        t = self._editor_table
        r = t.rowCount()
        t.insertRow(r)
        # 繼承上一列的主責代碼/名稱
        if r > 0:
            prev_resp_code = t.item(r - 1, 0)
            prev_resp_name = t.item(r - 1, 1)
            t.setItem(r, 0, QTableWidgetItem(prev_resp_code.text() if prev_resp_code else ""))
            t.setItem(r, 1, QTableWidgetItem(prev_resp_name.text() if prev_resp_name else ""))
        t.setItem(r, 5, QTableWidgetItem("3"))
        t.scrollToBottom()
        t.setCurrentCell(r, 2)

    def _table_del_row(self):
        rows = sorted({idx.row() for idx in self._editor_table.selectedIndexes()}, reverse=True)
        for r in rows:
            self._editor_table.removeRow(r)

    def _extract_rows_from_table(self) -> List[Dict]:
        """將 QTableWidget 的內容提取為 row dict list。"""
        t = self._editor_table
        rows = []
        for r in range(t.rowCount()):
            def cell(c):
                item = t.item(r, c)
                return (item.text() if item else "").strip()
            task_code = cell(2)
            if not task_code:
                continue
            # 找回原始 _behaviors/_knowledge/_skills（若存在）
            orig = next(
                (o for o in self._competency_rows if o.get("task_code") == task_code), {})
            rows.append({
                "resp_code":        cell(0),
                "resp_name":        cell(1),
                "task_code":        task_code,
                "task_name":        cell(3),
                "output":           cell(4),
                "level":            max(1, min(5, int(cell(5)))) if cell(5).isdigit() else 3,
                "_behaviors":       orig.get("_behaviors", []),
                "_knowledge":       orig.get("_knowledge", []),
                "_skills":          orig.get("_skills", []),
                "user_description": orig.get("user_description", ""),
                "user_output":      orig.get("user_output", ""),
                "behavior_accepted": [],
            })
        return rows

    def _goto_detail(self):
        rows = self._extract_rows_from_table()
        if not rows:
            QMessageBox.information(self, "提示", "請至少填寫一列工作任務（任務代碼欄不可為空）")
            return
        # 檢查 task_code 重複（P7）
        seen: set = set()
        dups: list = []
        for r in rows:
            tc = r.get("task_code", "")
            if tc in seen:
                dups.append(tc)
            seen.add(tc)
        if dups:
            dup_str = "、".join(sorted(set(dups)))
            reply = QMessageBox.warning(
                self, "任務代碼重複",
                f"以下任務代碼重複出現：{dup_str}\n"
                "重複的代碼在 Excel 匯出時行為指標可能互相覆蓋。\n\n"
                "是否仍要繼續？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        self._competency_rows = rows
        self._level = self._level_spin.value()
        self._current_task_idx = 0

        # 初始化或更新逐任務狀態
        n = len(rows)
        if len(self._task_states) != n:
            self._task_states = [TaskLLMState.IDLE] * n
            self._task_hashes = [""] * n
        else:
            # 任務數量相同：已完成但內容有變的標記為 STALE
            for i, row in enumerate(rows):
                if (self._task_states[i] == TaskLLMState.DONE
                        and _task_hash(row) != self._task_hashes[i]):
                    self._task_states[i] = TaskLLMState.STALE

        self._ensure_llm_worker()
        self._detail_update_display()
        self.stack.setCurrentIndex(3)

    # ─────────────────────────────────────
    # Worker 管理
    # ─────────────────────────────────────

    def _ensure_llm_worker(self):
        if self._llm_worker is not None and self._llm_worker.isRunning():
            return
        if self._llm_worker is not None:
            try:
                self._llm_worker.task_done.disconnect()
                self._llm_worker.error.disconnect()
                self._llm_worker.worker_ready.disconnect()
            except Exception:
                pass
        self._llm_worker = PersistentLLMWorker()
        self._llm_worker.task_done.connect(self._on_background_result)
        self._llm_worker.worker_ready.connect(
            lambda: self._status_label.setText(
                self._status_label.text().replace("AI 模型載入中...", "") + "  AI 就緒"
            )
        )
        self._llm_worker.error.connect(
            lambda msg: self._status_label.setText(f"⚠ AI 子程序：{msg}")
        )
        self._status_label.setText(
            self._status_label.text().split("  AI")[0] + "  AI 模型載入中..."
        )
        self._llm_worker.start_worker()

    def _build_task_args_for_worker(self, row: dict) -> dict:
        return {
            "position":           self._position,
            "task_name":          row.get("task_name", ""),
            "user_description":   row.get("user_description", ""),
            "standard_behaviors": row.get("_behaviors", []),
            "template":           row.get("template", self._analysis_template),
            "level":              row.get("level", 3),
            "user_output":        row.get("user_output", ""),
        }

    # ─────────────────────────────────────
    # 詳細填寫（Page 3）
    # ─────────────────────────────────────

    def _detail_update_display(self):
        """更新詳細填寫頁的顯示內容。"""
        rows = self._competency_rows
        idx  = self._current_task_idx
        total = len(rows)

        self._detail_progress_label.setText(f"任務 {idx+1} / {total}")
        self._detail_progress_bar.setMaximum(total)
        self._detail_progress_bar.setValue(idx + 1)

        row = rows[idx]
        self._detail_task_card.setTitle(f"任務 {row.get('task_code','')}")
        self._detail_task_code_lbl.setText(
            f"主責：{row.get('resp_code','')} {row.get('resp_name','')}")
        self._detail_task_name_lbl.setText(row.get("task_name", ""))

        self._detail_desc.setText(row.get("user_description", ""))
        self._detail_output.setText(row.get("user_output", ""))

        self._update_detail_llm_badge(idx)
        self._update_detail_bg_status()

        # 恢復此任務的模板選擇，預設跟隨全域設定
        tpl_key = row.get("template", self._analysis_template)
        for i in range(self._detail_tpl_combo.count()):
            if self._detail_tpl_combo.itemData(i) == tpl_key:
                self._detail_tpl_combo.setCurrentIndex(i)
                break

        if idx == 0:
            self._btn_detail_prev.setText("← 返回編輯器")
            self._btn_detail_prev.setEnabled(True)
        else:
            self._btn_detail_prev.setText("← 上一個任務")
            self._btn_detail_prev.setEnabled(True)
        is_last = (idx == total - 1)
        if is_last:
            self._btn_detail_next.setText("完成，查看 AI 建議  →")
            self._btn_detail_next.setObjectName("success")
        else:
            self._btn_detail_next.setText("下一個任務  →")
            self._btn_detail_next.setObjectName("primary")
        style = self._btn_detail_next.style()
        if style is not None:
            style.unpolish(self._btn_detail_next)
            style.polish(self._btn_detail_next)

    _STATE_DISPLAY = {
        TaskLLMState.IDLE:    ("●  未提交", "#aab4be"),
        TaskLLMState.PENDING: ("●  分析中", "#e67e22"),
        TaskLLMState.DONE:    ("✓  已完成", "#27ae60"),
        TaskLLMState.STALE:   ("↻  需更新", "#2980b9"),
    }

    def _update_detail_bg_status(self):
        """更新 Detail 頁底部的後台分析全局進度提示。"""
        if not self._task_states:
            self._detail_bg_status.setText("")
            return
        n       = len(self._task_states)
        done    = sum(1 for s in self._task_states if s == TaskLLMState.DONE)
        pending = sum(1 for s in self._task_states if s == TaskLLMState.PENDING)
        if done == n:
            self._detail_bg_status.setText("後台分析：全部完成 ✓")
            self._detail_bg_status.setStyleSheet("color:#27ae60; font-size:8.5pt;")
        elif pending > 0 or done > 0:
            self._detail_bg_status.setText(
                f"後台分析中：已完成 {done} / {n} 個任務，{pending} 個等待中...")
            self._detail_bg_status.setStyleSheet("color:#e67e22; font-size:8.5pt;")
        else:
            self._detail_bg_status.setText(
                "（填寫並儲存每個任務描述後，AI 將在背景自動分析）")
            self._detail_bg_status.setStyleSheet("color:#aab4be; font-size:8.5pt;")

    def _update_detail_llm_badge(self, idx: int):
        if not self._task_states or idx >= len(self._task_states):
            self._detail_llm_badge.setText("")
            return
        txt, color = self._STATE_DISPLAY.get(self._task_states[idx], ("", "#888"))
        self._detail_llm_badge.setText(txt)
        self._detail_llm_badge.setStyleSheet(f"color:{color}; font-size:9pt; font-weight:bold;")

    def _detail_save_current(self):
        """把目前的輸入存回 _competency_rows，並在描述有變更時自動提交 LLM。"""
        idx = self._current_task_idx
        row = self._competency_rows[idx]
        row["user_description"] = self._detail_desc.toPlainText().strip()
        row["user_output"]      = self._detail_output.toPlainText().strip()
        row["template"]         = self._detail_tpl_combo.currentData()

        if not row.get("user_description"):
            return   # 描述為空，不提交

        if not self._task_states or idx >= len(self._task_states):
            return

        new_hash = _task_hash(row)
        if new_hash == self._task_hashes[idx]:
            return   # 內容未變，不重複提交

        self._task_hashes[idx] = new_hash
        if self._llm_worker and self._llm_worker.isRunning():
            # Worker 已就緒：立即提交，設為 PENDING
            self._task_states[idx] = TaskLLMState.PENDING
            self._llm_worker.submit(idx, self._build_task_args_for_worker(row), new_hash)
        else:
            # Worker 尚未就緒（模型載入中）：設為 IDLE，等 _render_suggest_page 統一提交
            self._task_states[idx] = TaskLLMState.IDLE
        self._update_detail_llm_badge(idx)
        self._update_detail_bg_status()

    def _detail_prev(self):
        self._detail_save_current()
        if self._current_task_idx == 0:
            self.stack.setCurrentIndex(2)
        else:
            self._current_task_idx -= 1
            self._detail_update_display()

    def _detail_next(self):
        self._detail_save_current()
        if self._current_task_idx < len(self._competency_rows) - 1:
            self._current_task_idx += 1
            self._detail_update_display()
        else:
            self._goto_suggest()

    # ─────────────────────────────────────
    # LLM 建議（Page 4）
    # ─────────────────────────────────────

    def _goto_suggest(self):
        self.stack.setCurrentIndex(4)
        self._render_suggest_page()

    _TPL_DESC: dict[str, str] = {
        "AUTO": "自動", "ABCD": "ABCD", "5W2H": "5W2H", "STAR": "STAR",
    }
    _TPL_BADGE: dict[str, tuple[str, str]] = {
        "5W2H": ("#e8f5e9", "#2e7d32"),
        "ABCD": ("#e3f2fd", "#1565c0"),
        "STAR": ("#fff3e0", "#e65100"),
    }

    def _render_suggest_page(self):
        """清空並重新渲染建議頁，已完成任務直接填入，待完成任務顯示佔位。"""
        while self._suggest_layout.count():
            item = self._suggest_layout.takeAt(0)
            if item is None:
                continue
            w = item.widget()
            if w is not None:
                w.deleteLater()

        rows = self._competency_rows
        n = len(rows)
        self._suggest_checks = [None] * n
        self._suggest_boxes  = [None] * n

        done_count = sum(1 for s in self._task_states if s == TaskLLMState.DONE)

        if done_count < n and any(s in (TaskLLMState.PENDING, TaskLLMState.IDLE, TaskLLMState.STALE)
                                  for s in self._task_states):
            self._suggest_status_lbl.setText(f"AI 分析中...（已完成 {done_count}/{n}）")
            self._suggest_status_lbl.setStyleSheet(
                "color:#ffffff; background:#e67e22; border-radius:4px; padding:3px 10px; font-size:9pt;")
        else:
            self._suggest_status_lbl.setText("AI 分析完成")
            self._suggest_status_lbl.setStyleSheet(
                "color:#ffffff; background:#27ae60; border-radius:4px; padding:3px 10px; font-size:9pt;")

        self._suggest_progress.setMaximum(n)
        self._suggest_progress.setValue(done_count)
        has_pending = any(s == TaskLLMState.PENDING for s in self._task_states)
        self._btn_confirm_suggest.setEnabled(not has_pending)

        # 模板摘要標籤
        tpl_keys = [row.get("template", self._analysis_template) for row in rows]
        unique = sorted(set(tpl_keys), key=tpl_keys.index)
        if len(unique) == 1:
            self._tpl_label.setText(
                f"框架：{self._TPL_DESC.get(unique[0], unique[0])}（全部任務）")
        else:
            summary = "、".join(self._TPL_DESC.get(k, k) for k in unique)
            self._tpl_label.setText(f"框架：{summary}（各任務獨立設定，見下方標籤）")

        # 建立各任務佔位框並填入已完成的結果
        for idx, row in enumerate(rows):
            box = self._populate_task_box(idx, row)
            self._suggest_layout.addWidget(box)
            self._suggest_boxes[idx] = box
            if self._task_states[idx] == TaskLLMState.DONE:
                self._fill_task_result(
                    idx,
                    row.get("_llm_indicators", []),
                    row.get("_llm_template_used", ""),
                )

        # 提交所有尚未完成的任務（IDLE / STALE / PENDING 都重提，保證 Worker 就緒後有接到）
        # PENDING 可能是模型載入中時未實際進入 queue 的任務，重提不影響正確性（hash 驗證）
        if self._llm_worker and self._llm_worker.isRunning():
            for idx, row in enumerate(rows):
                state = self._task_states[idx]
                if state != TaskLLMState.DONE and row.get("user_description"):
                    new_hash = _task_hash(row)
                    self._task_hashes[idx] = new_hash
                    self._task_states[idx] = TaskLLMState.PENDING
                    self._llm_worker.submit(idx, self._build_task_args_for_worker(row), new_hash)

    def _rerun_llm(self):
        """重設所有任務狀態並重新分析。"""
        for i in range(len(self._task_states)):
            self._task_states[i] = TaskLLMState.IDLE
            self._task_hashes[i] = ""
            row = self._competency_rows[i]
            row.pop("_llm_indicators", None)
            row.pop("_llm_template_used", None)
        self._render_suggest_page()

    def _populate_task_box(self, idx: int, row: dict) -> "QGroupBox":
        """建立任務 GroupBox，預設顯示「分析中」佔位。"""
        box = QGroupBox(f"{row.get('task_code','')}  {row.get('task_name','')}")
        box_v = QVBoxLayout(box)
        box_v.setSpacing(4)

        state = self._task_states[idx] if idx < len(self._task_states) else TaskLLMState.IDLE
        if state == TaskLLMState.DONE:
            placeholder = None
        else:
            if not row.get("user_description"):
                placeholder = QLabel("（此任務未填寫描述，跳過 AI 分析）")
                placeholder.setStyleSheet("color:#aab4be; font-style:italic; font-size:9pt;")
            else:
                placeholder = QLabel("● AI 分析中，請稍候...")
                placeholder.setStyleSheet("color:#e67e22; font-style:italic; font-size:9pt;")
            box_v.addWidget(placeholder)

        extra_lbl = QLabel("手動補充行為指標（每行一條）：")
        extra_lbl.setStyleSheet("color:#4a5568; font-size:9pt;")
        box_v.addWidget(extra_lbl)
        extra = QTextEdit()
        extra.setFixedHeight(60)
        extra.setPlaceholderText("選填，直接輸入...")
        box_v.addWidget(extra)

        self._suggest_checks[idx] = {"checks": [], "extra": extra, "placeholder": placeholder}
        return box

    def _fill_task_result(self, idx: int, behaviors: list, template_used: str):
        """用 LLM 結果填入對應 GroupBox，移除佔位、插入模板標籤與核取方塊。"""
        entry = self._suggest_checks[idx] if idx < len(self._suggest_checks) else None
        box   = self._suggest_boxes[idx]  if idx < len(self._suggest_boxes)  else None
        if entry is None or box is None:
            return

        box_v = box.layout()

        placeholder = entry.get("placeholder")
        if placeholder is not None:
            box_v.removeWidget(placeholder)
            placeholder.deleteLater()
            entry["placeholder"] = None

        insert_pos = 0
        checks: list[tuple[QCheckBox, QLineEdit]] = []

        if template_used:
            bg, fg = self._TPL_BADGE.get(template_used, ("#f0f0f0", "#555555"))
            tpl_lbl = QLabel(f"框架：{template_used}")
            tpl_lbl.setStyleSheet(
                f"color:{fg}; background:{bg}; border-radius:3px;"
                f" padding:1px 10px; font-size:8pt; font-weight:bold;")
            tpl_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
            box_v.insertWidget(insert_pos, tpl_lbl)
            insert_pos += 1

        if behaviors:
            for b in behaviors:
                row_w = QWidget()
                row_h = QHBoxLayout(row_w)
                row_h.setContentsMargins(0, 0, 0, 0)
                row_h.setSpacing(6)
                cb = QCheckBox()
                cb.setChecked(True)
                cb.setFixedWidth(20)
                le = QLineEdit(b)
                le.setStyleSheet(
                    "border:1px solid #cbd5e0; border-radius:3px;"
                    "padding:3px 6px; background:#fff; font-size:9pt;")
                cb.toggled.connect(le.setEnabled)
                row_h.addWidget(cb, 0)
                row_h.addWidget(le, 1)
                box_v.insertWidget(insert_pos, row_w)
                insert_pos += 1
                checks.append((cb, le))
        else:
            no_result = QLabel("（AI 未能生成行為指標，可手動填寫）")
            no_result.setStyleSheet("color:#e74c3c; font-style:italic;")
            box_v.insertWidget(insert_pos, no_result)

        entry["checks"] = checks

    def _on_background_result(self, idx: int, indicators: list,
                               template_used: str, task_hash: str):
        """接收長駐 worker 回傳結果，更新狀態與 UI。"""
        if idx >= len(self._task_states):
            return
        if task_hash != self._task_hashes[idx]:
            return   # 描述已更新，丟棄過期結果

        row = self._competency_rows[idx]
        row["_llm_indicators"]    = indicators
        row["_llm_template_used"] = template_used
        self._task_states[idx]    = TaskLLMState.DONE

        # Detail 頁：更新 badge 與全局進度
        if self.stack.currentIndex() == 3:
            if self._current_task_idx == idx:
                self._update_detail_llm_badge(idx)
            self._update_detail_bg_status()

        # Suggest 頁：填入結果並更新進度
        if self.stack.currentIndex() == 4:
            self._fill_task_result(idx, indicators, template_used)
            done = sum(1 for s in self._task_states if s == TaskLLMState.DONE)
            n    = len(self._task_states)
            self._suggest_progress.setValue(done)
            has_pending = any(s == TaskLLMState.PENDING for s in self._task_states)
            if not has_pending:
                self._on_all_tasks_done()
            else:
                self._suggest_status_lbl.setText(f"AI 分析中...（已完成 {done}/{n}）")

    def _on_all_tasks_done(self):
        self._suggest_status_lbl.setText("AI 分析完成")
        self._suggest_status_lbl.setStyleSheet(
            "color:#ffffff; background:#27ae60; border-radius:4px; padding:3px 10px; font-size:9pt;")
        self._btn_confirm_suggest.setEnabled(True)

    # ─────────────────────────────────────
    # 補充 & 匯出（Page 5）
    # ─────────────────────────────────────

    def _goto_supplement(self):
        """從 LLM 建議頁收集採用的行為指標，進入補充頁。"""
        for idx, entry in enumerate(self._suggest_checks):
            if entry is None:
                continue
            accepted = [le.text().strip() for cb, le in entry["checks"]
                        if cb.isChecked() and le.text().strip()]
            extra_text = entry["extra"].toPlainText().strip()
            if extra_text:
                accepted.extend([l.strip() for l in extra_text.split("\n") if l.strip()])
            self._competency_rows[idx]["behavior_accepted"] = accepted

        # 更新摘要
        lines = [
            f"職業名稱：{self._position}",
            f"職能等級：{self._level}",
            f"職能基準：{(self._matched_std or {}).get('metadata', {}).get('name', '（未使用基準）')}",
            f"工作任務數：{len(self._competency_rows)} 個",
        ]
        covered = sum(1 for r in self._competency_rows if r.get("behavior_accepted"))
        lines.append(f"已生成行為指標：{covered} 個任務")
        self._summary_label.setText("\n".join(lines))
        self.stack.setCurrentIndex(5)

    def _on_export(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "儲存職能說明書", f"{self._position}_職能說明書.xlsx",
            "Excel 檔案 (*.xlsx)")
        if not path:
            return
        try:
            from excel_exporter import export_competency
            data = {
                "position": self._position,
                "level":    self._level,
                "standard_code": (self._matched_std or {}).get(
                    "metadata", {}).get("code", ""),
                "standard_name": (self._matched_std or {}).get(
                    "metadata", {}).get("name", ""),
                "supplement": self._supplement_text.toPlainText().strip(),
                "rows": self._competency_rows,
                "attitudes": (self._matched_std or {}).get(
                    "competency_attitudes", []),
            }
            out = export_competency(data, Path(path))
            QMessageBox.information(self, "匯出完成", f"已儲存至：\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "匯出失敗", str(e))

    # ─────────────────────────────────────
    # 資料管理
    # ─────────────────────────────────────

    def _open_data_manager(self):
        dlg = DataManagerDialog(self._rag, self)
        dlg.rebuild_requested.connect(self._start_init)
        dlg.exec()

    # ─────────────────────────────────────
    # 視窗關閉清理
    # ─────────────────────────────────────

    def closeEvent(self, event):
        if self._llm_worker is not None:
            self._llm_worker.stop_worker()
        super().closeEvent(event)
