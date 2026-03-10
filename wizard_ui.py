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
    QScrollArea, QFrame, QComboBox, QCheckBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QColor

from wizard_rag import WizardRAG
from gap_analyzer import GapAnalyzer, UserInput5W2H, GapReport
from excel_exporter import export_to_excel


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
        """
        engine: 可選的 GraphRAGQueryEngine 實例。
                傳入時直接復用其 Embedding 模型與向量索引，省去重複載入時間。
        """
        super().__init__()
        self.setWindowTitle("職能說明書精靈")
        self.setMinimumSize(900, 680)

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
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 頂部狀態欄
        self._top_bar = self._make_top_bar()
        layout.addWidget(self._top_bar)

        # 主要頁面 (stack)
        self.stack = QStackedWidget()
        layout.addWidget(self.stack, 1)

        self._page_loading = self._make_loading_page()
        self._page_form    = self._make_form_page()
        self._page_result  = self._make_result_page()

        self.stack.addWidget(self._page_loading)  # index 0
        self.stack.addWidget(self._page_form)      # index 1
        self.stack.addWidget(self._page_result)    # index 2

    def _make_top_bar(self) -> QWidget:
        bar = QFrame()
        bar.setFixedHeight(48)
        bar.setStyleSheet("background:#2F5496; color:white;")
        h = QHBoxLayout(bar)
        h.setContentsMargins(16, 0, 16, 0)

        title = QLabel("職能說明書精靈")
        title.setFont(QFont("Microsoft JhengHei", 14, QFont.Weight.Bold))
        title.setStyleSheet("color:white;")
        h.addWidget(title)
        h.addStretch()

        self._status_label = QLabel("初始化中...")
        self._status_label.setStyleSheet("color:#cce4ff; font-size:11px;")
        h.addWidget(self._status_label)

        return bar

    def _make_loading_page(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self._loading_label = QLabel("正在載入 Embedding 模型，請稍候...")
        self._loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._loading_label.setFont(QFont("Microsoft JhengHei", 11))
        v.addWidget(self._loading_label)

        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 0)   # 不確定進度
        self._progress_bar.setFixedWidth(400)
        v.addWidget(self._progress_bar)

        btn_rebuild = QPushButton("強制重建索引")
        btn_rebuild.setFixedWidth(160)
        btn_rebuild.clicked.connect(self._on_force_rebuild)
        v.addWidget(btn_rebuild, alignment=Qt.AlignmentFlag.AlignCenter)

        return w

    def _make_form_page(self) -> QWidget:
        """Step 1: 5W2H 輸入表單"""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        scroll.setWidget(inner)
        v = QVBoxLayout(inner)
        v.setContentsMargins(24, 16, 24, 16)
        v.setSpacing(12)

        def section(title):
            gb = QGroupBox(title)
            gb.setFont(QFont("Microsoft JhengHei", 10, QFont.Weight.Bold))
            return gb

        # What
        gb_what = section("What — 做什麼")
        f = QFormLayout(gb_what)
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
        self._why_purpose = QLineEdit()
        self._why_purpose.setPlaceholderText("工作目的（例：提升品牌知名度、達成業績目標）")
        f2.addRow("工作目的：", self._why_purpose)
        v.addWidget(gb_why)

        # Who
        gb_who = section("Who — 誰做 / 與誰協作")
        f3 = QFormLayout(gb_who)
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
        self._when_frequency = QComboBox()
        self._when_frequency.addItems(["每日", "每週", "每月", "每季", "專案型（不固定）", "其他"])
        self._when_frequency.setEditable(True)
        f4.addRow("執行頻率：", self._when_frequency)
        v.addWidget(gb_when)

        # Where
        gb_where = section("Where — 在哪做")
        f5 = QFormLayout(gb_where)
        self._where_env = QLineEdit()
        self._where_env.setPlaceholderText("工作環境/地點（例：辦公室、工廠現場、遠端居家）")
        f5.addRow("工作環境：", self._where_env)
        v.addWidget(gb_where)

        # How
        gb_how = section("How — 如何做")
        f6 = QFormLayout(gb_how)
        self._how_skills = QTextEdit()
        self._how_skills.setPlaceholderText("使用的技能/工具/方法（例：Excel 資料分析、Python 自動化、溝通協商）")
        self._how_skills.setFixedHeight(72)
        f6.addRow("技能/工具：", self._how_skills)
        v.addWidget(gb_how)

        # How Much
        gb_howmuch = section("How Much — 做到什麼程度")
        f7 = QFormLayout(gb_howmuch)
        self._how_much = QLineEdit()
        self._how_much.setPlaceholderText("績效指標（例：按時完成率 95%、客戶滿意度 4.5/5、錯誤率 <2%）")
        f7.addRow("績效指標：", self._how_much)
        v.addWidget(gb_howmuch)

        # 按鈕列
        btn_row = QHBoxLayout()
        btn_clear = QPushButton("清除")
        btn_clear.clicked.connect(self._clear_form)
        self._btn_analyze = QPushButton("開始分析 →")
        self._btn_analyze.setFixedHeight(36)
        self._btn_analyze.setStyleSheet("background:#2F5496; color:white; font-weight:bold; font-size:13px;")
        self._btn_analyze.clicked.connect(self._on_analyze)
        btn_analyze = self._btn_analyze
        btn_row.addWidget(btn_clear)
        btn_row.addStretch()
        btn_row.addWidget(btn_analyze)
        v.addLayout(btn_row)

        return scroll

    def _make_result_page(self) -> QWidget:
        """Step 2+3: 結果與缺口"""
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(16, 12, 16, 12)
        v.setSpacing(8)

        # 頂部工具列（無返回按鈕）
        toolbar = QHBoxLayout()
        self._result_status = QLabel("")
        self._result_status.setFont(QFont("Microsoft JhengHei", 10))
        self._btn_export = QPushButton("匯出 Excel")
        self._btn_export.setStyleSheet(
            "background:#217346; color:white; font-weight:bold; padding:4px 12px;"
        )
        self._btn_export.setEnabled(False)
        self._btn_export.clicked.connect(self._on_export)
        toolbar.addWidget(self._result_status, 1)
        toolbar.addWidget(self._btn_export)
        v.addLayout(toolbar)

        # 水平分割：左（可編輯輸入 + 下拉選單）、右（缺口詳情）
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # ── 左側面板 ──────────────────────────────────
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(0, 0, 4, 0)
        lv.setSpacing(6)

        # 職能基準下拉選單
        lbl_match = QLabel("相似職能基準：")
        lbl_match.setFont(QFont("Microsoft JhengHei", 9, QFont.Weight.Bold))
        lv.addWidget(lbl_match)
        self._match_combo = QComboBox()
        self._match_combo.setFont(QFont("Microsoft JhengHei", 9))
        self._match_combo.currentIndexChanged.connect(self._on_match_selected)
        lv.addWidget(self._match_combo)

        # 分隔線
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#ccc;")
        lv.addWidget(sep)

        # 可編輯的 5W2H 欄位
        lbl_edit = QLabel("工作內容（可直接修改後重新分析）：")
        lbl_edit.setFont(QFont("Microsoft JhengHei", 9, QFont.Weight.Bold))
        lv.addWidget(lbl_edit)

        edit_scroll = QScrollArea()
        edit_scroll.setWidgetResizable(True)
        edit_scroll.setFrameShape(QFrame.Shape.NoFrame)
        edit_inner = QWidget()
        edit_scroll.setWidget(edit_inner)
        form = QFormLayout(edit_inner)
        form.setContentsMargins(2, 4, 2, 4)
        form.setSpacing(6)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self._r_who_role = QLineEdit()
        self._r_who_role.setPlaceholderText("自身職稱")
        form.addRow("角色：", self._r_who_role)

        self._r_what_tasks = QTextEdit()
        self._r_what_tasks.setPlaceholderText("主要工作任務")
        self._r_what_tasks.setFixedHeight(58)
        form.addRow("工作任務：", self._r_what_tasks)

        self._r_what_outputs = QLineEdit()
        self._r_what_outputs.setPlaceholderText("工作產出/交付物")
        form.addRow("工作產出：", self._r_what_outputs)

        self._r_why_purpose = QLineEdit()
        self._r_why_purpose.setPlaceholderText("工作目的")
        form.addRow("工作目的：", self._r_why_purpose)

        self._r_how_skills = QTextEdit()
        self._r_how_skills.setPlaceholderText("技能/工具/方法")
        self._r_how_skills.setFixedHeight(58)
        form.addRow("技能/工具：", self._r_how_skills)

        self._r_how_much = QLineEdit()
        self._r_how_much.setPlaceholderText("績效指標")
        form.addRow("績效指標：", self._r_how_much)

        self._r_when_frequency = QLineEdit()
        self._r_when_frequency.setPlaceholderText("執行頻率")
        form.addRow("執行頻率：", self._r_when_frequency)

        self._r_where_env = QLineEdit()
        self._r_where_env.setPlaceholderText("工作環境")
        form.addRow("工作環境：", self._r_where_env)

        self._r_who_collaborate = QLineEdit()
        self._r_who_collaborate.setPlaceholderText("協作對象")
        form.addRow("協作對象：", self._r_who_collaborate)

        lv.addWidget(edit_scroll, 1)

        # 重新分析按鈕
        self._btn_reanalyze = QPushButton("重新分析 →")
        self._btn_reanalyze.setFixedHeight(32)
        self._btn_reanalyze.setStyleSheet(
            "background:#2F5496; color:white; font-weight:bold;"
        )
        self._btn_reanalyze.clicked.connect(self._on_reanalyze)
        lv.addWidget(self._btn_reanalyze)

        splitter.addWidget(left)

        # ── 右側面板 ──────────────────────────────────
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(4, 0, 0, 0)
        lbl_gap = QLabel("缺口分析結果：")
        lbl_gap.setFont(QFont("Microsoft JhengHei", 9, QFont.Weight.Bold))
        rv.addWidget(lbl_gap)
        self._gap_text = QTextEdit()
        self._gap_text.setReadOnly(True)
        self._gap_text.setFont(QFont("Microsoft JhengHei", 10))
        rv.addWidget(self._gap_text)
        splitter.addWidget(right)

        splitter.setSizes([320, 540])
        v.addWidget(splitter, 1)

        # ── 底部確認列 ────────────────────────────────
        confirm_bar = QHBoxLayout()
        self._confirm_check = QCheckBox(
            "我已確認以上缺口分析結果正確無誤，同意匯出職能說明書"
        )
        self._confirm_check.setFont(QFont("Microsoft JhengHei", 10))
        self._confirm_check.setStyleSheet("color:#333;")
        self._confirm_check.toggled.connect(self._btn_export.setEnabled)
        confirm_bar.addWidget(self._confirm_check)
        confirm_bar.addStretch()
        v.addLayout(confirm_bar)

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
        # 重置確認狀態（每次新分析都需要重新確認）
        self._confirm_check.setChecked(False)
        self._btn_export.setEnabled(False)

        # 填入可編輯的 5W2H 欄位
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

        # 填入下拉選單（暫時斷開 signal 避免觸發 _on_match_selected）
        self._match_combo.blockSignals(True)
        self._match_combo.clear()
        for r in report.matched_standards:
            self._match_combo.addItem(
                f"[{r['score']:.2f}] {r['standard_name']}（{r['standard_code']}）"
            )
        self._match_combo.blockSignals(False)

        score = report.completeness_score
        self._result_status.setText(
            f"最佳匹配：{report.best_standard_name}  |  完整度：{score}%"
        )

        if self.analyzer:
            text = self.analyzer.get_summary_text(report)
            self._gap_text.setPlainText(text)

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

        lines = []
        basic = std_data.get("basic_info") or std_data.get("metadata") or {}
        lines.append(f"【{basic.get('name', std_code)}】（{std_code}）")
        lines.append(f"職能級別：{basic.get('level', '')}")
        lines.append(f"工作描述：{basic.get('description', '')[:200]}")
        lines.append("")
        knowledge = std_data.get("competency_knowledge") or std_data.get("knowledge") or []
        skills = std_data.get("competency_skills") or std_data.get("skills") or []
        lines.append(f"知識項目（{len(knowledge)} 項）：")
        for k in knowledge[:8]:
            lines.append(f"  [{k.get('code','')}] {k.get('name','')}")
        lines.append("")
        lines.append(f"技能項目（{len(skills)} 項）：")
        for s in skills[:8]:
            lines.append(f"  [{s.get('code','')}] {s.get('name','')}")
        lines.append("")
        tasks = std_data.get("competency_tasks") or []
        lines.append(f"工作任務（{len(tasks)} 項）：")
        for t in tasks[:6]:
            lines.append(f"  [{t.get('task_id','')}] {t.get('task_name','')}")

        self._gap_text.setPlainText("\n".join(lines))

    def _collect_result_input(self) -> UserInput5W2H:
        """從結果頁可編輯欄位收集 5W2H 輸入"""
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
