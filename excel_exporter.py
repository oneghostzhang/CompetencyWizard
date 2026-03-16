"""
competency_wizard/excel_exporter.py
將 5W2H 輸入與缺口分析結果輸出為 Excel 檔案
需要 openpyxl
"""

from pathlib import Path
from typing import Optional
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side
    )
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

from gap_analyzer import GapReport, UserInput5W2H


# ─────────────────────────────────────────
# 顏色常數
# ─────────────────────────────────────────
COLOR_HEADER_BG   = "2F5496"   # 深藍色表頭
COLOR_HEADER_FONT = "FFFFFF"   # 白字
COLOR_SUB_BG      = "D9E1F2"   # 淺藍色小節
COLOR_COVERED_BG  = "E2EFDA"   # 淺綠：已涵蓋
COLOR_GAP_HIGH    = "FFE0E0"   # 淺紅：高嚴重度缺口
COLOR_GAP_MED     = "FFF2CC"   # 淺黃：中嚴重度缺口
COLOR_GAP_LOW     = "F8F8F8"   # 淺灰：低嚴重度缺口
COLOR_SCORE_BG    = "F4B942"   # 橘色分數欄


def export_to_excel(
    report: GapReport,
    output_path: Optional[Path] = None,
    role_name: str = "",
) -> Path:
    """
    輸出職能說明書 Excel

    Sheets:
      1. 職能說明書摘要   — 5W2H + 匹配職能基準基本資訊
      2. 工作任務對照     — 使用者任務 vs 標準任務
      3. 知識技能對照     — 知識/技能 covered / gap
      4. 缺口分析報告     — 所有缺口項目 + 建議
      5. 完整職能基準資料 — 直接列出 best_standard_data 供參考

    回傳輸出路徑。
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("缺少 openpyxl，請執行 pip install openpyxl")

    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_role = role_name.replace("/", "_").replace("\\", "_") or "職能說明書"
        output_path = Path.cwd() / f"{safe_role}_{ts}.xlsx"

    output_path = Path(output_path)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 移除預設空白頁

    _sheet_summary(wb, report, role_name)
    _sheet_confirmed(wb, report, role_name)
    _sheet_tasks(wb, report)
    _sheet_knowledge_skills(wb, report)
    _sheet_gap(wb, report)
    _sheet_full_standard(wb, report)

    wb.save(output_path)
    return output_path


# ─────────────────────────────────────────
# Sheet 1: 職能說明書摘要
# ─────────────────────────────────────────

def _sheet_summary(wb, report: GapReport, role_name: str):
    ws = wb.create_sheet("職能說明書摘要")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 60

    ui: UserInput5W2H = report.user_input

    # 標題
    ws.merge_cells("A1:B1")
    _write(ws, "A1", f"職能說明書　{role_name}", bold=True, size=14,
           bg=COLOR_HEADER_BG, font_color=COLOR_HEADER_FONT, align="center")

    ws["A2"] = "建立日期"
    ws["B2"] = datetime.now().strftime("%Y-%m-%d")

    ws["A3"] = "匹配職能基準"
    ws["B3"] = f"{report.best_standard_name}（{report.best_standard_code}）"

    ws["A4"] = "員工確認完整度"
    ws["B4"] = f"{report.completeness_score}%"
    _style(ws["B4"], bg=COLOR_SCORE_BG, bold=True)

    # 5W2H 表格
    _section_header(ws, "A6", "B6", "5W2H 工作內容描述")
    rows = [
        ("What — 工作任務", ui.what_tasks),
        ("What — 工作產出", ui.what_outputs),
        ("Why — 工作目的", ui.why_purpose),
        ("Who — 自身角色", ui.who_role),
        ("Who — 協作對象", ui.who_collaborate),
        ("When — 執行頻率", ui.when_frequency),
        ("Where — 工作環境", ui.where_environment),
        ("How — 技能/工具", ui.how_skills),
        ("How Much — 績效指標", ui.how_much_kpi),
    ]
    for i, (label, value) in enumerate(rows, start=7):
        ws[f"A{i}"] = label
        ws[f"B{i}"] = value
        _style(ws[f"A{i}"], bold=True)
        ws.row_dimensions[i].height = 30

    _auto_wrap(ws, "B", 7, 7 + len(rows))


# ─────────────────────────────────────────
# Sheet 2: 我的職能確認（員工已確認具備的項目）
# ─────────────────────────────────────────

def _sheet_confirmed(wb, report: GapReport, role_name: str):
    ws = wb.create_sheet("我的職能確認")
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 44

    # 標題
    ws.merge_cells("A1:C1")
    _write(ws, "A1", f"我的職能確認 — {role_name}",
           bold=True, size=13, bg=COLOR_HEADER_BG,
           font_color=COLOR_HEADER_FONT, align="center")

    std = report.best_standard_data or {}
    task_id_map = {t.get("task_name", ""): t.get("task_id", "")
                   for t in std.get("competency_tasks", [])}
    k_code_map  = {k.get("name", ""): k.get("code", "")
                   for k in std.get("competency_knowledge", [])}
    s_code_map  = {s.get("name", ""): s.get("code", "")
                   for s in std.get("competency_skills", [])}

    row = 3

    # ── 工作任務 ─────────────────────────────
    ws.merge_cells(f"A{row}:C{row}")
    _write(ws, f"A{row}", "📋 工作任務", bold=True, bg=COLOR_SUB_BG)
    row += 1

    if report.covered_tasks:
        _header_row(ws, row, ["任務代碼", "任務名稱", "說明"])
        row += 1
        for name in report.covered_tasks:
            tid  = task_id_map.get(name, "")
            # 取得對應任務的 output 說明
            output = next((t.get("output", "") for t in std.get("competency_tasks", [])
                           if t.get("task_name") == name), "")
            ws[f"A{row}"] = tid
            ws[f"B{row}"] = name
            ws[f"C{row}"] = output if isinstance(output, str) else ""
            for col in ["A", "B", "C"]:
                _style(ws[f"{col}{row}"], bg=COLOR_COVERED_BG)
            row += 1
    else:
        ws[f"A{row}"] = "（尚未確認任何工作任務）"
        row += 1

    row += 1

    # ── 具備知識 ─────────────────────────────
    ws.merge_cells(f"A{row}:C{row}")
    _write(ws, f"A{row}", "📖 具備知識", bold=True, bg=COLOR_SUB_BG)
    row += 1

    if report.covered_knowledge:
        _header_row(ws, row, ["類別", "代碼", "知識名稱"])
        row += 1
        for name in report.covered_knowledge:
            ws[f"A{row}"] = "知識"
            ws[f"B{row}"] = k_code_map.get(name, "")
            ws[f"C{row}"] = name
            for col in ["A", "B", "C"]:
                _style(ws[f"{col}{row}"], bg=COLOR_COVERED_BG)
            row += 1
    else:
        ws[f"A{row}"] = "（尚未確認任何知識項目）"
        row += 1

    row += 1

    # ── 具備技能 ─────────────────────────────
    ws.merge_cells(f"A{row}:C{row}")
    _write(ws, f"A{row}", "🔧 具備技能", bold=True, bg=COLOR_SUB_BG)
    row += 1

    if report.covered_skills:
        _header_row(ws, row, ["類別", "代碼", "技能名稱"])
        row += 1
        for name in report.covered_skills:
            ws[f"A{row}"] = "技能"
            ws[f"B{row}"] = s_code_map.get(name, "")
            ws[f"C{row}"] = name
            for col in ["A", "B", "C"]:
                _style(ws[f"{col}{row}"], bg=COLOR_COVERED_BG)
            row += 1
    else:
        ws[f"A{row}"] = "（尚未確認任何技能項目）"
        row += 1

    _auto_wrap(ws, "C", 3, row)


# ─────────────────────────────────────────
# Sheet 3: 工作任務對照
# ─────────────────────────────────────────

def _sheet_tasks(wb, report: GapReport):
    ws = wb.create_sheet("工作任務對照")
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 14

    _header_row(ws, 1, ["任務代碼", "工作任務名稱", "狀態"])

    # 建立 task_name → task_id 的查找表
    std = report.best_standard_data or {}
    task_id_map = {
        t.get("task_name", ""): t.get("task_id", "")
        for t in std.get("competency_tasks", [])
    }

    row = 2
    for name in report.covered_tasks:
        ws[f"A{row}"] = task_id_map.get(name, "")
        ws[f"B{row}"] = name
        ws[f"C{row}"] = "✓ 已涵蓋"
        for col in ["A", "B", "C"]:
            _style(ws[f"{col}{row}"], bg=COLOR_COVERED_BG)
        row += 1

    for g in report.gap_tasks:
        ws[f"A{row}"] = g.code
        ws[f"B{row}"] = g.name
        ws[f"C{row}"] = "△ 缺口"
        bg = COLOR_GAP_HIGH if g.severity == "high" else (
             COLOR_GAP_LOW if g.severity == "low" else COLOR_GAP_MED)
        for col in ["A", "B", "C"]:
            _style(ws[f"{col}{row}"], bg=bg)
        row += 1

    _auto_wrap(ws, "B", 2, row)


# ─────────────────────────────────────────
# Sheet 3: 知識技能對照
# ─────────────────────────────────────────

def _sheet_knowledge_skills(wb, report: GapReport):
    ws = wb.create_sheet("知識技能對照")
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 14

    _header_row(ws, 1, ["類別", "代碼", "名稱", "狀態"])

    # 建立 name → code 查找表
    std = report.best_standard_data or {}
    k_code_map = {k.get("name", ""): k.get("code", "") for k in std.get("competency_knowledge", [])}
    s_code_map = {s.get("name", ""): s.get("code", "") for s in std.get("competency_skills", [])}

    row = 2
    for name in report.covered_knowledge:
        ws[f"A{row}"] = "知識"
        ws[f"B{row}"] = k_code_map.get(name, "")
        ws[f"C{row}"] = name
        ws[f"D{row}"] = "✓ 已涵蓋"
        for col in ["A", "B", "C", "D"]:
            _style(ws[f"{col}{row}"], bg=COLOR_COVERED_BG)
        row += 1

    for g in report.gap_knowledge:
        ws[f"A{row}"] = "知識"
        ws[f"B{row}"] = g.code
        ws[f"C{row}"] = g.name
        ws[f"D{row}"] = "△ 缺口"
        bg = COLOR_GAP_HIGH if g.severity == "high" else (
             COLOR_GAP_LOW if g.severity == "low" else COLOR_GAP_MED)
        for col in ["A", "B", "C", "D"]:
            _style(ws[f"{col}{row}"], bg=bg)
        row += 1

    for name in report.covered_skills:
        ws[f"A{row}"] = "技能"
        ws[f"B{row}"] = s_code_map.get(name, "")
        ws[f"C{row}"] = name
        ws[f"D{row}"] = "✓ 已涵蓋"
        for col in ["A", "B", "C", "D"]:
            _style(ws[f"{col}{row}"], bg=COLOR_COVERED_BG)
        row += 1

    for g in report.gap_skills:
        ws[f"A{row}"] = "技能"
        ws[f"B{row}"] = g.code
        ws[f"C{row}"] = g.name
        ws[f"D{row}"] = "△ 缺口"
        bg = COLOR_GAP_HIGH if g.severity == "high" else (
             COLOR_GAP_LOW if g.severity == "low" else COLOR_GAP_MED)
        for col in ["A", "B", "C", "D"]:
            _style(ws[f"{col}{row}"], bg=bg)
        row += 1

    _auto_wrap(ws, "C", 2, row)


# ─────────────────────────────────────────
# Sheet 4: 缺口分析報告
# ─────────────────────────────────────────

def _sheet_gap(wb, report: GapReport):
    ws = wb.create_sheet("缺口分析報告")
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 50
    ws.column_dimensions["E"].width = 12

    _header_row(ws, 1, ["缺口類別", "代碼", "名稱", "說明", "嚴重度"])

    row = 2
    all_gaps = (
        [(g, "任務") for g in report.gap_tasks] +
        [(g, "知識") for g in report.gap_knowledge] +
        [(g, "技能") for g in report.gap_skills] +
        [(g, "行為指標") for g in report.gap_behaviors] +
        [(g, "工作產出") for g in report.gap_outputs]
    )

    severity_map = {"high": "高", "medium": "中", "low": "低"}
    bg_map = {"high": COLOR_GAP_HIGH, "medium": COLOR_GAP_MED, "low": COLOR_GAP_LOW}

    for g, cat in all_gaps:
        ws[f"A{row}"] = cat
        ws[f"B{row}"] = g.code
        ws[f"C{row}"] = g.name
        ws[f"D{row}"] = g.description
        ws[f"E{row}"] = severity_map.get(g.severity, g.severity)
        bg = bg_map.get(g.severity, COLOR_GAP_LOW)
        for col in ["A", "B", "C", "D", "E"]:
            _style(ws[f"{col}{row}"], bg=bg)
        row += 1

    if not all_gaps:
        ws.merge_cells(f"A{row}:E{row}")
        _write(ws, f"A{row}", "✓ 無明顯缺口，工作內容與職能基準高度吻合",
               bg=COLOR_COVERED_BG, bold=True, align="center")

    _auto_wrap(ws, "D", 2, row)


# ─────────────────────────────────────────
# Sheet 5: 完整職能基準資料
# ─────────────────────────────────────────

def _sheet_full_standard(wb, report: GapReport):
    ws = wb.create_sheet("完整職能基準資料")
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 70

    std = report.best_standard_data
    if not std:
        ws["A1"] = "（無職能基準資料）"
        return

    basic = std.get("basic_info") or std.get("metadata") or {}
    row = 1

    def add_row(label, value, sub=False):
        nonlocal row
        ws[f"A{row}"] = label
        ws[f"B{row}"] = str(value) if value is not None else ""
        if sub:
            _style(ws[f"A{row}"], bg=COLOR_SUB_BG, bold=True)
        else:
            _style(ws[f"A{row}"], bold=True)
        row += 1

    meta = std.get("metadata", {})
    add_row("職能基準代碼", meta.get("code", "") or basic.get("code", report.best_standard_code))
    add_row("職能基準名稱", meta.get("name", "") or basic.get("name", report.best_standard_name))
    add_row("職能類別", basic.get("category", ""))
    add_row("職能級別", basic.get("level", ""))
    add_row("工作描述", basic.get("job_description", ""))
    row += 1

    # 知識
    add_row("── 知識項目 ──", "", sub=True)
    for k in std.get("competency_knowledge", []):
        add_row(k.get("code", ""), k.get("name", ""))

    row += 1
    # 技能
    add_row("── 技能項目 ──", "", sub=True)
    for s in std.get("competency_skills", []):
        add_row(s.get("code", ""), s.get("name", ""))

    row += 1
    # 工作任務
    add_row("── 工作任務 ──", "", sub=True)
    for task in std.get("competency_tasks", []):
        add_row(task.get("task_id", ""), task.get("task_name", ""))

    _auto_wrap(ws, "B", 1, row)


# ─────────────────────────────────────────
# 樣式輔助函式
# ─────────────────────────────────────────

def _write(ws, cell_ref: str, value: str, bold=False, size=11,
           bg=None, font_color="000000", align="left"):
    cell = ws[cell_ref]
    cell.value = value
    cell.font = Font(bold=bold, size=size, color=font_color)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)


def _style(cell, bg=None, bold=False):
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    if bold:
        cell.font = Font(bold=True)
    cell.alignment = Alignment(vertical="center", wrap_text=True)


def _header_row(ws, row_num: int, labels: list):
    for col, label in enumerate(labels, start=1):
        cell = ws.cell(row=row_num, column=col, value=label)
        cell.font = Font(bold=True, color=COLOR_HEADER_FONT)
        cell.fill = PatternFill("solid", fgColor=COLOR_HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_num].height = 22


def _section_header(ws, cell1: str, cell2: str, title: str):
    ws.merge_cells(f"{cell1}:{cell2}")
    _write(ws, cell1, title, bold=True, bg=COLOR_SUB_BG)


def _auto_wrap(ws, col: str, start_row: int, end_row: int):
    for r in range(start_row, end_row):
        cell = ws[f"{col}{r}"]
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if len(str(cell.value or "")) > 50:
            ws.row_dimensions[r].height = 45
