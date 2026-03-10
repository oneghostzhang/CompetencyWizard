"""
competency_wizard/gap_analyzer.py
比對使用者的 5W2H 輸入與 RAG 找到的職能基準，輸出缺口分析報告
"""

import re
from typing import Dict, List, Optional
from dataclasses import dataclass, field

# ── 中文正規化（借鑒自 Graph_RAG_test/competency_store.py）──────────────────
_CJK_SPACE_RE = re.compile(
    r'(?<=[\u4e00-\u9fff\u3400-\u4dbf\uff00-\uffef])\s+'
    r'(?=[\u4e00-\u9fff\u3400-\u4dbf\uff00-\uffef])'
)

def _normalize_cjk(text: str) -> str:
    """移除 PDF 擷取產生的中文字元間多餘空格"""
    return _CJK_SPACE_RE.sub('', text)


# ─────────────────────────────────────────
# 資料結構
# ─────────────────────────────────────────

@dataclass
class UserInput5W2H:
    """使用者的 5W2H 輸入"""
    # What — 做什麼
    what_tasks: str = ""          # 主要工作任務描述
    what_outputs: str = ""        # 工作產出/交付物

    # Why — 為何做
    why_purpose: str = ""         # 工作目的/意義

    # Who — 誰做/與誰協作
    who_role: str = ""            # 自己的職稱/角色
    who_collaborate: str = ""     # 協作對象

    # When — 何時做
    when_frequency: str = ""      # 執行頻率（每日/每週/每月/專案型）

    # Where — 在哪做
    where_environment: str = ""   # 工作環境/地點

    # How — 如何做
    how_skills: str = ""          # 使用的技能/工具/方法

    # How Much — 做到什麼程度
    how_much_kpi: str = ""        # 績效指標/品質標準

    def to_search_query(self) -> str:
        """組合成 RAG 查詢字串"""
        parts = []
        if self.what_tasks:
            parts.append(self.what_tasks)
        if self.what_outputs:
            parts.append(self.what_outputs)
        if self.why_purpose:
            parts.append(self.why_purpose)
        if self.how_skills:
            parts.append(self.how_skills)
        if self.how_much_kpi:
            parts.append(self.how_much_kpi)
        return " ".join(parts)

    def to_dict(self) -> Dict:
        return {
            "what_tasks": self.what_tasks,
            "what_outputs": self.what_outputs,
            "why_purpose": self.why_purpose,
            "who_role": self.who_role,
            "who_collaborate": self.who_collaborate,
            "when_frequency": self.when_frequency,
            "where_environment": self.where_environment,
            "how_skills": self.how_skills,
            "how_much_kpi": self.how_much_kpi,
        }


@dataclass
class GapItem:
    """單一缺口項目"""
    category: str           # 缺口類別：knowledge / skill / behavior / output
    code: str               # 職能代碼（K01, S02...）
    name: str               # 名稱
    description: str = ""   # 說明
    severity: str = "medium"  # high / medium / low


@dataclass
class GapReport:
    """缺口分析報告"""
    user_input: UserInput5W2H
    matched_standards: List[Dict] = field(default_factory=list)   # RAG 匹配結果
    best_standard_code: str = ""
    best_standard_name: str = ""
    best_standard_data: Optional[Dict] = None

    # 使用者已涵蓋的項目
    covered_tasks: List[str] = field(default_factory=list)
    covered_knowledge: List[str] = field(default_factory=list)
    covered_skills: List[str] = field(default_factory=list)

    # 缺口項目
    gap_tasks: List[GapItem] = field(default_factory=list)
    gap_knowledge: List[GapItem] = field(default_factory=list)
    gap_skills: List[GapItem] = field(default_factory=list)
    gap_behaviors: List[GapItem] = field(default_factory=list)
    gap_outputs: List[GapItem] = field(default_factory=list)

    completeness_score: float = 0.0   # 0-100


# ─────────────────────────────────────────
# 分析器
# ─────────────────────────────────────────

class GapAnalyzer:
    """5W2H 輸入 vs 職能基準的缺口分析"""

    def __init__(self, wizard_rag):
        """
        wizard_rag: 已初始化的 WizardRAG 實例
        """
        self.rag = wizard_rag

    def analyze(self, user_input: UserInput5W2H, top_k: int = 3) -> GapReport:
        """
        執行缺口分析
        1. 用 5W2H 查詢 RAG，找最相似的職能基準
        2. 取最佳匹配，逐項比對缺口
        3. 回傳 GapReport
        """
        report = GapReport(user_input=user_input)

        # Step 1: RAG 搜尋
        query = user_input.to_search_query()
        if not query.strip():
            return report

        results = self.rag.search(query, top_k=top_k)
        if not results:
            return report

        report.matched_standards = results

        # Step 2: 取最佳匹配（score 最高）
        best = results[0]
        report.best_standard_code = best["standard_code"]
        report.best_standard_name = best["standard_name"]

        std_data = self.rag.get_standard(best["standard_code"])
        if not std_data:
            return report

        report.best_standard_data = std_data

        # Step 3: 缺口比對
        self._analyze_tasks(user_input, std_data, report)
        self._analyze_knowledge(user_input, std_data, report)
        self._analyze_skills(user_input, std_data, report)
        self._analyze_behaviors(user_input, std_data, report)
        self._analyze_outputs(user_input, std_data, report)

        # Step 4: 計算完整度分數
        report.completeness_score = self._calc_completeness(report, std_data)

        return report

    # ─── 各維度分析 ───────────────────────────

    def _analyze_tasks(self, ui: UserInput5W2H, std: Dict, report: GapReport):
        """比對工作任務（What）"""
        user_text = _normalize_cjk((ui.what_tasks + " " + ui.what_outputs).lower())
        tasks = std.get("competency_tasks", [])

        for task in tasks:
            task_name = task.get("task_name", "")
            task_id = task.get("task_id", "")
            if not task_name:
                continue
            if self._text_contains_any(user_text, task_name.lower().split()):
                report.covered_tasks.append(task_name)
            else:
                report.gap_tasks.append(GapItem(
                    category="task",
                    code=task_id,
                    name=task_name,
                    severity=self._task_severity(task),
                ))

    def _analyze_knowledge(self, ui: UserInput5W2H, std: Dict, report: GapReport):
        """比對知識項目（How 相關）"""
        user_text = _normalize_cjk((ui.how_skills + " " + ui.what_tasks + " " + ui.why_purpose).lower())
        knowledge_list = std.get("competency_knowledge", []) or std.get("knowledge", [])

        for k in knowledge_list:
            code = k.get("code", "")
            name = k.get("name", "")
            desc = k.get("description", "")
            if not name:
                continue
            if self._text_contains_any(user_text, name.lower().split()):
                report.covered_knowledge.append(name)
            else:
                report.gap_knowledge.append(GapItem(
                    category="knowledge",
                    code=code,
                    name=name,
                    description=desc,
                    severity="medium",
                ))

    def _analyze_skills(self, ui: UserInput5W2H, std: Dict, report: GapReport):
        """比對技能項目（How 相關）"""
        user_text = _normalize_cjk((ui.how_skills + " " + ui.what_tasks).lower())
        skill_list = std.get("competency_skills", []) or std.get("skills", [])

        for s in skill_list:
            code = s.get("code", "")
            name = s.get("name", "")
            desc = s.get("description", "")
            if not name:
                continue
            if self._text_contains_any(user_text, name.lower().split()):
                report.covered_skills.append(name)
            else:
                report.gap_skills.append(GapItem(
                    category="skill",
                    code=code,
                    name=name,
                    description=desc,
                    severity="medium",
                ))

    def _analyze_behaviors(self, ui: UserInput5W2H, std: Dict, report: GapReport):
        """比對行為指標（How Much 相關）"""
        user_text = (ui.how_much_kpi + " " + ui.what_tasks).lower()
        tasks = std.get("competency_tasks", [])

        seen_behaviors = set()
        for task in tasks:
            for behavior in task.get("behaviors", []):
                if not isinstance(behavior, dict):
                    continue
                code = behavior.get("code", "")
                desc = behavior.get("description", "")
                if not desc or code in seen_behaviors:
                    continue
                seen_behaviors.add(code)
                # 行為指標通常較長，只取前50字比對
                short = desc[:50].lower()
                keywords = [w for w in short.split() if len(w) > 1]
                if not self._text_contains_any(user_text, keywords):
                    report.gap_behaviors.append(GapItem(
                        category="behavior",
                        code=code,
                        name=desc[:60] + ("..." if len(desc) > 60 else ""),
                        description=desc,
                        severity="low",
                    ))

    def _analyze_outputs(self, ui: UserInput5W2H, std: Dict, report: GapReport):
        """比對工作產出（What outputs）"""
        user_text = (ui.what_outputs + " " + ui.what_tasks).lower()
        tasks = std.get("competency_tasks", [])

        seen_outputs = set()
        for task in tasks:
            for output in task.get("output", []):
                if not isinstance(output, dict):
                    continue
                code = output.get("code", "")
                name = output.get("name", "")
                if not name or name in seen_outputs:
                    continue
                seen_outputs.add(name)
                if not self._text_contains_any(user_text, name.lower().split()):
                    report.gap_outputs.append(GapItem(
                        category="output",
                        code=code,
                        name=name,
                        severity="medium",
                    ))

    # ─── 輔助方法 ─────────────────────────────

    @staticmethod
    def _text_contains_any(text: str, keywords: List[str]) -> bool:
        """text 是否包含 keywords 中的任一詞。
        先做精確比對；中文詞再以字元 bigram/trigram 模糊比對，
        提升使用者未使用精確術語時的召回率。
        """
        meaningful = [k for k in keywords if len(k) > 1]
        if not meaningful:
            return False
        for kw in meaningful:
            if kw in text:
                return True
            # CJK 字元 n-gram 模糊比對（取長度 2-3 的子字串）
            cjk_chars = re.findall(r'[\u4e00-\u9fff\u3400-\u4dbf]', kw)
            if len(cjk_chars) >= 2:
                max_n = min(3, len(cjk_chars))
                for n in range(2, max_n + 1):
                    for i in range(len(cjk_chars) - n + 1):
                        if ''.join(cjk_chars[i:i + n]) in text:
                            return True
        return False

    @staticmethod
    def _task_severity(task: Dict) -> str:
        """根據任務層級判定嚴重度（相容 int 與 str 型別）"""
        try:
            level = int(task.get("level", 0))
        except (ValueError, TypeError):
            level = 0
        if level >= 4:
            return "high"
        elif level >= 2:
            return "medium"
        return "low"

    def _calc_completeness(self, report: GapReport, std: Dict) -> float:
        """計算整體完整度（0-100）"""
        scores = []

        # 任務完整度
        total_tasks = len(report.covered_tasks) + len(report.gap_tasks)
        if total_tasks > 0:
            scores.append(len(report.covered_tasks) / total_tasks)

        # 知識完整度
        total_k = len(report.covered_knowledge) + len(report.gap_knowledge)
        if total_k > 0:
            scores.append(len(report.covered_knowledge) / total_k)

        # 技能完整度
        total_s = len(report.covered_skills) + len(report.gap_skills)
        if total_s > 0:
            scores.append(len(report.covered_skills) / total_s)

        if not scores:
            return 0.0

        return round(sum(scores) / len(scores) * 100, 1)

    def get_summary_text(self, report: GapReport) -> str:
        """生成人類可讀的缺口摘要"""
        lines = []
        lines.append(f"【最佳匹配職能基準】")
        lines.append(f"  {report.best_standard_name}（{report.best_standard_code}）")
        lines.append(f"  整體完整度：{report.completeness_score}%")
        lines.append("")

        if report.gap_tasks:
            lines.append(f"【工作任務缺口】（{len(report.gap_tasks)} 項）")
            for g in report.gap_tasks[:5]:
                lines.append(f"  • [{g.code}] {g.name}")
            if len(report.gap_tasks) > 5:
                lines.append(f"  ... 等共 {len(report.gap_tasks)} 項")
            lines.append("")

        if report.gap_knowledge:
            lines.append(f"【知識缺口】（{len(report.gap_knowledge)} 項）")
            for g in report.gap_knowledge[:5]:
                lines.append(f"  • [{g.code}] {g.name}")
            if len(report.gap_knowledge) > 5:
                lines.append(f"  ... 等共 {len(report.gap_knowledge)} 項")
            lines.append("")

        if report.gap_skills:
            lines.append(f"【技能缺口】（{len(report.gap_skills)} 項）")
            for g in report.gap_skills[:5]:
                lines.append(f"  • [{g.code}] {g.name}")
            if len(report.gap_skills) > 5:
                lines.append(f"  ... 等共 {len(report.gap_skills)} 項")
            lines.append("")

        if not report.gap_tasks and not report.gap_knowledge and not report.gap_skills:
            lines.append("✓ 主要工作內容與職能基準高度吻合，可直接輸出職能說明書。")

        return "\n".join(lines)
