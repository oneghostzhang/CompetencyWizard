"""
ai_chat.py
LM Studio 對答式 AI 引導填寫模組
透過自然對話引導員工逐一描述工作任務，並整理成 5W2H 格式。
"""

import json
import re
from typing import Optional

# ── 對話系統提示 ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
你是一位親切的 HR 助理，正在透過對話幫助員工填寫「工作職能說明書」。

你的目標：蒐集員工的工作任務，每個任務需收集以下 5W2H 資訊：
  • What  — 任務名稱、主要工作內容、具體產出成果
  • Why   — 為什麼要做這個任務、目的和價值
  • Who   — 由誰執行、和哪些人或部門協作
  • When  — 何時執行、頻率（每日/每週/每月/不定期）
  • Where — 在哪裡執行、使用的系統或工作環境
  • How   — 如何完成、用什麼方法或技術工具
  • How much — 量化指標或 KPI

規則：
1. 每次只問一個問題，語氣親切自然，使用繁體中文。
2. 先請員工說出任務名稱，再依序收集各欄位。
3. 若員工回答不夠完整，可以追問一次，但不要重複追問同一問題。
4. 若員工說「不確定」、「沒有」、「不適用」，接受並繼續下一欄位。
5. 一個任務的 5W2H 都收集完後，詢問是否還有其他任務。
6. 當所有任務收集完畢後，告訴員工「我已整理完畢，請確認以下清單」，
   然後輸出以下格式（必須在回覆的最後，且完整輸出）：

[TASKS_JSON]
[
  {
    "what_tasks": "任務名稱",
    "what_outputs": "產出成果",
    "why_purpose": "目的",
    "who_role": "執行者角色",
    "who_collaborate": "協作對象",
    "when_frequency": "頻率",
    "where_environment": "環境/系統",
    "how_skills": "方法/技術",
    "how_much_kpi": "KPI"
  }
]
[/TASKS_JSON]

請先向員工問好，並詢問第一個工作任務的名稱。
"""

# ── LM Studio 模型清單 ────────────────────────────────────────────────────────

# 預設優先順序（名稱部分匹配即可）
PREFERRED_MODELS = [
    "TAIDE",
    "Qwen3",
    "Qwen2.5",
    "gemma",
]

DEFAULT_MODEL = "TAIDE-LX-7B-Chat"


# ── 工具函式 ─────────────────────────────────────────────────────────────────

def extract_tasks_json(text: str) -> Optional[list]:
    """從 AI 回應文字中提取 [TASKS_JSON]...[/TASKS_JSON] 區塊並解析。"""
    pattern = r"\[TASKS_JSON\](.*?)\[/TASKS_JSON\]"
    match = re.search(pattern, text, re.DOTALL)
    if not match:
        return None
    try:
        return json.loads(match.group(1).strip())
    except json.JSONDecodeError:
        return None


def strip_tasks_json(text: str) -> str:
    """移除回應中的 JSON 區塊，只留給使用者看的文字。"""
    return re.sub(r"\[TASKS_JSON\].*?\[/TASKS_JSON\]", "", text, flags=re.DOTALL).strip()


# ── 主要類別 ─────────────────────────────────────────────────────────────────

class LMStudioChat:
    """管理與 LM Studio 的對話歷史和 API 呼叫。"""

    def __init__(self, model: str = DEFAULT_MODEL):
        self.model = model
        self.history: list[dict] = [
            {"role": "system", "content": SYSTEM_PROMPT}
        ]
        self.extracted_tasks: Optional[list] = None
        self._client = None

    # ── 私有方法 ──────────────────────────────────────────────────────────────

    def _get_client(self):
        if self._client is None:
            try:
                from openai import OpenAI
            except ImportError:
                raise ImportError("請先安裝 openai 套件：pip install openai")
            self._client = OpenAI(
                base_url="http://localhost:1234/v1",
                api_key="lm-studio",
            )
        return self._client

    def _call(self, max_tokens: int = 1024) -> str:
        client = self._get_client()
        response = client.chat.completions.create(
            model=self.model,
            messages=self.history,
            temperature=0.7,
            max_tokens=max_tokens,
        )
        reply = response.choices[0].message.content or ""
        self.history.append({"role": "assistant", "content": reply})

        tasks = extract_tasks_json(reply)
        if tasks is not None:
            self.extracted_tasks = tasks

        return reply

    # ── 公開方法 ──────────────────────────────────────────────────────────────

    def start(self) -> str:
        """取得 AI 的第一句問候語（不需要使用者輸入）。"""
        return self._call(max_tokens=512)

    def send(self, user_message: str) -> str:
        """送出使用者訊息，取得 AI 回應。"""
        self.history.append({"role": "user", "content": user_message})
        return self._call(max_tokens=1024)

    def is_done(self) -> bool:
        """AI 是否已輸出完整任務清單。"""
        return self.extracted_tasks is not None

    def get_tasks(self) -> list[dict]:
        """取得解析後的任務清單（僅在 is_done() 為 True 後有效）。"""
        return self.extracted_tasks or []

    # ── 靜態工具 ──────────────────────────────────────────────────────────────

    @staticmethod
    def get_available_models() -> list[str]:
        """取得 LM Studio 目前載入的模型清單。"""
        try:
            from openai import OpenAI
            client = OpenAI(base_url="http://localhost:1234/v1", api_key="lm-studio")
            models = client.models.list()
            return [m.id for m in models.data]
        except Exception:
            return []

    @staticmethod
    def check_server() -> bool:
        """檢查 LM Studio Server 是否正在運作。"""
        try:
            import urllib.request
            urllib.request.urlopen("http://localhost:1234/v1/models", timeout=2)
            return True
        except Exception:
            return False
