"""
ai_chat.py  v2.0
職能說明書 AI 對話引導模組

對話分五個階段：
  Phase 1 — 基本資訊收集（職位、公司、工作描述）
  Phase 2 — ICAP 職能基準確認（由系統注入候選標準）
  Phase 3 — 主要職責清單確認（T 代碼確認 / 增刪）
  Phase 4 — 逐職責深度訪談（子任務 × 工作產出 × 行為指標 × 知識技能）
  Phase 5 — 輸出完整職能說明書 JSON

輸出格式對齊 AI詠唱人員職能基準.xlsx：
  Sheet 2 欄位：主要職責 / 工作任務 / 工作產出 / 行為指標 / 職能級別 / 知識（K）/ 技能（S）
"""

import json
import logging
import re
from typing import Optional

logger = logging.getLogger(__name__)

# ── 常數 ─────────────────────────────────────────────────────────────────────

DEFAULT_MODEL       = "taide-lx-7b-chat"
_MAX_HISTORY_TURNS  = 30   # 最多保留 30 輪，避免 context 爆掉

# ── 固定開場白（不呼叫 API，瞬間顯示）──────────────────────────────────────

GREETING = """\
您好！我是您的職能說明書填寫助理。

我會透過幾個步驟，幫您建立一份符合 ICAP 標準的職能說明書：

  步驟 1 ► 收集基本資訊（職位、公司、工作描述）
  步驟 2 ► 比對 ICAP 職能基準，確認最接近的標準
  步驟 3 ► 確認您的主要職責清單
  步驟 4 ► 逐項深入了解各職責的工作任務與產出
  步驟 5 ► 整理完成，輸出職能說明書

請問您的職位名稱是什麼？（例如：AI 詠唱人員、行銷專員、電腦維修工讀生）\
"""

# ── 系統提示（5 階段狀態機）─────────────────────────────────────────────────

SYSTEM_PROMPT = """\
你是一位親切專業的 HR 助理，任務是透過對話引導員工完成「職能說明書」的建立。
整份對話分為以下五個階段，你必須按序進行，不可跳過：

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【Phase 1】基本資訊收集
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
依序收集（每次只問一個）：
  1. 職位名稱（必問）
  2. 公司 / 組織名稱（可省略）
  3. 所屬部門（可省略）
  4. 用自己的話描述工作內容（必問，鼓勵詳細描述）
  5. 目前的職能等級（1–5，不知道填 3）

收集完畢後說：「感謝您的描述，我會根據這些資訊為您找到最合適的職能基準。」
等待系統注入職能基準資料後進入 Phase 2。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【Phase 2】ICAP 職能基準確認
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
系統會注入一段【STANDARD_DATA】資料。
你的任務：
  • 向員工說明找到的職能基準（名稱、代碼、工作描述）
  • 詢問：「這個職能基準與您的工作描述是否相符？」
  • 員工確認 → 進入 Phase 3
  • 員工否認 → 詢問差異，記錄後仍進入 Phase 3（備注差異）

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【Phase 3】主要職責清單確認
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
從注入的標準資料中列出主要職責（T1、T2...），例如：

  T1 需求分析
  T2 協助建置 AI 模型運行環境
  T3 訓練 AI 模型
  T4 協助測試及優化 AI 模型

詢問員工：
  「以上是標準的主要職責，請用同樣格式告訴我您實際的職責清單。
   您可以新增、刪除或修改任何項目。」

確認最終職責清單後，說明接下來會逐項深入討論，進入 Phase 4。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【Phase 4】逐職責深度訪談
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
針對每一個主要職責（T1、T2...），依序進行：

Step A — 開放式描述
  「請用您自己的話，描述「Tx 職責名稱」這個職責您實際在做什麼？」

Step B — 拆解子任務
  根據員工描述，整理出 2–4 個子任務（Tx.1、Tx.2...），請員工確認。

Step C — 逐子任務收集細節（每個子任務問以下 4 項）：
  1. 工作產出（Output）：「這個任務完成後，具體的成果或交付物是什麼？」
  2. 行為指標：「您怎麼知道這個任務做得好？請描述一個實際例子。」
     （引導使用 STAR：情境→任務→行動→結果）
  3. 需要的知識（K）：「完成這個任務需要具備哪些知識或理論基礎？」
  4. 需要的技能（S）：「需要哪些實際操作技能或工具使用能力？」

每個子任務完成後：「好的，我已記錄完 Tx.x，接下來討論 Tx.x+1。」
整個職責完成後：「Tx 已完成，接下來討論 Tx+1。」
所有職責完成後，進入 Phase 5。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【Phase 5】輸出職能說明書
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
所有職責訪談完成後，說：
「感謝您的配合！我已整理完成您的職能說明書，請稍候系統輸出 Excel 檔案。」

然後輸出以下格式（必須完整，放在回覆最末尾）：

[COMPETENCY_JSON]
{
  "basic_info": {
    "position": "職位名稱",
    "company": "公司名稱",
    "department": "部門",
    "description": "工作描述",
    "level": 3,
    "matched_standard_code": "SET4130-002v1",
    "matched_standard_name": "AI 詠唱人員"
  },
  "main_responsibilities": [
    {
      "code": "T1",
      "name": "主要職責名稱",
      "tasks": [
        {
          "code": "T1.1",
          "name": "子任務名稱",
          "output_code": "O1.1.1",
          "output": "工作產出",
          "behavior_indicator": "行為指標（STAR 格式）",
          "level": 3,
          "knowledge": ["知識項目1", "知識項目2"],
          "skills": ["技能項目1", "技能項目2"]
        }
      ]
    }
  ]
}
[/COMPETENCY_JSON]

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【通用規則】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• 全程使用繁體中文
• 每次只問一個問題，語氣親切自然
• 員工說「不確定 / 沒有 / 不適用」→ 記錄為空，繼續下一項
• 每個 Phase 結束前先摘要確認，再進入下一 Phase
• 不要一次問多個問題
• 不要向員工說明「Phase 幾」這種系統術語
\
"""

# ── 工具函式 ─────────────────────────────────────────────────────────────────

def extract_competency_json(text: str) -> Optional[dict]:
    """提取 [COMPETENCY_JSON]...[/COMPETENCY_JSON] 並解析為 dict。"""
    match = re.search(r"\[COMPETENCY_JSON\](.*?)\[/COMPETENCY_JSON\]", text, re.DOTALL)
    if not match:
        return None
    try:
        return json.loads(match.group(1).strip())
    except json.JSONDecodeError:
        return None


def strip_output_json(text: str) -> str:
    """移除 AI 回應中的 JSON 區塊，只保留給使用者看的對話文字。"""
    return re.sub(
        r"\[COMPETENCY_JSON\].*?\[/COMPETENCY_JSON\]",
        "", text, flags=re.DOTALL
    ).strip()


def competency_to_task_list(competency: dict) -> list[dict]:
    """
    將職能 JSON 展開為 _added_tasks 格式（5W2H dict），
    以便與現有表單系統相容。
    """
    tasks = []
    for resp in competency.get("main_responsibilities", []):
        for t in resp.get("tasks", []):
            tasks.append({
                "what_tasks":        f"{t.get('code','')} {t.get('name','')}".strip(),
                "what_outputs":      t.get("output", ""),
                "why_purpose":       t.get("behavior_indicator", ""),
                "who_role":          competency.get("basic_info", {}).get("position", ""),
                "who_collaborate":   "",
                "when_frequency":    [],
                "where_environment": "",
                "how_skills":        "\n".join(t.get("skills", [])),
                "how_much_kpi":      "",
            })
    return tasks


# ── 主要類別 ─────────────────────────────────────────────────────────────────

class LMStudioChat:
    """
    管理與 LM Studio 的 5 階段對話。

    典型使用流程：
      session = LMStudioChat()
      greeting = session.start()           # 顯示固定開場白
      reply = session.send(user_text)      # 逐輪對話
      session.inject_standard(std_data)    # Phase 2：注入 ICAP 基準
      ...
      if session.is_done():
          data = session.get_competency()  # 取得完整職能 JSON
    """

    def __init__(self, model: str = DEFAULT_MODEL):
        self.model   = model
        self.history: list[dict] = [
            {"role": "system", "content": SYSTEM_PROMPT}
        ]
        self._competency: Optional[dict] = None
        self._client = None

    # ── 私有 ────────────────────────────────────────────────────────────────

    def _get_client(self):
        if self._client is None:
            try:
                from openai import OpenAI
            except ImportError:
                raise ImportError("請安裝 openai 套件：pip install openai")
            self._client = OpenAI(
                base_url="http://localhost:1234/v1",
                api_key="lm-studio",
            )
        return self._client

    def _trim_history(self) -> None:
        system = [m for m in self.history if m["role"] == "system"]
        rest   = [m for m in self.history if m["role"] != "system"]
        max_msgs = _MAX_HISTORY_TURNS * 2
        if len(rest) > max_msgs:
            rest = rest[-max_msgs:]
        self.history = system + rest

    def _call(self, max_tokens: int = 1024) -> str:
        self._trim_history()
        client = self._get_client()
        resp = client.chat.completions.create(
            model=self.model,
            messages=self.history,
            temperature=0.65,
            max_tokens=max_tokens,
            timeout=120,
        )
        reply = resp.choices[0].message.content or ""
        self.history.append({"role": "assistant", "content": reply})

        parsed = extract_competency_json(reply)
        if parsed is not None:
            self._competency = parsed

        return reply

    # ── 公開 ────────────────────────────────────────────────────────────────

    def start(self) -> str:
        """回傳固定開場白並寫入歷史（不呼叫 API）。"""
        self.history.append({"role": "assistant", "content": GREETING})
        return GREETING

    def send(self, user_message: str) -> str:
        """送出員工訊息，取得 AI 回應。"""
        self.history.append({"role": "user", "content": user_message})
        return self._call()

    def inject_standard(self, standard_data: dict) -> str:
        """
        Phase 2：將 ICAP 職能基準資料注入對話，
        讓 AI 向員工介紹並請求確認。
        standard_data 格式同 WizardRAG.get_standard() 的回傳值。
        """
        meta  = standard_data.get("metadata", {})
        bi    = standard_data.get("basic_info", {})
        tasks_raw = standard_data.get("competency_tasks", [])

        # 整理主要職責清單
        resp_list = "\n".join(
            f"  {t.get('task_id','')} {t.get('task_name','')}"
            for t in tasks_raw if t.get("task_name")
        )

        inject_text = (
            f"【STANDARD_DATA】\n"
            f"職能基準名稱：{meta.get('name','')}\n"
            f"代碼：{meta.get('code','')}\n"
            f"基準級別：{meta.get('level','')}\n"
            f"工作描述：{bi.get('job_description','')}\n"
            f"主要職責清單：\n{resp_list}"
        )

        # 以 system 角色注入（員工看不到，只有 AI 看到）
        self.history.append({"role": "system", "content": inject_text})
        # 觸發 AI 向員工說明
        self.history.append({
            "role": "user",
            "content": "（系統已找到可能符合的職能基準，請向我介紹並確認是否符合）"
        })
        return self._call(max_tokens=512)

    def is_done(self) -> bool:
        """AI 是否已輸出完整職能說明書 JSON。"""
        return self._competency is not None

    def get_competency(self) -> dict:
        """取得完整職能說明書 dict（Phase 5 後才有效）。"""
        return self._competency or {}

    def get_tasks_for_import(self) -> list[dict]:
        """將職能 JSON 轉為 _added_tasks 格式，供現有表單匯入。"""
        return competency_to_task_list(self._competency or {})

    # ── 靜態工具 ────────────────────────────────────────────────────────────

    @staticmethod
    def check_server() -> bool:
        """TCP socket 測試 LM Studio Server 是否啟動。"""
        import socket
        try:
            with socket.create_connection(("127.0.0.1", 1234), timeout=2):
                return True
        except Exception:
            return False

    @staticmethod
    def get_available_models() -> list[str]:
        """取得 LM Studio 目前可用模型清單。"""
        try:
            from openai import OpenAI
            client = OpenAI(base_url="http://localhost:1234/v1", api_key="lm-studio")
            return [m.id for m in client.models.list().data]
        except Exception as e:
            logger.warning("無法取得模型清單：%s", e)
            return []
