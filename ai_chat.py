"""
ai_chat.py  v2.1
職能說明書 AI 對話引導模組

後端優先順序：
  1. LlamaCpp（直接載入 GGUF，無 HTTP timeout 問題）
  2. LM Studio REST API（fallback，需開啟 LM Studio Server）

對話分五個階段：
  Phase 1 — 基本資訊收集
  Phase 2 — ICAP 職能基準確認（由系統注入候選標準）
  Phase 3 — 主要職責清單確認
  Phase 4 — 逐職責深度訪談（子任務 × 產出 × 行為指標 × 知識技能）
  Phase 5 — 輸出完整職能說明書 JSON
"""

import json
import logging
import re
import tomllib
from pathlib import Path
from typing import Optional

from openai.types.chat import ChatCompletionMessageParam

logger = logging.getLogger(__name__)

# ── 從 config.toml 讀取設定 ──────────────────────────────────────────────────

def _load_config() -> dict:
    cfg_path = Path(__file__).parent / "config.toml"
    if cfg_path.exists():
        with open(cfg_path, "rb") as f:
            return tomllib.load(f)
    return {}

_cfg = _load_config()

TAIDE_MODEL_PATH: str = _cfg.get("model", {}).get(
    "taide_path",
    r"C:\Users\User\.lmstudio\models\ZoneTwelve"
    r"\TAIDE-LX-7B-Chat-GGUF\TAIDE-LX-7B-Chat.Q4_K_S.gguf",
)

# ── LLM 參數 ─────────────────────────────────────────────────────────────────

_llm_cfg     = _cfg.get("llm", {})
N_CTX        = _llm_cfg.get("n_ctx",       4096)
N_THREADS    = _llm_cfg.get("n_threads",   8)
TEMPERATURE  = _llm_cfg.get("temperature", 0.3)
MAX_TOKENS   = _llm_cfg.get("max_tokens",  512)
STOP_TOKENS  = ["\n使用者:", "\n員工:", "\n問題:", "使用者：", "員工："]

DEFAULT_MODEL = "taide-lx-7b-chat"

_MAX_HISTORY_TURNS = 20

# ── 固定開場白 ────────────────────────────────────────────────────────────────

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

# ── 系統提示（精簡版，降低 token 消耗）──────────────────────────────────────

SYSTEM_PROMPT = """你是親切的HR助理，用繁體中文引導員工完成職能說明書。依序進行以下5個步驟，每次只問一個問題：

步驟1【基本資訊】依序詢問：職位名稱→公司/部門（可略）→工作內容描述→職能等級(1-5,不知道填3)。收集完說「謝謝，我將為您比對職能基準」。

步驟2【確認基準】若收到【STANDARD_DATA】，向員工說明找到的職能基準名稱與工作描述，問是否符合。符合→步驟3；不符合→記錄差異後仍進步驟3。

步驟3【確認職責】列出標準的主要職責，請員工對照增刪修改，用表格格式呈現：
| 代碼 | 主要職責名稱 |
|------|------------|
| T1 | 職責名稱 |
確認清單後進步驟4。

步驟4【逐項訪談】針對每個主要職責依序：
  A. 問員工描述這個職責實際做什麼
  B. 整理子任務，用表格請員工確認：| 子任務 | 說明 | 例：| T1.1 | 說明 |
  C. 問産出成果、成效例子、知識、技能，收集完用表格摘要：
     | 項目 | 內容 |
     |------|------|
     | 工作產出 | ... |
     | 行為指標 | ... |
     | 知識 | ... |
     | 技能 | ... |
  全部職責完成後進步驟5。

步驟5【輸出】說「職能說明書已整理完成，請系統輸出Excel。」然後輸出：

[COMPETENCY_JSON]
{"basic_info":{"position":"","company":"","department":"","description":"","level":3,"matched_standard_code":"","matched_standard_name":""},"main_responsibilities":[{"code":"T1","name":"","tasks":[{"code":"T1.1","name":"","output_code":"O1.1.1","output":"","behavior_indicator":"","level":3,"knowledge":[],"skills":[]}]}]}
[/COMPETENCY_JSON]

規則：員工說不確定/沒有→空白繼續；摘要用表格；不向員工說明步驟編號。"""

# ── 工具函式 ─────────────────────────────────────────────────────────────────

def extract_competency_json(text: str) -> Optional[dict]:
    match = re.search(r"\[COMPETENCY_JSON\](.*?)\[/COMPETENCY_JSON\]", text, re.DOTALL)
    if not match:
        return None
    try:
        return json.loads(match.group(1).strip())
    except json.JSONDecodeError:
        return None


def strip_output_json(text: str) -> str:
    return re.sub(
        r"\[COMPETENCY_JSON\].*?\[/COMPETENCY_JSON\]",
        "", text, flags=re.DOTALL
    ).strip()


def competency_to_task_list(competency: dict) -> list[dict]:
    """將職能 JSON 轉為 _added_tasks 5W2H 格式，供現有表單匯入。"""
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


# ── 後端：LlamaCpp（優先）────────────────────────────────────────────────────

class _LlamaCppBackend:
    """直接載入 GGUF，無 HTTP timeout，借鑑 Graph_RAG_test 參數設定。"""

    def __init__(self, model_path: str):
        from langchain_community.llms import LlamaCpp
        logger.info("載入 LlamaCpp 模型：%s", model_path)
        self._llm = LlamaCpp(
            model_path=model_path,
            n_ctx=N_CTX,
            n_threads=N_THREADS,
            temperature=TEMPERATURE,
            max_tokens=MAX_TOKENS,
            verbose=False,
            stop=STOP_TOKENS,
        )
        logger.info("LlamaCpp 載入完成")

    def chat(self, messages: list[ChatCompletionMessageParam]) -> str:
        """將對話歷史轉為 TAIDE chat template 格式後推論。"""
        prompt = _build_taide_prompt(messages)
        result = self._llm.invoke(prompt)
        if isinstance(result, str):
            return result.strip()
        if hasattr(result, "content"):
            return result.content.strip()
        return str(result).strip()


def _build_taide_prompt(messages: list[ChatCompletionMessageParam]) -> str:
    """
    將 OpenAI 格式的 messages 轉為 TAIDE llama chat template。
    TAIDE 使用 [INST] ... [/INST] 格式（和 Llama 2 相容）。
    """
    parts = []
    system_content = ""

    for msg in messages:
        role = msg["role"]
        content = msg.get("content", "")
        if not isinstance(content, str):
            content = str(content)
        if role == "system":
            system_content = content
        elif role == "user":
            if system_content:
                parts.append(f"[INST] <<SYS>>\n{system_content}\n<</SYS>>\n\n{content} [/INST]")
                system_content = ""  # system 只附加一次
            else:
                parts.append(f"[INST] {content} [/INST]")
        elif role == "assistant":
            parts.append(f"{content}")

    return "".join(parts)


# ── 後端：LM Studio REST API（fallback）─────────────────────────────────────

class _LMStudioBackend:
    """透過 LM Studio OpenAI 相容 API 呼叫（需開啟 Server）。"""

    def __init__(self, model: str = DEFAULT_MODEL):
        from openai import OpenAI
        self._client = OpenAI(base_url="http://localhost:1234/v1", api_key="lm-studio")
        self._model = model

    def chat(self, messages: list[ChatCompletionMessageParam]) -> str:
        resp = self._client.chat.completions.create(
            model=self._model,
            messages=messages,
            temperature=TEMPERATURE,
            max_tokens=MAX_TOKENS,
            timeout=300,
        )
        return (resp.choices[0].message.content or "").strip()


# ── 主要類別 ─────────────────────────────────────────────────────────────────

class LMStudioChat:
    """
    管理職能說明書 5 階段對話。
    後端自動選擇：LlamaCpp（優先）→ LM Studio API（fallback）。
    """

    def __init__(self, model: str = DEFAULT_MODEL, model_path: str = TAIDE_MODEL_PATH):
        self._model      = model
        self._model_path = model_path
        self._backend    = None          # 延遲初始化（在 ChatWorker 執行緒中載入）
        self.history: list[ChatCompletionMessageParam] = [
            {"role": "system", "content": SYSTEM_PROMPT}
        ]
        self._competency: Optional[dict] = None

    # ── 後端初始化（在背景執行緒呼叫）─────────────────────────────────────

    def init_backend(self) -> str:
        """
        初始化推論後端，回傳後端名稱。
        優先 LlamaCpp；若模型檔不存在或套件未安裝則 fallback LM Studio。
        """
        if self._backend is not None:
            return "already_init"

        # 嘗試 LlamaCpp
        if Path(self._model_path).exists():
            try:
                self._backend = _LlamaCppBackend(self._model_path)
                return "llamacpp"
            except Exception as e:
                logger.warning("LlamaCpp 初始化失敗，改用 LM Studio：%s", e)

        # Fallback：LM Studio REST API
        try:
            self._backend = _LMStudioBackend(self._model)
            return "lmstudio"
        except Exception as e:
            raise RuntimeError(f"所有後端均初始化失敗：{e}")

    # ── 私有 ────────────────────────────────────────────────────────────────

    def _trim_history(self) -> None:
        system = [m for m in self.history if m["role"] == "system"]
        rest   = [m for m in self.history if m["role"] != "system"]
        max_msgs = _MAX_HISTORY_TURNS * 2
        if len(rest) > max_msgs:
            rest = rest[-max_msgs:]
        self.history = system + rest

    def _call(self) -> str:
        if self._backend is None:
            raise RuntimeError("後端尚未初始化，請先呼叫 init_backend()")
        self._trim_history()
        reply = self._backend.chat(self.history)
        self.history.append({"role": "assistant", "content": reply})
        parsed = extract_competency_json(reply)
        if parsed is not None:
            self._competency = parsed
        return reply

    # ── 公開 API ────────────────────────────────────────────────────────────

    def start(self) -> str:
        """回傳固定開場白（不呼叫 LLM）。"""
        self.history.append({"role": "assistant", "content": GREETING})
        return GREETING

    def send(self, user_message: str) -> str:
        """送出員工訊息，取得 AI 回應。"""
        self.history.append({"role": "user", "content": user_message})
        return self._call()

    def inject_standard(self, standard_data: dict) -> str:
        """Phase 2：注入 ICAP 職能基準資料，AI 向員工介紹並確認。"""
        meta      = standard_data.get("metadata", {})
        bi        = standard_data.get("basic_info", {})
        tasks_raw = standard_data.get("competency_tasks", [])
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
            f"主要職責：\n{resp_list}"
        )
        self.history.append({"role": "system", "content": inject_text})
        self.history.append({
            "role": "user",
            "content": "（系統已找到可能符合的職能基準，請向我介紹並確認是否符合）"
        })
        return self._call()

    def is_done(self) -> bool:
        return self._competency is not None

    def get_competency(self) -> dict:
        return self._competency or {}

    def get_tasks_for_import(self) -> list[dict]:
        return competency_to_task_list(self._competency or {})

    # ── 靜態工具 ────────────────────────────────────────────────────────────

    @staticmethod
    def check_server() -> bool:
        """檢查 LM Studio Server（fallback 路徑用）。"""
        import socket
        try:
            with socket.create_connection(("127.0.0.1", 1234), timeout=2):
                return True
        except Exception:
            return False

    @staticmethod
    def llamacpp_available() -> bool:
        """檢查 LlamaCpp 和 GGUF 模型是否都可用。"""
        if not Path(TAIDE_MODEL_PATH).exists():
            return False
        try:
            from langchain_community.llms import LlamaCpp  # noqa
            return True
        except ImportError:
            return False


# ── 模組級別行為指標分析（子 process 隔離，防止 llama.cpp abort 崩潰）────────

_WORKER_TIMEOUT = 120   # 每個任務最長等待秒數


def _build_prompt_messages(
    position: str,
    task_name: str,
    user_description: str,
    standard_behaviors: list,
) -> list[ChatCompletionMessageParam]:
    """組裝 system/user prompt messages（供 worker process 使用）。"""
    std_lines = []
    for b in standard_behaviors[:5]:
        if isinstance(b, dict):
            std_lines.append(b.get("description", ""))
        elif isinstance(b, str):
            std_lines.append(b)
    std_text = "\n".join(f"- {l}" for l in std_lines if l) or "（無標準行為指標）"
    user_desc = user_description.strip() or "（員工未填寫）"

    system_prompt = (
        "你是職能說明書專家，使用繁體中文。\n"
        "請根據員工描述與參考標準，生成 2～3 條行為指標。\n"
        "【格式規定】\n"
        "・每條限 1 句，30～60 字，以行動動詞開頭（依據/根據/核對/確認/整理/協助/記錄/編製/彙整/執行）\n"
        "・描述「做什麼」與「達到什麼結果」，不用「能夠」「在...的角色下」等虛詞\n"
        "・風格參考：「依據財會法規，處理各類收支傳票並登錄於系統，經主管簽核後妥善歸檔。」\n"
        "・嚴禁輸出解釋、標題、編號、多餘文字，只輸出 JSON。"
    )
    user_prompt = (
        f"職位：{position}\n"
        f"工作任務：{task_name}\n"
        f"員工實際描述：{user_desc}\n"
        f"參考標準行為指標（僅供風格參考，勿照抄）：\n{std_text}\n\n"
        '只輸出 JSON，格式：{"behavior_indicators":["指標1","指標2","指標3"]}'
    )
    return [
        {"role": "system", "content": system_prompt},
        {"role": "user",   "content": user_prompt},
    ]


def _worker_main(tasks: list, q):
    """
    子 process 入口：依序處理所有任務，每完成一個往 Queue 放 (idx, indicators)。
    結束後放 None 作為 sentinel。llama.cpp abort 只殺死此 process。
    """
    import re as _re, json as _json
    from pathlib import Path as _Path

    # 建立後端（子 process 內獨立初始化）
    backend = None
    if _Path(TAIDE_MODEL_PATH).exists():
        try:
            backend = _LlamaCppBackend(TAIDE_MODEL_PATH)
        except Exception:
            pass
    if backend is None:
        backend = _LMStudioBackend()

    for idx, task_args in tasks:
        try:
            messages = _build_prompt_messages(**task_args)
            reply = backend.chat(messages)
            match = _re.search(r'\{.*?"behavior_indicators".*?\}', reply, _re.DOTALL)
            if match:
                data = _json.loads(match.group())
                indicators = data.get("behavior_indicators", [])
                if isinstance(indicators, list):
                    q.put((idx, _split_indicators(indicators)))
                    continue
            lines = [l.strip().lstrip("-•・ ")
                     for l in reply.split("\n") if len(l.strip()) > 8]
            q.put((idx, _split_indicators(lines[:6])))
        except Exception as e:
            q.put((idx, []))
    q.put(None)   # sentinel


def analyze_task(
    position: str,
    task_name: str,
    user_description: str,
    standard_behaviors: list,
    backend=None,
) -> dict:
    """
    單次 LLM 呼叫（子 process 隔離版）：生成 ICAP 格式行為指標。
    llama.cpp abort 時子 process 崩潰，主程式不受影響。
    """
    import multiprocessing as _mp
    task_args = dict(position=position, task_name=task_name,
                     user_description=user_description,
                     standard_behaviors=standard_behaviors)
    q = _mp.Queue()
    p = _mp.Process(target=_worker_main, args=([(0, task_args)], q), daemon=True)
    p.start()
    try:
        item = q.get(timeout=_WORKER_TIMEOUT)
        if item is None:
            return {"behavior_indicators": [], "error": "worker 無回應"}
        _, indicators = item
        logger.info("analyze_task 完成：%s", task_name)
        return {"behavior_indicators": indicators, "error": None}
    except Exception as e:
        logger.error("analyze_task 失敗：%s", e)
        return {"behavior_indicators": [], "error": str(e)}
    finally:
        p.kill()


def analyze_tasks_batch(
    rows: list,
    position: str,
    result_cb,
    done_cb,
    error_cb,
):
    """
    批次分析（子 process 隔離版）：一次啟動一個子 process 處理所有任務。
    每完成一個任務呼叫 result_cb(idx, indicators)；
    全部完成呼叫 done_cb()；子 process 意外崩潰時呼叫 error_cb(msg)。

    回傳 (Process, Queue)，呼叫端可監控。
    """
    import multiprocessing as _mp
    tasks = [
        (i, dict(
            position=position,
            task_name=row.get("task_name", ""),
            user_description=row.get("user_description", ""),
            standard_behaviors=row.get("_behaviors", []),
        ))
        for i, row in enumerate(rows)
    ]
    q = _mp.Queue()
    p = _mp.Process(target=_worker_main, args=(tasks, q), daemon=True)
    p.start()
    return p, q


def _split_indicators(raw: list) -> list:
    """將 LLM 可能合併成單一字串的多條指標拆開，去除「指標N:」等前綴，限回傳 3 條。"""
    result = []
    for item in raw:
        item = str(item).strip()
        if not item:
            continue
        # 若含換行或「指標N:」模式，視為多條合併，拆開
        parts = re.split(r'\n|(?:指標\s*\d+\s*[:：])', item)
        for p in parts:
            p = p.strip().lstrip("-•・ ")
            # 去除行首的「N.」「N、」「N)」等編號
            p = re.sub(r'^\d+[\.、\)）]\s*', '', p)
            if len(p) > 5:
                result.append(p)
    return result[:3]
