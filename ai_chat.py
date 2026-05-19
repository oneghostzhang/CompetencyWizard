"""
ai_chat.py  v2.3
職能說明書行為指標生成模組

後端優先順序：
  1. LlamaCpp（直接載入 GGUF，無 HTTP timeout 問題）
  2. LM Studio REST API（fallback，需開啟 LM Studio Server）

主要進入點：create_persistent_worker()
  建立長駐子 process，模型只載入一次。
  AUTO 模式採兩步式推論：先分類框架，再以選定框架生成指標。
"""

import logging
import re
import tomllib

# LLM 說明前言黑名單（fallback 逐行解析時排除這類 meta 文字）
_PREAMBLE_RE = re.compile(
    r'以下是|以下為|以下提供|行為指標如下|此.*格式|符合要求|不含其他|多餘字詞'
    r'|僅包含行為|指標如下|以下.*指標|根據.*描述.*以下|如下所示|以下幾點'
    r'|以下三|以下兩|以下為您'
)
from pathlib import Path

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
    "",   # 未設定時為空字串，程式啟動時自動 fallback 至 LM Studio API
)

# ── LLM 參數 ─────────────────────────────────────────────────────────────────

_llm_cfg     = _cfg.get("llm", {})
N_CTX        = _llm_cfg.get("n_ctx",       4096)
N_THREADS    = _llm_cfg.get("n_threads",   8)
TEMPERATURE  = _llm_cfg.get("temperature", 0.3)
MAX_TOKENS   = _llm_cfg.get("max_tokens",  512)
STOP_TOKENS  = ["\n使用者:", "\n員工:", "\n問題:", "使用者：", "員工："]

DEFAULT_MODEL = "taide-lx-7b-chat"

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


# ── 職能等級差異化提示 ────────────────────────────────────────────────────────
_LEVEL_HINT: dict[int, str] = {
    1: "初階（依指示執行，著重『依規定/依指示完成』）",
    2: "基礎（獨立完成標準作業，著重『依程序獨立處理』）",
    3: "中階（處理例外狀況，著重『判斷並調整』）",
    4: "進階（優化流程或指導他人，著重『改善/培訓』）",
    5: "專家（制定標準或跨部門協調，著重『規劃/制度建立』）",
}

# ── 三種分析模板 Prompt ───────────────────────────────────────────────────────
PROMPT_TEMPLATES: dict[str, dict[str, str]] = {

    # ── 5W2H：固定 SOP、重複性操作、有明確頻率要求的工作 ────────────────────
    "5W2H": {
        "system": (
            "你是 ICAP 職能說明書專家，使用繁體中文。\n"
            "請根據員工描述，以 5W2H 框架生成 2～3 條行為指標。\n\n"
            "【格式規定】每條為 1 句完整句子，以行動動詞開頭，"
            "將執行方式、使用工具/系統、執行頻率或時機、可驗證的結果整合在同一句中。\n\n"
            "【格式示意（勿照抄，需依員工描述生成）】\n"
            "  ✓「每[頻率]於[系統/工具]執行[操作]，確保[可驗證結果]。」\n"
            "  ✓「[時機]依[流程]完成[任務]，並於[時限]內提交[產出]。」\n\n"
            "禁止輸出解釋或多餘文字。"
        ),
        "user": (
            "職位：{position}｜職能等級：{level} — {level_hint}\n"
            "工作任務：{task_name}\n\n"
            "【員工描述（如何執行）】\n{user_desc}\n\n"
            "【工作產出/成果】\n{user_output}\n\n"
            "【任務相關背景知識（了解任務範疇用，輸出格式無需仿照）】\n{std_text}\n\n"
            "請依員工描述生成 5W2H 格式行為指標。\n"
            '只輸出 JSON，格式：{{"behavior_indicators":["指標1","指標2","指標3"]}}'
        ),
    },

    # ── ABCD：有明確 KPI 或驗收條件的工作 ──────────────────────────────────
    "ABCD": {
        "system": (
            "你是 ICAP 職能說明書專家，使用繁體中文。\n"
            "請根據員工描述，以 ABCD 框架生成 2～3 條行為指標。\n\n"
            "【格式規定】每條限 1 句，須同時包含以下三個要素：\n"
            "  ・C（Condition）：執行條件或依據，如「依規定」「在系統中」\n"
            "  ・B（Behavior）：具體行動動詞，描述做什麼\n"
            "  ・D（Degree）：可衡量或可驗證的達成標準，有具體數字時優先使用；"
            "若描述中無數字，改用可觀察的完成條件（如「無誤差」「按時完成」）\n"
            "建議句型：C，執行 B，達成 D。\n\n"
            "【格式示意（勿照抄，需依員工描述生成）】\n"
            "  ✓「依[條件/規定]（C），[動詞+操作]（B），達成[量化或可驗證標準]（D）。」\n"
            "  ✓「在[執行環境]（C），[動詞+操作]（B），確保[成果條件]（D）。」\n\n"
            "禁止輸出解釋或多餘文字。"
        ),
        "user": (
            "職位：{position}｜職能等級：{level} — {level_hint}\n"
            "工作任務：{task_name}\n\n"
            "【員工描述（如何執行）】\n{user_desc}\n\n"
            "【工作產出/成果】\n{user_output}\n\n"
            "【任務相關背景知識（了解任務範疇用，輸出格式無需仿照）】\n{std_text}\n\n"
            "請從員工描述中識別不同的執行行為，每個行為各自套用 C+B+D 結構生成獨立的一條指標。\n"
            "每條指標須為陣列的獨立元素，勿將多個行為合併在同一條。\n"
            '只輸出 JSON，格式：{{"behavior_indicators":["第一條指標","第二條指標","第三條指標"]}}'
        ),
    },

    # ── STAR：非例行任務、情境判斷、跨部門協調 ──────────────────────────────
    "STAR": {
        "system": (
            "你是 ICAP 職能說明書專家，使用繁體中文。\n"
            "請根據員工描述，以 STAR 框架生成 2～3 條行為指標。\n\n"
            "【格式規定】每條限 1～2 句，須同時包含以下三個要素：\n"
            "  ・S（Situation）：觸發任務的情境或條件，如「當[事件]發生時」\n"
            "  ・A（Action）：具體採取的行動，以行動動詞開頭\n"
            "  ・R（Result）：可觀察或可衡量的成果\n"
            "建議句型：當[情境]（S），採取[行動]（A），達成[成果]（R）。\n\n"
            "【格式示意（勿照抄，需依員工描述生成）】\n"
            "  ✓「當[觸發情境]（S），立即[具體行動]（A），使[可衡量成果]達成（R）。」\n\n"
            "禁止輸出解釋或多餘文字。"
        ),
        "user": (
            "職位：{position}｜職能等級：{level} — {level_hint}\n"
            "工作任務：{task_name}\n\n"
            "【員工描述（如何執行）】\n{user_desc}\n\n"
            "【工作產出/成果】\n{user_output}\n\n"
            "【任務相關背景知識（了解任務範疇用，輸出格式無需仿照）】\n{std_text}\n\n"
            "請從員工描述中找出觸發情境(S)、具體行動(A)、可衡量成果(R)，生成行為指標。\n"
            '只輸出 JSON，格式：{{"behavior_indicators":["指標1","指標2","指標3"]}}'
        ),
    },
}


def _build_prompt_messages(
    position: str,
    task_name: str,
    user_description: str,
    standard_behaviors: list,
    template: str = "ABCD",
    level: int = 3,
    user_output: str = "",
) -> list[ChatCompletionMessageParam]:
    """組裝 system/user prompt messages（供 worker process 使用）。

    template: 分析框架，"5W2H" / "ABCD" / "STAR"，預設 "ABCD"
    level:    職能等級 1–5，用於差異化提示
    user_output: 工作產出描述，補充 context 給 LLM
    """
    std_lines = []
    for b in standard_behaviors[:5]:
        if isinstance(b, dict):
            std_lines.append(b.get("description", ""))
        elif isinstance(b, str):
            std_lines.append(b)
    std_text  = "\n".join(f"- {l}" for l in std_lines if l) or "（無標準行為指標）"
    user_desc = user_description.strip() or "（員工未填寫）"
    out_text  = user_output.strip() or "（未填寫）"
    level_hint = _LEVEL_HINT.get(level, "")

    tpl = PROMPT_TEMPLATES.get(template, PROMPT_TEMPLATES["ABCD"])
    system_prompt = tpl["system"]
    user_prompt = tpl["user"].format(
        position=position,
        level=level,
        level_hint=level_hint,
        task_name=task_name,
        user_desc=user_desc,
        user_output=out_text,
        std_text=std_text,
    )
    return [
        {"role": "system", "content": system_prompt},
        {"role": "user",   "content": user_prompt},
    ]


def _build_classify_messages(position: str, task_name: str, user_desc: str) -> list:
    """第一步：僅判斷框架，輸出 {"template":"ABCD"}，不生成指標。"""
    return [
        {"role": "system", "content": (
            "你是 ICAP 職能說明書專家。\n"
            "請依下列順序判斷最適合的行為指標分析框架，只輸出 JSON，禁止輸出其他文字。\n\n"
            "【判斷順序（依序排除）】\n"
            "1. STAR：員工描述的核心是「回應突發事件或非預期狀況」，且需自主判斷或跨部門協調。"
            "僅在任務明確屬於非例行應變時才選用。\n"
            "2. 5W2H：描述的核心是「按既定步驟反覆執行的操作流程」，強調執行方式與頻率。\n"
            "3. ABCD：其餘情況，包含任何有明確可衡量產出的任務，以此作為預設選項。\n"
        )},
        {"role": "user", "content": (
            f"職位：{position}\n"
            f"工作任務：{task_name}\n"
            f"員工描述：{user_desc or '（未填寫）'}\n\n"
            '只輸出 JSON，例如：{"template":"ABCD"}'
        )},
    ]


def _persistent_worker(input_q, result_q):
    """
    長駐子 process：模型只載入一次，持續從 input_q 讀取任務並寫入 result_q。
    input_q 格式：(idx, task_args, task_hash) 或 None（結束 sentinel）
    result_q 格式：{"type":"ready"} | {"type":"result", "idx":..., ...}
    """
    import re as _re, json as _json
    from pathlib import Path as _Path

    backend = None
    if _Path(TAIDE_MODEL_PATH).exists():
        try:
            backend = _LlamaCppBackend(TAIDE_MODEL_PATH)
        except Exception:
            pass
    if backend is None:
        backend = _LMStudioBackend()

    result_q.put({"type": "ready"})

    _VALID_TPLS = {"5W2H", "ABCD", "STAR"}

    while True:
        try:
            item = input_q.get(timeout=300)
        except Exception:
            continue
        if item is None:
            break
        idx, task_args, task_hash = item
        tpl_param = task_args.get("template", "ABCD")
        try:
            if tpl_param == "AUTO":
                # ── 兩步式 AUTO ──────────────────────────────────────────────
                # Step 1：僅判斷框架，不生成指標
                tpl_used = "5W2H"   # 預設 fallback（分類失敗或無法匹配時）
                try:
                    cls_messages = _build_classify_messages(
                        position=task_args.get("position", ""),
                        task_name=task_args.get("task_name", ""),
                        user_desc=task_args.get("user_description", ""),
                    )
                    cls_reply = backend.chat(cls_messages)
                    cs = cls_reply.find('{')
                    ce = cls_reply.rfind('}')
                    if cs != -1 and ce > cs:
                        cls_data = _json.loads(cls_reply[cs:ce + 1])
                        detected = cls_data.get("template", "5W2H")
                        tpl_used = detected if detected in _VALID_TPLS else "5W2H"
                except Exception:
                    pass   # 分類失敗 → 以 5W2H 繼續

                # Step 2：用選定框架生成指標
                gen_args = {**task_args, "template": tpl_used}
                messages = _build_prompt_messages(**gen_args)
                reply = backend.chat(messages)
            else:
                tpl_used = tpl_param
                messages = _build_prompt_messages(**task_args)
                reply = backend.chat(messages)

            # 從第一個 { 到最後一個 } 提取 JSON（避免非貪婪 regex 被內嵌花括號截斷）
            json_str = None
            s = reply.find('{')
            e = reply.rfind('}')
            if s != -1 and e > s:
                json_str = reply[s:e + 1]

            indicators = None
            if json_str:
                try:
                    data = _json.loads(json_str)
                    raw = data.get("behavior_indicators", [])
                    if isinstance(raw, str) and raw.strip():
                        raw = [raw]
                    if isinstance(raw, list) and raw:
                        indicators = _split_indicators(raw) or None
                except Exception:
                    pass   # JSON 解析失敗 → fallback 到逐行解析
            if not indicators:
                lines = [l.strip().lstrip("-•・ ")
                         for l in reply.split("\n")
                         if len(l.strip()) > 10
                         and not l.strip().startswith('"')
                         and not l.strip().startswith('{')
                         and not l.strip().startswith('[')
                         and not l.strip().endswith('：')
                         and not l.strip().endswith(':')
                         and not _PREAMBLE_RE.search(l.strip())]
                indicators = _split_indicators(lines[:6])
            result_q.put({"type": "result", "idx": idx,
                          "indicators": indicators or [],
                          "template_used": tpl_used, "task_hash": task_hash})
        except Exception:
            safe_tpl = tpl_param if tpl_param in _VALID_TPLS else "5W2H"
            result_q.put({"type": "result", "idx": idx, "indicators": [],
                          "template_used": safe_tpl, "task_hash": task_hash})


def create_persistent_worker():
    """
    建立並啟動長駐 LLM 子 process。
    回傳 (process, input_q, result_q)，呼叫端透過 input_q 送任務、result_q 取結果。
    """
    import multiprocessing as _mp
    input_q  = _mp.Queue()
    result_q = _mp.Queue()
    p = _mp.Process(target=_persistent_worker, args=(input_q, result_q), daemon=True)
    p.start()
    return p, input_q, result_q


def _split_indicators(raw: list) -> list:
    """將 LLM 可能合併成單一字串的多條指標拆開，去除「指標N:」等前綴，限回傳 3 條。"""
    result = []
    for item in raw:
        # LLM 有時把指標包成 {"指標1": "text"} dict，取 value
        if isinstance(item, dict):
            item = next(iter(item.values()), "")
        item = str(item).strip()
        if not item:
            continue
        # 去除 LLM 在字串外多加的引號（如 '"每日..."'）
        if item.startswith('"') and item.endswith('"') and len(item) > 2:
            item = item[1:-1].strip()
        # 跳過看起來是 JSON 結構的行（不是真正的指標）
        if item.startswith('{') or item.startswith('['):
            continue
        # 若含換行、中文分號或「指標N:」模式，視為多條合併，拆開
        raw_parts = re.split(r'\n|；|(?:指標\s*\d+\s*[:：])', item)
        # 合併因 JSON 字串內嵌換行產生的短斷片（如 "採購、銷\n售及費用..." 被拆成 ["採購、銷","售及費用..."]）
        parts: list[str] = []
        buf = ""
        for p in raw_parts:
            p = p.strip()
            if not p:
                continue
            if buf and len(buf) < 12:
                buf += p        # 短斷片與下一段合併，視為同一指標
            elif buf:
                parts.append(buf)
                buf = p
            else:
                buf = p
        if buf:
            parts.append(buf)

        for p in parts:
            p = p.strip().lstrip("-•・ ")
            # 去除行首的「N.」「N、」「N)」等編號
            p = re.sub(r'^\d+[\.、\)）]\s*', '', p)
            if len(p) > 10:
                result.append(p)
    return result[:3]
