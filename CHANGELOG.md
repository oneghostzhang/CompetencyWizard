# CHANGELOG

> 記錄各版本的架構決策與重要修復背景。
> 近期修復（v2.0.x）的細節見各 git commit message；此處保留 git 訊息未涵蓋的「為什麼這樣做」。

---

## v2.1.x — AI 模板選擇（2026-05）

**新增 5W2H / ABCD / STAR / AUTO 四種分析框架**

- 每個工作任務在 Detail 頁各自選擇框架，預設為 AUTO
- AUTO 模式：LLM 在同一次呼叫中先判斷框架再生成指標（不額外增加 API 呼叫次數）
- AUTO 輸出 JSON 格式：`{"template":"ABCD","behavior_indicators":[...]}`，解析失敗時 fallback 到 ABCD
- `analyze_tasks_batch()` 改為 `row.get("template", global_default)`，全域設定僅作 fallback

---

## v2.0.x — 穩定性修復（2026-04）

見 git commits `a66a927`～`820270d`，包含 P1–P10、C1–C5、T1–T3、O1–O6 共 24 項（6 項驗證為誤報）。

**關鍵決策：**
- P3 舊執行緒未停止 → `_run_llm()` 建立新執行緒前先 `stop()` + `disconnect()` + `wait(2000)`
- T1 初始化無限等待 → 5 分鐘 QTimer + 取消按鈕（軟逾時，不強殺）
- T3 搜尋卡住 → 30 秒 QTimer 斷開 signal（不用 `TerminateThread`，避免 Qt 資源未釋放）
- C4 FAISS 維度不符 → `_try_load_cache()` 比對 `index.d` 與模型維度，不符時強制重建

---

## v2.0.0 — 架構重設計（2026-03）

**從 AI 對話式改為逐任務填寫精靈**

六頁流程取代原本的聊天介面：
載入 → 搜尋職能基準 → 確認基準書 → 逐任務填寫 → LLM 建議確認 → 補充匯出

---

## v1.4.10 — LLM 後端整合（2026-03）

**問題：LM Studio HTTP 連線逾時**
初始版本依賴 LM Studio REST API（`http://localhost:1234/v1`），使用者未啟動伺服器時會長時間等待後 timeout，行為指標生成完全失效。

**決策：改用 LlamaCpp 直接載入 GGUF，保留 LM Studio 作為 fallback**
- 主後端：`_LlamaCppBackend`，本機推論，無需網路
- 備用後端：`_LMStudioBackend`，OpenAI 相容 API
- 啟動時 `llamacpp_available()` 自動偵測可用後端

---

## v1.4.9 — llama.cpp 崩潰隔離（2026-03）

**問題：`GGML_ASSERT abort` 讓整個 App 閃退**
llama.cpp 底層為 C/C++，斷言失敗時直接呼叫 `abort()`，繞過 Python `try/except`，整個 Qt 應用程式瞬間終止，使用者所有已輸入資料遺失。

**決策：LLM 推論移至獨立 `multiprocessing.Process`**
- 主程式 ↔ 子程序透過 `Queue` 通訊
- 子程序崩潰時（`is_alive()` = False）：已完成任務保留，未完成任務標記為空，主程式繼續運行
- `daemon=True` 確保主程式退出時子程序自動清理
- Windows 打包需在 `main.py` 加入 `freeze_support()`

---

## v1.4.11 — PDF 解析修正（2026-03）

**問題：主要職責解析只保留最後一筆（狀態機邏輯錯誤）**
`_parse_responsibilities_from_tables()` 偵測到新 T-code 時，未先儲存前一個職責就直接覆蓋，T1～T(n-1) 全部遺失。

**修正：**存新職責前先 `current_resp → responsibilities`，再重置 `current_task = None`。
修正後對 908 份 ICAP PDF 全量重新解析。

**問題：3 位數知識/技能代碼解析錯誤**
原正規表達式 `K\d{2}` 只匹配 2 位，`K004` 被切成代碼 `K00`、名稱 `4工藝...`。
修正：4 處 `\d{2}` 改為 `\d+`。

---

## v1.x — RAG 向量系統修正

**快取命中時仍讀取過期資料**
`_standards` 字典原本從快取還原，PDF 更新後舊資料不會刷新。
決策：快取命中時僅復用 FAISS 索引（重建耗時 1–3 分鐘），`_standards` 每次從磁碟 JSON 重新讀取。

**新增欄位後舊快取不相容**
新增 `standard_category` 後，舊快取 chunk 缺少此欄位。
處理：載入時檢查第一個 chunk 是否含目標欄位，不含則強制重建。

---

## v1.x — PyQt6 介面修正

**`QCheckBox.setWordWrap()` 不存在（Qt6 API 差異）**
`QCheckBox` 繼承自 `QAbstractButton`，Qt6 未提供 `setWordWrap`，呼叫即閃退。
修正：`QHBoxLayout` 橫排 `QCheckBox`（僅勾選） + `QLineEdit`（顯示文字），以 `List[Tuple[QCheckBox, QLineEdit]]` 儲存，收集時用 `le.text()` 而非 `cb.text()`。

---

## v1.x — 開發環境

**config.toml 設定檔**
模型路徑與推論參數（`n_ctx`、`n_threads`、`temperature`、`max_tokens`）原本硬編碼在 `ai_chat.py`。
改以 `tomllib`（Python 3.11+ 內建）讀取 `config.toml`，找不到時退回預設值。
`config.toml` 列入 `.gitignore`，提供 `config.example.toml` 作為範本。
