# CHANGELOG

> 記錄各版本的架構決策與重要修復背景。
> 近期修復（v2.0.x）的細節見各 git commit message；此處保留 git 訊息未涵蓋的「為什麼這樣做」。

---

## v2.3.x — Hub-and-Spoke UI 與預覽頁（2026-05）

**從線性多頁流程改為任務卡片 Hub + 單任務編輯頁**

- Step 3/4 合併重設計：任務總覽 Hub（QGridLayout 卡片網格） + 單任務編輯頁，取代原本的逐任務 Detail 頁與批次 Suggest 頁
- 任務卡片以邊框顏色顯示 LLM 狀態（灰=未提交、橘=分析中、綠=完成、紅=需更新）
- Task Edit 頁：AI 結果區塊在分析完成前隱藏（`setVisible(False)`），完成後才顯示，避免空白佔位干擾使用者
- 匯出前新增唯讀預覽頁（Page 5）：`QTabWidget` 含四個 `QTableWidget` 分頁（職能說明書、知識清單、技能清單、態度清單），色彩配置與 `excel_exporter.py` 完全一致（深藍表頭、淺藍主責欄、淺綠行為指標欄、淺黃知識、淺紫技能、淺粉紅態度）

**AI 指標解析三層修復（`ai_chat.py`）**

- **Bug 1 — `indicators is None` 判斷失效**：`_split_indicators()` 回傳空列表 `[]` 時，`indicators is None` 為 False，導致 fallback 逐行解析被跳過，最終送出空結果。修正：改用 `if not indicators:` + `_split_indicators(...) or None`，空列表同樣觸發 fallback。
- **Bug 2 — `behavior_indicators` 為字串而非列表**：部分 LLM 輸出 `{"behavior_indicators":"文字"}` 而非陣列，`isinstance(raw, list)` 為 False 直接跳過。修正：偵測到字串時包成 `[raw]` 再送進 `_split_indicators()`。
- **Bug 3 — 指標外層多餘引號被整行丟棄**：部分 LLM 輸出 `["\"每日...\""，...]`，外層引號觸發 `startswith('"')` 過濾。修正：先去除前後引號再判斷，不再整行丟棄。
- 以上三個修復同時套用至 AUTO 模式與固定模板模式。

---

## v2.2.x — 長駐 LLM Worker（2026-05）

**從「全部填完後批次分析」改為「逐任務即時分析」**

- 每次在 Detail 頁（Step 3）儲存任務描述，後端立即提交給長駐 LLM 子 process 分析
- 長駐 worker 模型只載入一次，不需要每批重新載入（省去 30–60 秒等待）
- Detail 頁每個任務右上角顯示 LLM 狀態 badge：`●未提交` / `●分析中` / `✓已完成` / `↻需更新`
- Suggest 頁（Step 4）改為「佔位渲染」：先建立所有任務的空框，結果逐一動態填入
- `_task_hash()`：MD5 計算 `task_name + user_description + user_output + template + level`，偵測內容是否變更；過期結果（hash 不符）自動丟棄
- `TaskLLMState` enum：IDLE / PENDING / DONE / STALE
- `PersistentLLMWorker(QThread)` 替代原 `LLMAnalyzeThread`，包裝長駐子 process
- `create_persistent_worker()`：工廠函式，建立 dual-Queue（input_q + result_q）長駐子 process
- 關閉視窗時 `closeEvent()` 送 sentinel 並 kill 子 process，避免殭屍 process

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
