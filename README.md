# CompetencyWizard 職能說明書精靈

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-lightgrey)
![UI](https://img.shields.io/badge/UI-PyQt6-41CD52?logo=qt&logoColor=white)
![Version](https://img.shields.io/badge/Version-v2.0.3-orange)
![AI](https://img.shields.io/badge/AI-LlamaCpp%20TAIDE-blueviolet)

> 以 RAG + LLM 為核心的職能說明書製作工具。員工只需輸入職業名稱，系統自動搜尋最相近的 ICAP 職能基準並預填結構化欄位，員工逐任務填寫工作詳情後，LLM 自動生成符合 ICAP 格式的行為指標，最終輸出標準格式 Excel 職能說明書。

---

## 目錄

- [系統架構](#️-系統架構)
- [核心功能](#-核心功能)
- [專案結構](#-專案結構)
- [快速開始](#-快速開始)
- [使用流程](#-使用流程)
- [AI 引導填寫](#-ai-引導填寫)
- [模組說明](#-模組說明)
- [效能](#-效能)
- [技術棧](#️-技術棧)
- [注意事項](#️-注意事項)
- [更新日誌](#-更新日誌)

---

## 🏗️ 系統架構

```
┌─────────────────────────────────────────────────────────┐
│                    資料輸入層                            │
│   ICAP 職能基準 PDF  →  pdf_parser_v2.py  →  JSON       │
└─────────────────────────┬───────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────┐
│                    向量索引層                            │
│   wizard_rag.py (WizardRAG)                             │
│   ├── BAAI/bge-base-zh-v1.5 Embedding 模型             │
│   └── FAISS 向量索引快取（_index_cache/）               │
└─────────────────────────┬───────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────┐
│                    分析引擎層                            │
│   gap_analyzer.py (GapAnalyzer)                         │
│   ├── 逐任務 5W2H → 搜尋查詢字串轉換                   │
│   ├── Layer 1：語義相似度排名（Top-K 匹配職能基準）     │
│   └── 完整度分數計算 + 缺口嚴重度分級                  │
└─────────────────────────┬───────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────┐
│                    介面層                                │
│   wizard_ui.py (PyQt6 桌面應用程式)                     │
│   ├── Search：職業名稱搜尋 → Top-3 職能基準             │
│   ├── Editor：職能基準書編輯（主責/任務/產出/等級）     │
│   ├── Detail：逐任務填寫工作描述                        │
│   ├── LLMSuggest：LLM 自動生成行為指標（可編輯）        │
│   ├── Supplement：補充說明                              │
│   ├── Step 2：分析結果 + 職能基準分頁瀏覽               │
│   ├── Step 3：StandardAdoptionWizard 逐項確認精靈       │
│   └── Step 4：確認缺口 → 匯出 Excel（6 個 Sheet）      │
└─────────────────────────────────────────────────────────┘
```

---

## 🎯 核心功能

| 功能 | 說明 |
|------|------|
| 🔍 **語意向量職能搜尋** | 輸入職業名稱，`BAAI/bge-base-zh-v1.5` + FAISS 從 900+ 份 ICAP 職能基準中找出 Top-3 最相似標準 |
| 📋 **職能基準書編輯器** | 系統預填主責代碼 / 主責名稱 / 任務代碼 / 任務名稱 / 工作產出 / 職能等級；員工可新增、刪除、直接點格子修改任一欄位 |
| ✍️ **逐任務工作描述填寫** | 每個任務分別填寫「實際如何執行此任務」與「主要工作產出」，引導員工具體描述實際工作內容 |
| 🤖 **LLM 行為指標自動生成** | 根據員工填寫的任務描述，呼叫本地 LlamaCpp（TAIDE GGUF）自動生成 2–3 條 ICAP 格式行為指標；子 process 隔離，生成中程式不凍結 |
| ✏️ **行為指標可直接編輯** | LLM 生成結果以可編輯文字框呈現，員工可勾選採用、手動修改內容，或按「重新 AI 分析」重新生成 |
| 📤 **匯出 Excel（5 Sheet）** | 職能說明書 / 知識清單 / 技能清單 / 態度清單 / 補充說明，完整對齊 ICAP 職能基準書格式 |
| ⚡ **索引快取** | FAISS 索引首次建立後自動快取，後續啟動 2–5 秒直接載入，JSON 更新後自動同步 |
| 🗂️ **資料管理** | 頂部「資料管理」按鈕：新增／刪除 PDF、解析 PDF→JSON、搜尋過濾清單、重建向量索引 |

---

## 📁 專案結構

```
CompetencyWizard/
│
├── 🐍 核心模組
│   ├── main.py              # 程式入口（QApplication 啟動）
│   ├── wizard_ui.py         # PyQt6 主視窗 UI（四頁 Stack）
│   ├── ai_chat.py           # AI 對話模組 v2.1（5 階段職能說明書引導）
│   ├── wizard_rag.py        # RAG 核心（Embedding + FAISS 檢索）
│   ├── gap_analyzer.py      # 5W2H 缺口分析邏輯與資料結構
│   ├── logger.py            # 集中式日誌設定（RotatingFileHandler）
│   ├── excel_exporter.py    # openpyxl Excel 輸出（6 Sheet）
│   └── pdf_parser_v2.py     # PDF 職能基準解析工具
│
├── 📂 data/
│   ├── raw_pdf/             # 原始職能基準 PDF（不納入版控）
│   └── parsed_json_v2/      # 解析後的結構化 JSON（不納入版控）
│
├── 📄 test_cases.txt        # 測試資料（9 個案例，逐任務 5W2H 格式）
└── 📂 _index_cache/         # FAISS 向量索引快取（自動產生，不納入版控）
```

---

## 🚀 快速開始

### 1. 下載專案

```bash
git clone https://github.com/oneghostzhang/CompetencyWizard.git
cd CompetencyWizard
```

### 2. 安裝依賴

#### 方式 A（推薦）：使用 uv

```bash
uv sync
```

若需啟用 AI 對話 / 行為指標生成功能（LlamaCpp / OpenAI fallback）：

```bash
uv sync --extra ai
```

#### 方式 B：使用 pip

```bash
pip install PyQt6 sentence-transformers faiss-cpu openpyxl openai
```

> 首次執行會自動下載 `BAAI/bge-base-zh-v1.5` 模型（約 400 MB），需要網路連線。

### 3. 準備職能基準資料

將 ICAP 職能基準 JSON 檔放入 `data/parsed_json_v2/` 資料夾：

```
CompetencyWizard/
└── data/
    └── parsed_json_v2/
        ├── TFB5120-003v3.json
        ├── HBR2431-001v4.json
        └── ...
```

### 4. 啟動程式

使用 uv：

```bash
uv run python main.py
```

或（已啟用虛擬環境時）：

```bash
python main.py
```

---

## 📋 使用流程

```
啟動 → 自動載入 Embedding 模型 + 建立 FAISS 索引
  ↓
Step 1：填寫職業名稱 → 系統向量搜尋 Top-3 最相近職能基準
  └── 選擇符合的基準（或不使用基準，從空白填寫）
  ↓
Step 2：職能基準書編輯器
  ├── 系統預填：主責代碼 / 主責名稱 / 任務代碼 / 任務名稱 / 工作產出 / 職能等級
  └── 員工可新增、刪除、直接點格子修改任一欄位
  ↓
Step 3：逐任務填寫工作詳情
  ├── 每個任務填寫：「您實際如何執行此任務？」
  └── 每個任務填寫：「主要工作成果或產出？」
  ↓
Step 4：LLM 自動生成行為指標
  ├── 每個任務 → LLM 根據描述生成 2-3 條 ICAP 格式行為指標
  ├── 員工勾選採用或手動補充修改
  └── 可按「重新 AI 分析」重跑
  ↓
Step 5：填寫說明與補充事項（選填）
  └── 員工姓名 + 備注說明
  ↓
匯出 Excel 職能說明書（3 個 Sheet）
```

---

## 🤖 LLM 行為指標生成

員工填寫工作任務的描述後，系統呼叫本地 LLM **自動生成 ICAP 格式行為指標**，員工勾選採用或手動修改。

### AI 推論後端

使用 **LlamaCpp** 直接載入 GGUF 模型，無 HTTP timeout，完全離線推論。

將 GGUF 模型放至預設路徑，系統啟動時自動載入：
```
C:\Users\<你的帳號>\.lmstudio\models\ZoneTwelve\TAIDE-LX-7B-Chat-GGUF\TAIDE-LX-7B-Chat.Q4_K_S.gguf
```

### 支援模型

| 模型 | 繁中支援 | 說明 |
|------|----------|------|
| `TAIDE-LX-7B-Chat` | ★★★★★ | 台灣政府出品，最適合繁中職場語境，首選 |
| `Qwen3-8B` | ★★★★☆ | 阿里巴巴，推理能力強 |
| `gemma-3n-E4B` | ★★★ | Google，速度較快 |

### 5 階段對話流程

```
表單頁點選「開始對話 →」（綠色按鈕）
  ↓（首次載入模型約 10–30 秒）
【Phase 1】AI 詢問基本資訊
  職位名稱 → 公司/部門 → 工作內容描述 → 職能等級(1–5)
  ↓
【Phase 2】ICAP 職能基準確認
  系統自動 RAG 搜尋最相似標準 → AI 介紹基準並確認是否符合
  ↓
【Phase 3】主要職責清單確認
  AI 列出標準職責（表格格式），員工增刪修改
  ↓
【Phase 4】逐職責深度訪談
  每項職責：描述實際工作 → 確認子任務（表格）→ 產出/知識/技能摘要（表格）
  ↓
【Phase 5】輸出職能說明書
  AI 輸出完整 JSON → 點選「確認並匯入任務 →」返回清單
```

> AI 對話和手動填寫可以**混用**，匯入不會覆蓋已有的手動任務。
> 所有資料在本機處理，不會上傳至任何外部伺服器。

---

## 🔧 模組說明

<details>
<summary><b>wizard_rag.py</b> — RAG 核心</summary>

- 使用 `BAAI/bge-base-zh-v1.5` 將職能基準 JSON 的 `chunks_for_rag` 欄位向量化
- FAISS IndexFlatIP 建立餘弦相似度索引，支援 Top-K 檢索
- 索引自動快取至 `_index_cache/`，避免每次啟動重建
</details>

<details>
<summary><b>gap_analyzer.py</b> — 缺口分析</summary>

- `UserInput5W2H` 資料類別：結構化儲存 9 個面向的使用者輸入，`task_list` 保留各任務獨立字串清單
- `to_search_query()` 將所有任務的 5W2H 合併為語義搜尋查詢字串
- `GapAnalyzer.analyze()` 執行 RAG 檢索 + 完整度評分 + 缺口嚴重度分級
- 缺口嚴重度分級：核心製程關鍵字 → 高；衛生/行政/清潔關鍵字 → 低；其餘 → 中
- `GapReport` 含 `matched_standards`（排名列表）、`completeness_score`、`gap_items`（含嚴重度）
</details>

<details>
<summary><b>ai_chat.py v2.1</b> — AI 對話模組（5 階段職能說明書引導）</summary>

- **推論後端**：`_LlamaCppBackend` 直接載入 GGUF，無 HTTP timeout，子 process 隔離防止 C-level abort 崩潰
- TAIDE Llama2 chat template：`[INST] <<SYS>>\n{system}\n<</SYS>>\n\n{user} [/INST]`
- 固定開場白（`GREETING`）瞬間顯示，不呼叫 LLM
- `SYSTEM_PROMPT`：5 階段狀態機，含 markdown 表格格式指令
- `inject_standard()`：Phase 2 將 RAG 搜尋結果注入對話歷史，觸發 AI 介紹基準
- `extract_competency_json()` 偵測 `[COMPETENCY_JSON]...[/COMPETENCY_JSON]` 區塊並解析
- `competency_to_task_list()`：職能 JSON → 5W2H dict 清單，供表單匯入
- 對話歷史自動截斷（最多 20 輪），防止 context 無限成長
- `llamacpp_available()` / `check_server()`：後端可用性偵測
</details>

<details>
<summary><b>wizard_ui.py</b> — PyQt6 桌面 UI</summary>

- 四頁 `QStackedWidget`：載入頁 → 輸入表單頁 → 結果頁 → AI 對話頁
- `InitThread` / `AnalyzeThread`：背景執行緒避免 UI 凍結
- **逐任務輸入架構**：
  - 表單頁頂部固定顯示「已加入任務清單」面板（不隨 5W2H 欄位捲動）
  - `_added_tasks: List[dict]`：每項任務儲存完整 9 欄位 5W2H dict
  - 「加入清單 ＋」在表單底部，填完後收集全部欄位 → 存入 dict → 清空表單 → 自動捲回頂端
  - `_collect_input()` 合併所有 dict 產生 `UserInput5W2H`
- **AI 對話頁**：
  - `ChatWorker(QThread)`：背景執行緒，含 `reply / status / error` 三個 Signal
  - 自動初始化 LlamaCpp 後端（首次載入 10–30 秒）
  - `_markdown_to_html()`：AI 回覆中的 markdown 表格自動轉為 HTML，在 `QTextBrowser` 視覺化渲染
  - 對話完成後顯示「確認並匯入任務 →」按鈕
- `TaskEditDialog`：彈出式任務編輯對話框；點「編輯」開啟，含完整 9 個 5W2H 欄位，儲存後原地更新清單
- `StandardSelectorDialog`：從職能基準資料庫選擇範本，每個標準任務帶入 task_name/output/behaviors/skills 等欄位加入清單
- `StandardAdoptionWizard`：分析完成後 Opt-out 確認精靈；三分頁（任務／知識／技能），預勾全部，綠色=自動偵測，藍色=預選待確認
- 結果頁右側 `QTabWidget`：基本資訊 / 工作職能（task_id 下拉）/ 缺口分析
- `DataManagerDialog`：新增／刪除 PDF、PDF→JSON 解析、搜尋過濾、重建索引
</details>

<details>
<summary><b>excel_exporter.py</b> — Excel 匯出（6 Sheet）</summary>

| Sheet | 內容 |
|-------|------|
| 職能說明書摘要 | 員工姓名、建立日期、匹配基準、完整度、5W2H 輸入摘要 |
| 我的職能確認 | 員工確認的任務／知識／技能清單（含代碼，綠色樣式） |
| 工作任務對照 | 確認任務 vs 缺口任務對照（含嚴重度色彩：高=紅、中=橘、低=黃） |
| 知識技能對照 | 確認知識技能 vs 缺口對照（含代碼與嚴重度） |
| 缺口分析報告 | 缺口摘要、建議補強方向 |
| 完整職能基準資料 | 匹配職能基準的完整任務、知識、技能清單 |

</details>

<details>
<summary><b>pdf_parser_v2.py</b> — PDF 解析工具</summary>

- 使用 pdfplumber 提取 ICAP 職能基準 PDF 的結構化資料
- 輸出含 `chunks_for_rag` 的 JSON，可直接供 WizardRAG 使用
</details>

---

## 📈 效能

| 操作 | 時間 | 備註 |
|------|------|------|
| Embedding 模型載入 | 30–60 秒 | 首次下載模型後快取 |
| FAISS 索引建立 | 1–3 分鐘 | 900+ JSON，建立後快取 |
| 索引快取載入 | 2–5 秒 | 第二次起直接讀取 |
| 單次分析查詢 | 1–3 秒 | 已建立索引後 |
| Excel 匯出 | < 1 秒 | |

---

## 🛠️ 技術棧

| 層級 | 技術 | 用途 |
|------|------|------|
| 桌面 UI | PyQt6 | 操作介面 |
| 本地 LLM | LlamaCpp + TAIDE GGUF | 行為指標自動生成（analyze_task，子 process 隔離） |
| Embedding | sentence-transformers | 文本向量化（bge-base-zh-v1.5） |
| 向量檢索 | FAISS | 高效相似度搜尋 |
| Excel 輸出 | openpyxl | 職能說明書格式化輸出 |
| PDF 解析 | pdfplumber | ICAP 職能基準 PDF 轉 JSON |

---

## ⚠️ 注意事項

1. **資料格式**：`parsed_json_v2/` 的 JSON 須包含 `chunks_for_rag` 欄位，可使用 `pdf_parser_v2.py` 從原始 PDF 生成
2. **記憶體需求**：Embedding 模型約 500 MB，完整 900+ JSON 索引約 300 MB，建議系統記憶體 ≥ 8 GB
3. **資料來源**：職能基準資料來自 [ICAP 職能發展應用平台](https://icap.wda.gov.tw/)，僅供學習研究使用
4. **索引更新**：可透過頂部「資料管理」按鈕新增 PDF、解析後重建索引；載入頁亦提供「強制重建索引」按鈕

---

## 📋 更新日誌

| 版本 | 日期 | 更新內容 |
|------|------|---------|
| v2.0.3 | 2026-03-30 | **防閃退**：LLM 分析改用子 process 隔離，llama.cpp `GGML_ASSERT abort` 只殺子 process，主程式不受影響，已完成任務保留結果；**LLM建議頁**：行為指標改為可直接編輯的 `QLineEdit`、新增 `_split_indicators()` 自動拆分合併格式、重新分析按鈕加資料來源說明；**Excel 新增態度清單 Sheet**（A代碼）、K/S 對應任務欄顯示全部來源任務；**Log** 改為每次啟動產生獨立時間戳記檔，保留最近 30 份 |
| v2.0.2 | 2026-03-30 | 修正 `WizardRAG` 快取命中時仍讀取舊版 standards 的問題：改為每次從 JSON 重新載入，確保重新解析 PDF 後資料立即生效；修正逐任務填寫頁第一筆任務無法返回編輯器（按鈕改顯示「← 返回編輯器」並始終啟用）；全量重解析 908 份職能基準 PDF，補齊先前解析器 bug 遺漏的中間任務 |
| v2.0.0 | 2026-03-30 | 全面重新設計系統流程：UI 改為 6 頁流程（搜索→編輯器→逐任務填寫→LLM建議確認→補充→匯出）；移除 5W2H 表單，改為直接填寫職能基準書格式；新增 `analyze_task()` 單次 LLM 呼叫自動生成行為指標；`excel_exporter.py` 全新格式對齊 ICAP 職能基準書欄位（3 Sheet：職能說明書/知識清單/技能清單） |
| v1.4.11 | 2026-03-27 | 修正 `pdf_parser_v2.py` 主要職責只抓到最後一筆的 bug：新職責覆蓋前未先儲存 `current_task`，導致 T1~T3 任務全部遺失。現在解析完整（如會計助理：T1.1~T4.1 皆正確輸出） |
| v1.4.8 | 2026-03-21 | 新增 AI 對答式引導填寫：本地 LLM 扮演 HR 助理，透過對話引導員工描述工作任務，完成後自動整理 5W2H 格式匯入清單；新增 ai_chat.py 模組；ChatWorker 背景執行緒；固定開場白 |
| v1.4.7 | 2026-03-18 | 任務清單面板加入收合/展開功能，防止任務過多時覆蓋表單操作區域 |
| v1.4.6 | 2026-03-18 | 新增 TaskEditDialog 彈出式任務編輯對話框；點「編輯」開啟，儲存後原地更新清單；移除舊的「載回主表單」編輯方式 |
| v1.4.5 | 2026-03-18 | 範本載入改為逐任務完整 5W2H dict：task_name/output/behaviors/skills 等欄位自動對應；載入後自動捲回頂端 |
| v1.4.4 | 2026-03-18 | 「加入清單 ＋」移至表單底部按鈕列；加入後自動捲回頂端；任務清單「編輯」「刪除」按鈕修正為 app 標準樣式 |
| v1.4.3 | 2026-03-18 | 每項任務儲存完整 5W2H dict；加入清單後清空全部欄位；任務清單面板顯示 What 摘要 + 角色/頻率小字 |
| v1.4.2 | 2026-03-18 | 任務清單面板移至表單頁固定頂部（不隨 5W2H 欄位捲動），始終可見 |
| v1.4.1 | 2026-03-18 | 重新設計任務輸入：移除 TaskListWidget 多行輸入，改為單一 QTextEdit + 「加入清單 ＋」按鈕；頂部新增已加入任務清單面板 |
| v1.4.0 | 2026-03-18 | 整頁一任務架構確立：每次填寫代表一項完整工作任務（含 Why / Who / When / Where / How / How Much），多任務逐項填入後統一分析 |
| v1.3.1 | 2026-03-16 | 修正 Excel 已涵蓋項目代碼欄空白、完整職能基準 sheet 關鍵欄位空白；新增「我的職能確認」sheet；確認精靈改為 opt-out 全選模式；知識/技能缺口嚴重度關鍵字分級（核心=高，衛生行政=低） |
| v1.3.0 | 2026-03-16 | 新增 `StandardAdoptionWizard` 逐項確認精靈（任務／知識／技能三頁勾選，綠色預勾已偵測項目，確認後重算完整度）；執行頻率改為勾選框；修正職能基準快速填入格式 |
| v1.2.0 | 2026-03-11 | 新增資料管理功能：DataManagerDialog（新增／刪除 PDF、PDF→JSON 解析、PDF 清單搜尋、重建向量索引） |
| v1.1.0 | 2026-03-11 | 重構 UI 樣式架構（ID 選擇器集中管理）；修正 Qt 樣式傳遞導致按鈕不可見問題；優化整體配色 |
| v1.0.0 | 2026-03-10 | 初始版本：5W2H 輸入、RAG 職能基準檢索、缺口分析、結果頁分頁瀏覽、Excel 匯出 |

---

## 📄 授權

本專案採用 [MIT License](LICENSE) 授權。

---

**版本**：v2.0.0　　**最後更新**：2026-03-30
