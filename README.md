# CompetencyWizard 職能說明書精靈

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-lightgrey)
![UI](https://img.shields.io/badge/UI-PyQt6-41CD52?logo=qt&logoColor=white)
![Version](https://img.shields.io/badge/Version-v1.4.6-orange)

> 以 RAG 技術為核心的職能說明書產生工具。員工以「逐任務填寫完整 5W2H」的方式描述工作內容，系統自動從台灣 ICAP 職能基準資料庫找出最相似標準，再由員工逐項確認形成完整缺口報告，最終輸出格式化 Excel 職能說明書。

---

## 目錄

- [系統架構](#️-系統架構)
- [核心功能](#-核心功能)
- [專案結構](#-專案結構)
- [快速開始](#-快速開始)
- [使用流程](#-使用流程)
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
│   ├── Step 1：逐任務 5W2H 輸入（任務清單面板 + 表單）   │
│   ├── Step 2：分析結果 + 職能基準分頁瀏覽               │
│   ├── Step 3：StandardAdoptionWizard 逐項確認精靈       │
│   └── Step 4：確認缺口 → 匯出 Excel（6 個 Sheet）      │
└─────────────────────────────────────────────────────────┘
```

---

## 🎯 核心功能

| 功能 | 說明 |
|------|------|
| 📝 **逐任務 5W2H 輸入** | 每項任務各自填寫完整的 What / Why / Who / When / Where / How / How Much，填完一項按「加入清單 ＋」後自動捲回頂端繼續填下一項 |
| 📋 **任務清單面板** | 固定顯示在表單頂部，列出所有已加入任務（含 What 摘要、角色、頻率），可隨時編輯或刪除任一任務 |
| ✏️ **任務編輯對話框** | 點「編輯」開啟 `TaskEditDialog` 彈出視窗，含完整 9 個 5W2H 欄位，儲存後原地更新清單，不影響其他任務 |
| 🚀 **職能基準快速填入** | 點選「選擇範本 →」從職能基準載入，每個標準任務自動帶入 what/output/behaviors/skills 等欄位加入清單 |
| 🔍 **語意向量檢索** | `BAAI/bge-base-zh-v1.5` + FAISS，從 900+ 職能基準中找出最相似標準 |
| 📊 **職能缺口分析** | 自動比對輸入與職能基準，輸出完整度分數（0–100%）與缺口嚴重度分級（高 / 中 / 低） |
| ✅ **逐項確認精靈（Opt-out）** | 分析完成後自動開啟 `StandardAdoptionWizard`，預勾選全部項目，員工取消不符合的項目；已自動偵測項目綠色標示，未偵測項目藍色標示 |
| 📤 **匯出 Excel（6 Sheet）** | 職能說明書摘要 / 我的職能確認 / 工作任務對照 / 知識技能對照 / 缺口分析報告 / 完整職能基準資料 |
| ⚡ **索引快取** | 首次建立後自動快取，後續啟動直接載入 |
| 🗂️ **資料管理** | 頂部「資料管理」按鈕：新增／刪除 PDF、解析 PDF→JSON、搜尋過濾清單、重建向量索引 |

---

## 📁 專案結構

```
CompetencyWizard/
│
├── 🐍 核心模組
│   ├── main.py              # 程式入口（QApplication 啟動）
│   ├── wizard_ui.py         # PyQt6 主視窗 UI（三頁 Stack）
│   ├── wizard_rag.py        # RAG 核心（Embedding + FAISS 檢索）
│   ├── gap_analyzer.py      # 5W2H 缺口分析邏輯與資料結構
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

```bash
pip install PyQt6 sentence-transformers faiss-cpu openpyxl
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

```bash
python main.py
```

---

## 📋 使用流程

```
啟動 → 自動載入 Embedding 模型 + 建立 FAISS 索引
  ↓
逐任務填寫 5W2H（可重複多次）
  ├── 填寫第一項任務的完整 What / Why / Who / When / Where / How / How Much
  ├── 按「加入清單 ＋」→ 任務加入頂部清單，表單清空
  ├── 繼續填寫下一項任務...
  └── （或點選「選擇範本 →」從職能基準自動預填全部任務）
  ↓
點選「開始分析 →」（系統合併所有任務 5W2H 進行 RAG 查詢）
  ↓
【自動開啟】職能基準逐項確認精靈（Opt-out 模式）：
  ├── 📋 工作任務頁：依主責分組，預勾全部 → 員工取消不符合的任務
  ├── 📖 知識頁：預勾全部 → 員工取消不具備的知識
  ├── 🔧 技能頁：預勾全部 → 員工取消不具備的技能
  └── 確認採用 → 系統重新計算完整度與缺口
  ↓
結果頁：
  ├── 左側：選擇匹配職能基準（下拉選單）
  │          「📋 重新確認職能」按鈕 → 隨時重開精靈調整
  └── 右側：
        ├── [基本資訊] metadata + basic_info
        ├── [工作職能] 依 task_id 瀏覽任務 + 知識/技能對應
        └── [缺口分析] 完整度分數 + 缺口嚴重度（高/中/低）
  ↓
勾選確認核取方塊 → 點選「匯出 Excel」
```

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
<summary><b>wizard_ui.py</b> — PyQt6 桌面 UI</summary>

- 三頁 `QStackedWidget`：載入頁 → 輸入表單頁 → 結果頁
- `InitThread` / `AnalyzeThread`：背景執行緒避免 UI 凍結
- **逐任務輸入架構**：
  - 表單頁頂部固定顯示「已加入任務清單」面板（不隨 5W2H 欄位捲動）
  - `_added_tasks: List[dict]`：每項任務儲存完整 9 欄位 5W2H dict
  - 「加入清單 ＋」在表單底部，填完後收集全部欄位 → 存入 dict → 清空表單 → 自動捲回頂端
  - `_collect_input()` 合併所有 dict 產生 `UserInput5W2H`
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

**版本**：v1.4.6　　**最後更新**：2026-03-18
