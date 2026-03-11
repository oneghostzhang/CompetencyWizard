# CompetencyWizard 職能說明書精靈

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-lightgrey)
![UI](https://img.shields.io/badge/UI-PyQt6-41CD52?logo=qt&logoColor=white)
![Version](https://img.shields.io/badge/Version-v1.1.0-orange)

> 以 RAG 技術為核心的職能說明書產生工具。透過 5W2H 結構化問卷，自動從台灣 ICAP 職能基準資料庫中找出最相似的標準，進行缺口分析並輸出格式化 Excel 職能說明書。

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
- [與 Graph\_RAG\_test 整合](#-與-graph_rag_test-整合)
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
│   ├── FAISS 向量索引（_index_cache/）                   │
│   └── 可復用外部 GraphRAGQueryEngine（省略重複載入）    │
└─────────────────────────┬───────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────┐
│                    分析引擎層                            │
│   gap_analyzer.py (GapAnalyzer)                         │
│   ├── 5W2H → 搜尋查詢字串轉換                          │
│   ├── 語義相似度排名（Top-K 匹配職能基準）              │
│   └── 完整度分數計算 + 缺口摘要生成                    │
└─────────────────────────┬───────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────┐
│                    介面層                                │
│   wizard_ui.py (PyQt6 桌面應用程式)                     │
│   ├── Step 1：5W2H 輸入表單                             │
│   ├── Step 2：分析結果 + 職能基準分頁瀏覽               │
│   └── Step 3：確認缺口 → 匯出 Excel                    │
└─────────────────────────────────────────────────────────┘
```

---

## 🎯 核心功能

| 功能 | 說明 |
|------|------|
| 📝 **5W2H 結構化輸入** | 引導填寫 What / Why / Who / When / Where / How / How Much 七個面向 |
| 🔍 **語意向量檢索** | `BAAI/bge-base-zh-v1.5` + FAISS，從 900+ 職能基準中找出最相似標準 |
| 📊 **職能缺口分析** | 自動比對輸入與職能基準，輸出完整度分數（0–100%）與缺口說明 |
| 📂 **分頁瀏覽職能基準** | 查看匹配標準的 metadata、basic_info、工作任務（含知識 / 技能對應） |
| ✏️ **即時修改重新分析** | 在結果頁直接修改 5W2H 欄位，一鍵重新執行分析 |
| 📤 **匯出 Excel** | 將分析結果輸出為格式化的職能說明書 `.xlsx` 檔 |
| ⚡ **索引快取** | 首次建立後自動快取，後續啟動直接載入（省略重複建立時間） |

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
│   ├── excel_exporter.py    # openpyxl Excel 輸出
│   └── pdf_parser_v2.py     # PDF 職能基準解析工具
│
├── 📂 data/
│   ├── raw_pdf/             # 原始職能基準 PDF（不納入版控）
│   └── parsed_json_v2/      # 解析後的結構化 JSON（不納入版控）
│
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

> 若已有 [Graph_RAG_test](https://github.com/oneghostzhang/RAG_test) 的 `parsed_json_v2` 資料，可直接複製或建立符號連結使用。

### 4. 啟動程式

```bash
python main.py
```

---

## 📋 使用流程

```
啟動 → 自動載入 Embedding 模型 + 建立 FAISS 索引
  ↓
填寫 5W2H 表單（What / Why / Who / When / Where / How / How Much）
  ↓
點選「開始分析 →」
  ↓
結果頁：
  ├── 左側：選擇匹配職能基準（下拉選單）
  │          可直接修改 5W2H 欄位 → 重新分析
  └── 右側：
        ├── [基本資訊] metadata + basic_info
        ├── [工作職能] 依 task_id 瀏覽任務 + 知識/技能對應
        └── [缺口分析] 完整度分數 + 缺口說明
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
- 支援傳入 `GraphRAGQueryEngine` 實例以復用已載入的 Embedding 模型
</details>

<details>
<summary><b>gap_analyzer.py</b> — 缺口分析</summary>

- `UserInput5W2H` 資料類別：結構化儲存 7 個面向的使用者輸入
- `to_search_query()` 將 5W2H 轉為語義搜尋查詢字串
- `GapAnalyzer.analyze()` 執行 RAG 檢索 + 完整度評分 + 缺口摘要
- `GapReport` 含 `matched_standards`（排名列表）、`completeness_score`、`gap_items`
</details>

<details>
<summary><b>wizard_ui.py</b> — PyQt6 桌面 UI</summary>

- 三頁 `QStackedWidget`：載入頁 → 輸入表單頁 → 結果頁
- `InitThread` / `AnalyzeThread`：背景執行緒避免 UI 凍結
- 結果頁右側 `QTabWidget`：基本資訊 / 工作職能（task_id 下拉）/ 缺口分析
- 樣式：集中於 `APP_STYLE` 常數，使用 ID 選擇器（`#objectName`）管理容器背景，避免 Qt 樣式傳遞問題
</details>

<details>
<summary><b>excel_exporter.py</b> — Excel 匯出</summary>

- 使用 openpyxl 輸出格式化 `.xlsx` 職能說明書
- 包含 5W2H 輸入內容、最佳匹配職能基準資訊、缺口分析摘要
- 自動套用欄寬、字型、填色等樣式
</details>

<details>
<summary><b>pdf_parser_v2.py</b> — PDF 解析工具</summary>

- 使用 pdfplumber 提取 ICAP 職能基準 PDF 的結構化資料
- 輸出含 `chunks_for_rag` 的 JSON，可直接供 WizardRAG 使用
- 與 [Graph_RAG_test](https://github.com/oneghostzhang/RAG_test) 共用相同格式
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

## 🔗 與 Graph_RAG_test 整合

若系統已有運行中的 [Graph_RAG_test](https://github.com/oneghostzhang/RAG_test) `GraphRAGQueryEngine` 實例，可直接傳入共用，**避免重複載入 Embedding 模型（節省 30–60 秒）**：

```python
from wizard_ui import WizardMainWindow, APP_STYLE
from PyQt6.QtWidgets import QApplication

# your_engine 為已初始化的 GraphRAGQueryEngine
app = QApplication([])
app.setStyle("Fusion")
app.setStyleSheet(APP_STYLE)
win = WizardMainWindow(engine=your_engine)
win.show()
app.exec()
```

兩個專案共用同一份 `parsed_json_v2/` 資料，無需重複存放。

---

## ⚠️ 注意事項

1. **資料格式**：`parsed_json_v2/` 的 JSON 須包含 `chunks_for_rag` 欄位，可使用 `pdf_parser_v2.py` 從原始 PDF 生成
2. **記憶體需求**：Embedding 模型約 500 MB，完整 900+ JSON 索引約 300 MB，建議系統記憶體 ≥ 8 GB
3. **資料來源**：職能基準資料來自 [ICAP 職能發展應用平台](https://icap.wda.gov.tw/)，僅供學習研究使用
4. **索引更新**：新增 JSON 後需重建索引，可在載入頁點選「強制重建索引」按鈕

---

## 📋 更新日誌

| 版本 | 日期 | 更新內容 |
|------|------|---------|
| v1.1.0 | 2026-03-11 | 重構 UI 樣式架構（ID 選擇器集中管理）；修正 Qt 樣式傳遞導致按鈕不可見問題；優化整體配色（參考 Graph_RAG_test 色系） |
| v1.0.0 | 2026-03-10 | 初始版本：5W2H 輸入、RAG 職能基準檢索、缺口分析、結果頁分頁瀏覽、Excel 匯出 |

---

## 📄 授權

本專案採用 [MIT License](LICENSE) 授權。

---

**版本**：v1.1.0　　**最後更新**：2026-03-11
