# CompetencyWizard 職能說明書精靈

以 RAG（Retrieval-Augmented Generation）技術為核心，協助使用者透過 5W2H 問卷快速產生職能說明書並進行缺口分析的桌面工具。

---

## 功能特色

- **5W2H 結構化輸入**：引導使用者填寫 What / Why / Who / When / Where / How / How Much 七個面向，系統化描述工作內容
- **語意向量檢索**：使用 `BAAI/bge-base-zh-v1.5` Embedding 模型 + FAISS 索引，從職能基準資料庫中找出最相似的標準
- **職能缺口分析**：自動比對使用者輸入與職能基準，輸出完整度分數與缺口說明
- **分頁瀏覽職能基準**：直接查看匹配職能基準的 metadata、basic_info、工作職能任務（含知識 / 技能對應）
- **可編輯後重新分析**：在結果頁直接修改 5W2H 內容，一鍵重新執行分析
- **匯出 Excel**：將分析結果輸出為格式化的職能說明書 `.xlsx` 檔

---

## 系統需求

| 項目 | 需求 |
|------|------|
| Python | 3.10 以上 |
| 作業系統 | Windows（已驗證）/ macOS / Linux |
| PyQt6 | 6.10+ |
| sentence-transformers | 5.0+ |
| faiss-cpu | 1.7+ |
| openpyxl | 3.1+ |

---

## 安裝

```bash
git clone https://github.com/oneghostzhang/CompetencyWizard.git
cd CompetencyWizard
pip install PyQt6 sentence-transformers faiss-cpu openpyxl
```

> 首次執行會自動下載 `BAAI/bge-base-zh-v1.5` 模型（約 400 MB），需要網路連線。

---

## 執行

```bash
python main.py
```

---

## 專案結構

```
CompetencyWizard/
├── main.py              # 程式入口
├── wizard_ui.py         # PyQt6 主視窗 UI
├── wizard_rag.py        # RAG 核心（Embedding + FAISS 檢索）
├── gap_analyzer.py      # 5W2H 缺口分析邏輯
├── excel_exporter.py    # Excel 輸出
├── pdf_parser_v2.py     # PDF 職能基準解析工具
├── data/
│   ├── raw_pdf/         # 原始職能基準 PDF
│   └── parsed_json_v2/  # 解析後的結構化 JSON
└── _index_cache/        # FAISS 向量索引快取（自動產生）
```

---

## 使用流程

```
啟動 → 初始化模型索引
  ↓
填寫 5W2H 表單
  ↓
開始分析（語意檢索 + 缺口比對）
  ↓
結果頁：選擇匹配職能基準、瀏覽任務/知識/技能、查看缺口摘要
  ↓
（選擇性）修改 5W2H → 重新分析
  ↓
勾選確認 → 匯出 Excel
```

---

## 與 Graph_RAG_test 整合

若系統已有運行中的 `GraphRAGQueryEngine` 實例，可直接傳入共用，避免重複載入 Embedding 模型：

```python
from wizard_ui import WizardMainWindow
from PyQt6.QtWidgets import QApplication

app = QApplication([])
win = WizardMainWindow(engine=your_engine)  # 傳入已載入的 engine
win.show()
app.exec()
```

---

## 授權

MIT License
