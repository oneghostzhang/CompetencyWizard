# CompetencyWizard 論文草稿

> **狀態**：草稿整理中　　**最後更新**：2026-04-03  
> ⚠️ 標記處為待補充內容（需實驗數據）

---

## 1. 摘要（Abstract）

職能說明書（Competency Profile）是企業人才管理的核心文件，傳統人工對照 ICAP 職能基準書逐欄填寫往往耗費 HR 專家 4–8 小時以上，且評審人員的主觀判斷難以標準化。本研究提出 **CompetencyWizard**，一個以 RAG（Retrieval-Augmented Generation）為核心的職能說明書半自動化製作系統。

系統以 BAAI/bge-base-zh-v1.5 Embedding 模型建立 FAISS 向量索引，從 908 份 ICAP 職能基準 JSON 中語義檢索最相似基準；並以 LlamaCpp 驅動 TAIDE-LX-7B-Chat 本地大型語言模型，依員工填寫的工作任務描述自動生成符合 ICAP 格式的行為指標。整合 PyQt6 桌面介面提供六頁引導式流程，員工可直接編輯生成結果並匯出標準 Excel 職能說明書（含職能說明書、知識清單、技能清單、態度清單、補充說明五個 Sheet）。

**主要貢獻：**
- 提出可落地部署的 RAG + LLM 職能說明書生成流水線
- 子 process 隔離機制防止 llama.cpp C-level abort 崩潰主程式
- 完整的 ICAP 格式對齊（主責 / 任務 / 知識 / 技能 / 態度五大結構）
- ⚠️ **待補充**：量化評估結果（生成品質分數、節省時間比例等）

---

## 2. 緒論（Introduction）

### 2.1 研究背景與動機

勞動部職能發展應用平台（ICAP）已累積超過 900 份各職業職能基準書，提供標準化的職能框架供企業參考。然而企業實際採用率偏低，主要原因如下：

| 面向 | 現況問題 |
|------|---------|
| 人工成本 | 一份職能說明書需 HR 專家 4–8 小時人工對照填寫 |
| 格式複雜 | ICAP 格式有特定句型（能＋動詞＋受詞＋條件），非專業人員難以遵循 |
| 一致性 | 不同評審人員對同一職位的行為指標描述差異大 |
| 查找困難 | 900+ 份基準需人工逐一比對，無語義搜尋支援 |

### 2.2 問題陳述

如何讓企業員工無需 HR 專業背景，透過人機協作方式，在合理時間內產出符合 ICAP 標準的職能說明書？

### 2.3 本文貢獻

1. 以語義向量檢索自動匹配最相近職能基準，取代人工查找
2. 六頁引導式流程降低使用門檻，任何員工均可操作
3. 本地 LLM 生成行為指標，保持資料完全離線不外傳
4. PDF 解析流水線 + 向量索引快取，支援批量職能基準入庫
5. 子 process 隔離機制解決 llama.cpp C-level 崩潰問題
6. ⚠️ **待補充**：用戶實驗結果數據

### 2.4 論文架構

本文結構如下：第 3 節回顧相關研究；第 4 節說明系統設計與架構；第 5 節描述實作細節；第 6 節報告實驗與評估結果；第 7 節討論系統限制與未來工作；第 8 節為結論。

---

## 3. 相關研究（Related Work）

### 3.1 職能管理系統

職能模型（Competency Model）由 Spencer & Spencer（1993）奠定理論基礎，定義職能為「能夠預測工作績效的個人特質」，並區分知識、技能、態度三大構面。現有商業 HR 系統（如 SAP SuccessFactors、Workday）雖支援職能管理，但均為雲端 SaaS 服務，不含台灣 ICAP 標準，且需將員工資料上傳外部伺服器。

### 3.2 Retrieval-Augmented Generation（RAG）

RAG 由 Lewis et al.（2020）提出，將大型語言模型與非參數式知識檢索結合，有效解決 LLM 知識截止日期與幻覺問題。本系統採用類似架構，以職能基準 JSON 作為知識庫，透過向量檢索提供 LLM 生成行為指標所需的領域知識。

### 3.3 中文 Sentence Embedding

Xiao et al.（2023）提出的 BGE（BAAI General Embedding）系列模型，在中文語義檢索任務上表現優異。本系統採用 bge-base-zh-v1.5，在保持合理推論速度的前提下，對職能基準的短文本匹配效果良好。

### 3.4 本系統與相關方法的差異

| 系統 / 方法 | 類型 | 主要限制 | 本系統優勢 |
|-----------|------|---------|-----------|
| 人工查表填寫 | 純人工 | 耗時、一致性差 | 自動檢索 + LLM 生成 |
| SAP SuccessFactors | 雲端 SaaS | 無 ICAP 基準、資料外傳 | 完全離線、ICAP 對齊 |
| ChatGPT / GPT-4 | 雲端 LLM | 不含 ICAP 知識、資料外傳 | RAG 注入基準、本地推論 |
| 問卷調查式系統 | Rule-based | 無法生成行為指標文字 | LLM 生成自然語言指標 |

---

## 4. 系統設計 / 架構（System Design）

### 4.1 系統架構圖

```
┌──────────────────────────────────────────────────────────┐
│                        資料輸入層                         │
│  ICAP PDF  ──→  pdf_parser_v2.py  ──→  parsed_json_v2/   │
└──────────────────────────┬───────────────────────────────┘
                           │
┌──────────────────────────▼───────────────────────────────┐
│                        向量索引層                         │
│  WizardRAG                                               │
│  ├── BAAI/bge-base-zh-v1.5  →  768 維向量               │
│  └── FAISS IndexFlatIP（餘弦相似度）                     │
│      快取：_index_cache/wizard_meta.pkl                  │
└──────────────────────────┬───────────────────────────────┘
                           │
┌──────────────────────────▼───────────────────────────────┐
│                        六頁 UI 流程（PyQt6）              │
│  Search → Editor → Detail → LLMSuggest → Supplement      │
│                               │                          │
│              analyze_tasks_batch()                       │
│              multiprocessing.Process（子 process 隔離）  │
│              LlamaCpp + TAIDE-LX-7B-Chat GGUF            │
└──────────────────────────┬───────────────────────────────┘
                           │
┌──────────────────────────▼───────────────────────────────┐
│                        輸出層                            │
│  excel_exporter.py  →  職能說明書.xlsx（5 Sheet）        │
│  職能說明書 / 知識清單 / 技能清單 / 態度清單 / 補充說明  │
└──────────────────────────────────────────────────────────┘
```

### 4.2 各模組功能說明

| 模組 | 功能 |
|------|------|
| `pdf_parser_v2.py` | 使用 pdfplumber 解析 ICAP PDF，狀態機擷取主責（T-code）、任務、知識、技能、態度，輸出結構化 JSON |
| `wizard_rag.py` | 載入 JSON → Embedding → FAISS 索引；提供 `search(query, top_k)` 介面；自動快取避免重建 |
| `wizard_ui.py` | PyQt6 六頁 QStackedWidget 主視窗；LLMAnalyzeThread 背景執行；DataManagerDialog 管理 PDF 資料 |
| `ai_chat.py` | 建立 LLM prompt；`analyze_tasks_batch()` 在隔離子 process 執行；`_split_indicators()` 解析輸出 |
| `excel_exporter.py` | 五個 Sheet 格式化輸出；`_collect_ks()` 彙整 K/S 對應任務；態度清單獨立 Sheet |
| `logger.py` | 每次啟動產生獨立時間戳記 log 檔，保留最近 30 份 |

### 4.3 技術選型理由

| 技術 | 選型理由 |
|------|---------|
| BAAI/bge-base-zh-v1.5 | 中文語意理解最優、開源可離線、768 維平衡效能與精度 |
| FAISS IndexFlatIP | 908 筆資料量不需近似搜尋，精確餘弦相似度保證檢索品質 |
| LlamaCpp + TAIDE-LX-7B | 繁中職場語境表現最優、完全離線、無 API 費用、GGUF 格式記憶體效率佳 |
| multiprocessing.Process | llama.cpp GGML_ASSERT abort 為 C-level，Python try/except 無法捕捉，需進程隔離 |
| PyQt6 | 跨平台桌面應用、原生控件、無需瀏覽器依賴 |
| pdfplumber | 表格結構擷取準確度優於 PyMuPDF，適合 ICAP PDF 格式 |

---

## 5. 實作細節（Implementation）

### 5.1 技術棧

| 層級 | 技術 | 版本需求 |
|------|------|---------|
| 桌面 UI | PyQt6 | ≥ 6.4 |
| 本地 LLM | llama-cpp-python + TAIDE GGUF | ≥ 0.2 |
| Embedding | sentence-transformers | ≥ 2.2 |
| 向量檢索 | faiss-cpu | ≥ 1.7 |
| Excel 輸出 | openpyxl | ≥ 3.1 |
| PDF 解析 | pdfplumber | ≥ 0.9 |
| 語言 | Python | ≥ 3.10 |

### 5.2 PDF 解析流水線

ICAP PDF 以表格形式排列職能基準，解析採狀態機逐行讀取：

```
PDF → pdfplumber 逐頁表格擷取
  → _parse_responsibilities_from_tables() 狀態機
      ├── 偵測 T-code（regex: r'T\d+'）→ 建立新主責記錄
      ├── 偵測任務代碼（regex: r'\d+\.\d+'）→ 建立新任務記錄
      └── 其餘行 → 累積至當前任務的知識 / 技能 / 態度欄位
  → 輸出 JSON（含 chunks_for_rag 欄位，供向量化使用）
```

**關鍵 Bug 修正**：原實作在偵測到新 T-code 時直接覆蓋 `current_resp`，未先儲存 `current_task`，導致每個主責只保留最後一筆任務。修正方式為在建立新主責前先執行儲存：

```python
if col_resp and re.match(r'T\d+', col_resp):
    # 修正：先儲存 current_task，再建立新主責
    if current_task and current_resp:
        current_resp["工作任務"].append(current_task)
        current_task = None
    if current_resp:
        responsibilities.append(current_resp)
    current_resp = { "代碼": ..., "名稱": ..., "工作任務": [] }
```

### 5.3 RAG 檢索流程

```python
# 查詢向量化 → FAISS 搜尋
query_vec = model.encode([query])           # shape: (1, 768)
scores, indices = index.search(query_vec, k=3)
matched_standards = [standards[i] for i in indices[0]]
```

向量快取機制：首次建立後序列化至 `_index_cache/wizard_meta.pkl`（含 chunks、embedding model 版本、index bytes）。為避免 JSON 更新後快取過期，每次載入快取後仍從 JSON 重新載入 `_standards` 字典，確保資料即時性。

### 5.4 LLM 行為指標生成

採 Llama2 chat template 建構 prompt：

```
[INST] <<SYS>>
你是一位職能基準書撰寫專家。請根據以下工作任務描述，
生成 2–3 條符合 ICAP 格式的行為指標。
格式要求：每條以「能」開頭，包含動詞、受詞與條件。
範例：能依據客戶需求，運用標準化流程完成服務，並確保品質達標。
<</SYS>>

職位：{position}　任務：{task_name}
描述：{description}　產出：{output} [/INST]
```

**子 process 隔離機制**：

```python
def analyze_tasks_batch(rows, position):
    q = multiprocessing.Queue()
    proc = multiprocessing.Process(
        target=_worker_main, args=(tasks, q)
    )
    proc.daemon = True
    proc.start()
    return proc, q

# UI 執行緒監控：
if not proc.is_alive():
    self.error.emit("AI 子程序意外中止")
```

**輸出後處理**：LLM 有時將多條指標合併為單一字串，`_split_indicators()` 以 `\n` 和 `指標N:` pattern 切分，確保每條獨立顯示於可編輯的 QLineEdit。

### 5.5 資料來源與處理

- **職能基準**：勞動部 ICAP 平台，共 908 份 PDF，全量解析後存為 JSON
- **向量維度**：768 dim（bge-base-zh-v1.5 輸出）
- **索引規模**：約 900 筆 chunks，FAISS 索引約 300 MB
- **模型大小**：TAIDE-LX-7B Q4_K_S 量化版約 4.5 GB

---

## 6. 實驗與評估（Experiments / Evaluation）

> ⚠️ **本節為最重要的待補充部分，以下為建議的實驗設計框架**

### 6.1 評估指標

**A. RAG 職能基準檢索準確度**

| 指標 | 說明 |
|------|------|
| Top-1 Accuracy | 第一個結果即為正確基準的比例 |
| Top-3 Recall | 正確基準出現在 Top-3 的比例 |
| MRR（Mean Reciprocal Rank） | 排名倒數平均值 |

**B. LLM 行為指標生成品質**

| 指標 | 說明 | 評估方式 |
|------|------|---------|
| ICAP 格式符合率 | 是否以「能＋動詞」開頭 | Regex 自動評估 |
| BERTScore | 與任務描述的語意相似度 | 自動計算 |
| 人工評分（1–5 分） | HR 專家評分相關性、專業性、可讀性 | 至少 3 位評審 |
| 採用率 | 使用者最終勾選 LLM 建議的比例 | 系統日誌統計 |

**C. 系統效能**

| 指標 | 目標值 |
|------|--------|
| 索引快取載入時間 | ≤ 5 秒 |
| RAG 單次查詢時間 | ≤ 3 秒 |
| LLM 行為指標生成（單任務） | 30–120 秒（CPU 推論） |
| Excel 匯出時間 | ≤ 1 秒 |

### 6.2 基線比較（Baseline Comparison）

| 系統 | Top-1 Acc. | 行為指標品質（1–5） | 完成時間（分鐘） |
|------|-----------|------------------|----------------|
| 人工填寫 | — | ⚠️ 待測 | ⚠️ 待測 |
| CompetencyWizard | ⚠️ 待測 | ⚠️ 待測 | ⚠️ 待測 |
| ChatGPT（無 RAG） | — | ⚠️ 待測 | ⚠️ 待測 |

### 6.3 消融實驗（Ablation Study）

| 系統變體 | 移除模組 | 預期影響 |
|---------|---------|---------|
| Full System | — | 基準線 |
| w/o RAG | 不注入職能基準 | 行為指標脫離 ICAP 格式、K/S/A 欄位空白 |
| w/o LLM | 只顯示基準預設指標 | 指標與實際工作描述脫節 |
| w/o cache | 每次重建 FAISS | 啟動時間增加 60–180 秒 |
| w/o subprocess | LLM 在主 process | GGML_ASSERT abort → 全程式崩潰 |

### 6.4 建議的使用者實驗設計

**受試者**：⚠️ 至少 10 位（建議 20 位），區分 HR 背景 vs 一般員工

**實驗流程**：
```
受試者隨機分配：
  A 組（控制組）：對照 ICAP PDF 人工填寫職能說明書
  B 組（實驗組）：使用 CompetencyWizard 製作

衡量指標：
  - 完成時間（分鐘）
  - 最終文件的 HR 專家評分（1–5 分）
  - 使用者滿意度問卷（System Usability Scale, SUS）
  - LLM 建議採用率（B 組）
```

---

## 7. 討論（Discussion）

### 7.1 系統限制

| 限制 | 說明 |
|------|------|
| 推論速度 | LlamaCpp CPU 推論約 30–120 秒 / 任務，GPU 加速待支援 |
| 模型路徑 | TAIDE GGUF 路徑硬編碼，非技術用戶難以更換模型 |
| PDF 解析相依性 | 依賴表格結構，格式不一致的 PDF 可能解析失敗 |
| 安裝門檻 | 需本機 Python 環境，無 Web 版本 |
| 語言限制 | 目前僅支援繁體中文 ICAP 基準 |

### 7.2 失敗案例分析

- **LLM 輸出格式不穩定**：多條指標合併為單一字串。已透過 `_split_indicators()` 後處理修正，但根本解法為 few-shot prompt 或 constrained decoding。
- **新興職業匹配度低**：如「直播主」、「NFT 創作者」等職業不在 ICAP 收錄範圍，RAG 僅能返回相似度偏低的近似結果。
- **口語化輸入**：職業名稱輸入過於口語（如「跑業務的」）時，向量匹配精度下降，需前處理標準化。
- **長任務描述截斷**：TAIDE 7B context window 限制，過長的任務描述可能被截斷，影響生成品質。

### 7.3 未來工作方向

1. **GPU 推論支援**：整合 llama-cpp-python CUDA build，預計推論加速 5–10 倍
2. **Web 版本**：FastAPI 後端 + Vue 前端，降低安裝門檻
3. **多基準合併匹配**：當職位跨多個 ICAP 基準時，支援合併多份基準內容
4. **使用者回饋微調**：收集使用者採用 / 拒絕的行為指標，作為 RLHF 訓練資料
5. **自動評估管線**：建立 BERTScore + 格式符合率的自動評估系統，持續監控生成品質
6. **多語言支援**：擴充英文職能基準（O*NET），支援跨語言匹配

---

## 8. 結論（Conclusion）

本研究提出 CompetencyWizard，結合語義向量檢索（RAG）與本地大型語言模型（TAIDE-LX-7B），實現職能說明書的半自動化製作。系統透過六頁引導式流程降低使用門檻，本地 LlamaCpp 推論確保資料完全離線，子 process 隔離機制解決 llama.cpp C-level abort 崩潰問題。從 908 份 ICAP 職能基準 PDF 解析建立向量索引，並對齊 ICAP 五大結構（主責 / 任務 / 知識 / 技能 / 態度）匯出標準 Excel 格式。

⚠️ **待補充**：根據實驗結果，系統能有效縮短職能說明書製作時間（XX%），並在 ICAP 格式符合率上達到 XX%，具備直接部署於企業環境的實用價值。

---

## 9. 參考文獻（References）

### RAG 與語言模型

[1] Lewis, P., Perez, E., Piktus, A., Petroni, F., Karpukhin, V., Goyal, N., ... & Kiela, D. (2020). **Retrieval-augmented generation for knowledge-intensive NLP tasks**. *Advances in Neural Information Processing Systems (NeurIPS)*, 33, 9459–9474.

[2] Gao, Y., Xiong, Y., Gao, X., Jia, K., Pan, J., Bi, Y., ... & Wang, H. (2023). **Retrieval-augmented generation for large language models: A survey**. *arXiv preprint arXiv:2312.10997*.

[3] Brown, T., Mann, B., Ryder, N., Subbiah, M., Kaplan, J. D., Dhariwal, P., ... & Amodei, D. (2020). **Language models are few-shot learners**. *Advances in Neural Information Processing Systems (NeurIPS)*, 33, 1877–1901.

[4] Touvron, H., Martin, L., Stone, K., Albert, P., Almahairi, A., Babaei, Y., ... & Scialom, T. (2023). **Llama 2: Open foundation and fine-tuned chat models**. *arXiv preprint arXiv:2307.09288*.

[5] Wei, J., Wang, X., Schuurmans, D., Bosma, M., Xia, F., Chi, E., ... & Zhou, D. (2022). **Chain-of-thought prompting elicits reasoning in large language models**. *Advances in Neural Information Processing Systems (NeurIPS)*, 35, 24824–24837.

[6] Kojima, T., Gu, S. S., Reid, M., Matsuo, Y., & Iwasawa, Y. (2022). **Large language models are zero-shot reasoners**. *Advances in Neural Information Processing Systems (NeurIPS)*, 35, 22199–22213.

### Embedding 與向量索引

[7] Xiao, S., Liu, Z., Zhang, P., & Muennighoff, N. (2023). **C-Pack: Packaged resources to advance general Chinese embedding**. *arXiv preprint arXiv:2309.07597*.  
（bge-base-zh-v1.5 原始論文）

[8] Reimers, N., & Gurevych, I. (2019). **Sentence-BERT: Sentence embeddings using Siamese BERT-networks**. In *Proceedings of the 2019 Conference on Empirical Methods in Natural Language Processing (EMNLP)*, 3982–3992.

[9] Johnson, J., Douze, M., & Jégou, H. (2021). **Billion-scale similarity search with GPUs**. *IEEE Transactions on Big Data*, 7(3), 535–547.  
（FAISS 原始論文）

### 職能管理理論

[10] Spencer, L. M., & Spencer, S. M. (1993). **Competence at work: Models for superior performance**. John Wiley & Sons.  
（Competency Framework 理論奠基，必引）

[11] Lucia, A. D., & Lepsinger, R. (1999). **The art and science of competency models: Pinpointing critical success factors in organizations**. Jossey-Bass/Pfeiffer.

### 人機協作

[12] Amershi, S., Chickering, M., Drucker, S. M., Lee, B., Simard, P., & Suh, J. (2019). **Software engineering for machine learning: A case study**. In *Proceedings of the 41st International Conference on Software Engineering: Software Engineering in Practice (ICSE-SEIP)*, 291–300.

---

## 附錄：待完成清單

| 優先度 | 項目 | 說明 |
|--------|------|------|
| 🔴 高 | RAG 檢索準確度實驗 | 手動標注 50 個職業對應正確 ICAP 基準，計算 Top-1 / Top-3 Accuracy / MRR |
| 🔴 高 | LLM 行為指標品質評估 | 找 3 位 HR / 相關領域人員評分（1–5 分），10 個測試案例 |
| 🔴 高 | 使用時間比較 | 人工填寫 vs 使用系統，各 5 個案例，記錄完成時間 |
| 🟡 中 | 消融實驗數據 | 實際執行 w/o RAG、w/o LLM 的結果對比 |
| 🟡 中 | 系統架構圖（向量圖） | 將第 4 節文字架構圖繪製為正式圖表（建議用 draw.io） |
| 🟢 低 | 英文翻譯 | 若投國際會議需翻譯全文 |
| 🟢 低 | 引用格式統一 | 依目標期刊 / 會議格式調整（APA / IEEE / ACM） |
