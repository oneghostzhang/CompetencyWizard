"""
competency_wizard/wizard_rag.py
職能說明書精靈專用 RAG

可獨立使用（自行載入模型 + 建立索引），
也可傳入已初始化的 GraphRAGQueryEngine 直接復用其資源（避免重複載入 Embedding 模型）。
"""

import json
import logging
import pickle
import numpy as np
from pathlib import Path
from typing import Any, List, Dict, Optional, cast

try:
    import faiss
    from sentence_transformers import SentenceTransformer
    DEPS_AVAILABLE = True
except ImportError:
    DEPS_AVAILABLE = False


# 預設路徑
_DEFAULT_JSON_DIR = Path(__file__).parent / "data" / "parsed_json_v2"
_DEFAULT_INDEX_DIR = Path(__file__).parent / "_index_cache"
_EMBEDDING_MODEL = "BAAI/bge-base-zh-v1.5"

logger = logging.getLogger(__name__)


class WizardRAG:
    """職能說明書精靈專用 RAG

    傳入 engine（GraphRAGQueryEngine）時，直接復用其 embedding_model 與 vector_index，
    省去重新載入模型（約 1-2 分鐘）的等待時間。
    不傳入 engine 時，使用獨立的索引快取（_index_cache/）正常初始化。
    """

    def __init__(self, json_dir: Path | None = None, index_dir: Path | None = None, engine=None):
        self.json_dir = Path(json_dir) if json_dir else _DEFAULT_JSON_DIR
        self.index_dir = Path(index_dir) if index_dir else _DEFAULT_INDEX_DIR
        self.index_dir.mkdir(parents=True, exist_ok=True)

        # 可選：GraphRAGQueryEngine 實例，用於復用其向量資源
        self._engine = engine

        self._model: Optional[SentenceTransformer] = None
        self._index: Optional["faiss.Index"] = None
        self._chunks: List[Dict] = []   # [{"text", "standard_code", "standard_name", "chunk_type", "task_id"}]
        self._standards: Dict[str, Dict] = {}  # standard_code -> raw JSON dict

        self.initialized = False
        self._stop_requested = False

    @property
    def chunk_count(self) -> int:
        return len(self._chunks)

    @property
    def using_shared_engine(self) -> bool:
        return self._engine is not None

    # ─────────────────────────────────────────
    # 初始化
    # ─────────────────────────────────────────

    def stop(self) -> None:
        """請求中止初始化（用於取消長時間的背景操作）。"""
        self._stop_requested = True

    def initialize(self, progress_cb=None):
        """載入模型與向量索引。

        優先嘗試從 engine 復用（若有提供）；否則退回獨立快取/建立流程。
        """
        self._stop_requested = False
        logger.info("開始初始化 WizardRAG...")
        if not DEPS_AVAILABLE:
            logger.error("缺少依賴套件：faiss-cpu 或 sentence-transformers")
            raise RuntimeError("缺少依賴套件：faiss-cpu 或 sentence-transformers")

        if self._engine is not None and self._try_init_from_engine(progress_cb):
            self.initialized = True
            return

        # ── 獨立初始化（無 engine 或 engine 尚未就緒）──
        if progress_cb:
            progress_cb("載入 Embedding 模型...")
        self._model = SentenceTransformer(_EMBEDDING_MODEL)

        if self._try_load_cache():
            if progress_cb:
                progress_cb(f"已載入快取索引（{len(self._chunks)} 個 chunks）")
        else:
            if progress_cb:
                progress_cb("建立向量索引（首次啟動需要數分鐘）...")
            self._build_index(progress_cb)
            self._save_cache()

        self.initialized = True
        logger.info("WizardRAG 初始化完成（%d chunks，%d 個標準）", len(self._chunks), len(self._standards))

    # ─── 從 engine 復用 ─────────────────────────

    def _try_init_from_engine(self, progress_cb=None) -> bool:
        """嘗試從 GraphRAGQueryEngine 復用 embedding_model 與 vector_index。"""
        engine = self._engine
        if engine is None:
            return False

        if getattr(engine, 'embedding_model', None) is None:
            if progress_cb:
                progress_cb("主程式 Embedding 尚未就緒，改為獨立初始化...")
            return False

        if getattr(engine, 'vector_index', None) is None:
            if progress_cb:
                progress_cb("主程式向量索引尚未就緒，改為獨立初始化...")
            return False

        if progress_cb:
            progress_cb("復用主程式 Embedding 模型與向量索引...")

        # 直接引用（不複製），節省記憶體
        self._model = engine.embedding_model
        self._index = engine.vector_index

        # 從 chunk_meta_map 重組 _chunks（格式對齊 wizard 自建索引）
        chunk_meta_map: Dict[int, Dict] = getattr(engine, 'chunk_meta_map', {})
        n = engine.vector_index.ntotal
        self._chunks = []
        for i in range(n):
            cm = chunk_meta_map.get(i, {})
            self._chunks.append({
                "text": cm.get("chunk_content", ""),
                "standard_code": cm.get("standard_code", ""),
                "standard_name": cm.get("standard_name", ""),
                "standard_category": cm.get("standard_category", ""),
                "chunk_type": cm.get("chunk_type", ""),
                "task_id": cm.get("task_id", ""),
            })

        # 從原始 JSON 載入 standards（輕量操作，格式與 wizard 其他模組一致）
        if progress_cb:
            progress_cb("載入職能基準資料...")
        self._standards = self._load_standards_from_json()

        if progress_cb:
            progress_cb(
                f"已從主程式索引載入（{len(self._chunks)} chunks，{len(self._standards)} 個標準）"
            )
        return True

    # ─── 獨立快取邏輯 ────────────────────────────

    def _try_load_cache(self) -> bool:
        faiss_file = self.index_dir / "wizard.faiss"
        meta_file = self.index_dir / "wizard_meta.pkl"
        if not faiss_file.exists() or not meta_file.exists():
            return False
        try:
            self._index = faiss.read_index(str(faiss_file))
            with open(meta_file, "rb") as f:
                meta = pickle.load(f)
            self._chunks = meta["chunks"]
            if meta.get("embedding_model", "") != _EMBEDDING_MODEL:
                return False
            # 舊快取沒有 standard_category 欄位時強制重建
            if self._chunks and "standard_category" not in self._chunks[0]:
                return False
            # 每次都從 JSON 重新載入 standards，確保資料是最新的
            self._standards = self._load_standards_from_json()
            return True
        except Exception as e:
            logger.warning("快取載入失敗（%s），將重新建立索引", e)
            return False

    def _save_cache(self):
        faiss.write_index(self._index, str(self.index_dir / "wizard.faiss"))
        with open(self.index_dir / "wizard_meta.pkl", "wb") as f:
            pickle.dump({
                "chunks": self._chunks,
                "standards": self._standards,
                "embedding_model": _EMBEDDING_MODEL,
            }, f)

    def _build_index(self, progress_cb=None):
        """從 parsed_json_v2 JSON 檔案建立 FAISS 索引（獨立模式用）。"""
        if self._model is None:
            raise RuntimeError("Embedding 模型尚未初始化")
        json_files = list(self.json_dir.glob("*.json"))
        texts = []
        self._chunks = []

        for fpath in json_files:
            try:
                data = json.loads(fpath.read_text(encoding="utf-8"))
            except Exception as e:
                logger.warning("跳過無效 JSON 檔 %s：%s", fpath.name, e)
                continue

            std_code = (
                data.get("metadata", {}).get("code")
                or data.get("basic_info", {}).get("code", "")
            )
            std_name = (
                data.get("metadata", {}).get("name")
                or data.get("basic_info", {}).get("name", "")
            )
            std_category = data.get("basic_info", {}).get("category", "")
            if not std_code:
                continue

            self._standards[std_code] = data

            for chunk in data.get("chunks_for_rag", []):
                content = chunk.get("content", "").strip()
                if not content:
                    continue
                meta = chunk.get("metadata", {})
                ct = meta.get("chunk_type", "")
                if ct == "summary":
                    continue

                # task chunk 補上職能基準名稱，提升跨標準查詢精度
                if ct == "task" and std_name and std_name not in content:
                    content = f"職能基準: {std_name}（{std_code}）\n" + content

                if len(content) > 750:
                    content = content[:750]

                texts.append(content)
                self._chunks.append({
                    "text": content,
                    "standard_code": std_code,
                    "standard_name": std_name,
                    "standard_category": std_category,
                    "chunk_type": ct,
                    "task_id": meta.get("task_id", ""),
                })

        if not texts:
            raise RuntimeError("找不到可用的 chunks，請確認 json_dir 路徑正確")

        # 分批 encode，每批更新進度（避免 UI 看起來卡死）
        batch_size = 64
        total = len(texts)
        all_embeddings = []
        for start in range(0, total, batch_size):
            if self._stop_requested:
                raise RuntimeError("使用者已取消初始化")
            batch = texts[start:start + batch_size]
            batch_emb = self._model.encode(batch, show_progress_bar=False)
            all_embeddings.append(batch_emb)
            if progress_cb:
                done = min(start + batch_size, total)
                progress_cb(f"向量化中 {done}/{total} chunks...")

        embeddings = np.asarray(np.vstack(all_embeddings), dtype="float32")
        faiss.normalize_L2(embeddings)

        self._index = faiss.IndexFlatIP(embeddings.shape[1])
        cast(Any, self._index).add(embeddings)

    def _load_standards_from_json(self) -> Dict[str, Dict]:
        """從原始 JSON 檔載入職能基準資料（raw dict，格式與 gap_analyzer 預期一致）。"""
        standards = {}
        for fpath in self.json_dir.glob("*.json"):
            try:
                data = json.loads(fpath.read_text(encoding="utf-8"))
                std_code = (
                    data.get("metadata", {}).get("code")
                    or data.get("basic_info", {}).get("code", "")
                )
                if std_code:
                    standards[std_code] = data
            except Exception as e:
                logger.warning("跳過無效 JSON 檔 %s：%s", fpath.name, e)
                continue
        return standards

    # ─────────────────────────────────────────
    # 查詢
    # ─────────────────────────────────────────

    def search(self, query: str, top_k: int = 5) -> List[Dict]:
        """向量搜尋最相似的職能基準，回傳 top_k 個不重複標準。"""
        if not self.initialized:
            raise RuntimeError("WizardRAG 尚未初始化，請先呼叫 initialize()")
        if self._model is None or self._index is None:
            raise RuntimeError("WizardRAG 內部狀態未完成初始化")

        vec = self._model.encode([query], show_progress_bar=False)
        vec = np.array(vec, dtype="float32")
        faiss.normalize_L2(vec)

        scores, indices = cast(Any, self._index).search(vec, top_k * 4)

        seen_codes: set = set()
        results = []
        for score, idx in zip(scores[0], indices[0]):
            if idx < 0 or idx >= len(self._chunks):
                continue
            chunk = self._chunks[idx]
            code = chunk["standard_code"]
            if not code or code in seen_codes:
                continue
            seen_codes.add(code)
            results.append({
                "standard_code": code,
                "standard_name": chunk["standard_name"],
                "standard_category": chunk.get("standard_category", ""),
                "score": float(score),
                "chunk_type": chunk["chunk_type"],
                "matched_text": chunk["text"][:200],
            })
            if len(results) >= top_k:
                break
        return results

    def get_standard(self, code: str) -> Optional[Dict]:
        """取得職能基準完整資料（含任務、知識、技能）。"""
        return self._standards.get(code)

    def match_to_tasks(
        self,
        query: str,
        task_names: List[str],
        threshold: float = 0.55,
    ) -> tuple:
        """Layer 2：將一段員工任務描述與標準任務名稱清單做語意相似度比對。

        回傳 (best_index, best_score)。
        若最高分未達 threshold，best_index 為 -1。
        """
        if not task_names or not self.initialized or not self._model:
            return -1, 0.0

        texts = [query] + task_names
        embs = np.asarray(self._model.encode(texts, show_progress_bar=False), dtype="float32")
        faiss.normalize_L2(embs)

        # 查詢向量 vs. 各任務向量的點積（= 正規化後的餘弦相似度）
        scores = embs[1:] @ embs[0]
        best_idx   = int(np.argmax(scores))
        best_score = float(scores[best_idx])
        return (best_idx, best_score) if best_score >= threshold else (-1, best_score)

    def invalidate_cache(self):
        """刪除獨立快取，下次初始化時強制重建。

        若使用 engine 模式，快取不適用，但仍重置 initialized 狀態。
        """
        if self._engine is None:
            for f in self.index_dir.glob("wizard.*"):
                f.unlink(missing_ok=True)
        self.initialized = False
