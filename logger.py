"""
competency_wizard/logger.py
集中式日誌設定 — 在 main.py 呼叫 setup() 一次後全域生效
每次啟動產生一份獨立 LOG 檔（logs/YYYYMMDD_HHMMSS.log），
保留最近 30 份，超過自動刪除最舊的。
"""
import logging
import sys
from datetime import datetime
from pathlib import Path

_LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
_MAX_KEEP = 30          # 最多保留幾份 log
_initialized = False


def setup(log_dir: Path | None = None, level: int = logging.INFO) -> None:
    """初始化 logging，輸出至 console 與每次啟動獨立的 log 檔。
    多次呼叫無副作用。
    """
    global _initialized
    if _initialized:
        return
    _initialized = True

    root = logging.getLogger()
    root.setLevel(level)
    fmt = logging.Formatter(_LOG_FORMAT, _DATE_FORMAT)

    # Console handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(level)
    ch.setFormatter(fmt)
    root.addHandler(ch)

    # File handler：每次啟動一份新檔
    if log_dir is not None:
        log_dir = Path(log_dir)
        log_dir.mkdir(parents=True, exist_ok=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = log_dir / f"{ts}.log"

        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setLevel(level)
        fh.setFormatter(fmt)
        root.addHandler(fh)

        # 清理舊檔，只保留最近 _MAX_KEEP 份
        logs = sorted(log_dir.glob("*.log"), key=lambda p: p.stat().st_mtime)
        for old in logs[:-_MAX_KEEP]:
            try:
                old.unlink()
            except OSError:
                pass


def get(name: str) -> logging.Logger:
    """取得具名 logger。setup() 可在之後再呼叫。"""
    return logging.getLogger(name)
