"""
competency_wizard/logger.py
集中式日誌設定 — 在 main.py 呼叫 setup() 一次後全域生效
"""
import logging
import logging.handlers
import sys
from pathlib import Path

_LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
_initialized = False


def setup(log_dir: Path = None, level: int = logging.INFO) -> None:
    """初始化 logging，輸出至 console 與（選擇性）旋轉檔案。
    多次呼叫無副作用。
    """
    global _initialized
    if _initialized:
        return
    _initialized = True

    root = logging.getLogger()
    root.setLevel(level)

    # Console handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(level)
    ch.setFormatter(logging.Formatter(_LOG_FORMAT, _DATE_FORMAT))
    root.addHandler(ch)

    # File handler（選擇性）
    if log_dir is not None:
        log_dir = Path(log_dir)
        log_dir.mkdir(parents=True, exist_ok=True)
        fh = logging.handlers.RotatingFileHandler(
            log_dir / "wizard.log",
            maxBytes=5 * 1024 * 1024,   # 5 MB per file
            backupCount=3,
            encoding="utf-8",
        )
        fh.setLevel(level)
        fh.setFormatter(logging.Formatter(_LOG_FORMAT, _DATE_FORMAT))
        root.addHandler(fh)


def get(name: str) -> logging.Logger:
    """取得具名 logger。setup() 可在之後再呼叫。"""
    return logging.getLogger(name)
