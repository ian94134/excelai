"""
結構化 JSON Lines 日誌。

設計原則（參見 PLAN_v4.md §C-1、SECURITY.md）：
- 一行一筆 JSON，方便 grep / jq / 匯入資料庫
- 輸出位置固定為 `~/.excel-ai/logs/{YYYY-MM-DD}.jsonl`，與 Telemetry 的 SQLite 同目錄
- 敏感參數用 SHA-256 前 8 碼 hash 記錄，不寫原值
- API key / token 永遠不寫 log
- 寫入失敗（如 home 無寫入權）自動退化為 stderr，不阻斷主流程

V4 Phase 1 引入。Phase 2/3 的 backup / retry / telemetry 會共用此 logger。
"""

import hashlib
import json
import logging
import os
import sys
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path


# ── 常數 ──────────────────────────────────────────────────────────────────────

LOG_DIR = Path.home() / ".excel-ai" / "logs"
LOG_FILENAME_PATTERN = "%Y-%m-%d"  # 每日一檔
MAX_BYTES = 5 * 1024 * 1024  # 單檔 5MB 後輪替
BACKUP_COUNT = 5             # 保留 5 份輪替檔

# 環境變數 EXCEL_AI_LOG_LEVEL 可覆寫；預設 INFO
_LEVEL = getattr(logging, os.getenv("EXCEL_AI_LOG_LEVEL", "INFO").upper(), logging.INFO)


# ── JSON Formatter ───────────────────────────────────────────────────────────

class JsonFormatter(logging.Formatter):
    """輸出結構化 JSON，每筆一行。"""

    def format(self, record: logging.LogRecord) -> str:
        payload: dict = {
            "ts": datetime.now(timezone.utc).isoformat(),
            "level": record.levelname,
            "logger": record.name,
            "msg": record.getMessage(),
        }
        # 使用者透過 logger.info(..., extra={...}) 傳入的額外欄位
        for key, value in record.__dict__.items():
            if key in {
                "name", "msg", "args", "levelname", "levelno", "pathname",
                "filename", "module", "exc_info", "exc_text", "stack_info",
                "lineno", "funcName", "created", "msecs", "relativeCreated",
                "thread", "threadName", "processName", "process", "message",
                "taskName",
            }:
                continue
            payload[key] = value
        # 例外 traceback（若有）
        if record.exc_info:
            payload["exc"] = self.formatException(record.exc_info)
        return json.dumps(payload, ensure_ascii=False, default=str)


# ── 建立 logger ──────────────────────────────────────────────────────────────

_INITIALIZED: set[str] = set()


def _ensure_log_dir() -> Path | None:
    """嘗試建立 log 目錄，失敗回傳 None（呼叫端會降級為 stderr）。"""
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        return LOG_DIR
    except OSError:
        return None


def get_logger(name: str) -> logging.Logger:
    """
    取得已設定 JsonFormatter 的 logger。重複呼叫同一 name 不會重複加 handler。

    用法：
        log = get_logger("executor")
        log.info("tool_executed", extra={"tool": "read_range", "duration_ms": 42})
    """
    logger = logging.getLogger(f"excel_ai.{name}")

    if name in _INITIALIZED:
        return logger

    logger.setLevel(_LEVEL)
    logger.propagate = False  # 不往 root 傳，避免重複輸出

    formatter = JsonFormatter()

    # 檔案 handler
    log_dir = _ensure_log_dir()
    if log_dir is not None:
        today = datetime.now().strftime(LOG_FILENAME_PATTERN)
        file_path = log_dir / f"{today}.jsonl"
        try:
            fh = RotatingFileHandler(
                file_path, maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT,
                encoding="utf-8",
            )
            fh.setFormatter(formatter)
            logger.addHandler(fh)
        except OSError:
            pass  # 寫入失敗時退化為純 stderr

    # Stderr handler：debug 模式或檔案 handler 無法建立時仍有輸出
    if os.getenv("EXCEL_AI_LOG_STDERR", "").lower() in ("1", "true", "yes") or not logger.handlers:
        sh = logging.StreamHandler(sys.stderr)
        sh.setFormatter(formatter)
        logger.addHandler(sh)

    _INITIALIZED.add(name)
    return logger


# ── 敏感資料處理 ──────────────────────────────────────────────────────────────

def hash_args(arguments: dict) -> str:
    """
    把 tool arguments 轉為穩定的 8 碼 hash。
    相同內容永遠產生相同 hash，方便在 log 裡比對「是否為同一種呼叫」。
    """
    try:
        canonical = json.dumps(arguments, ensure_ascii=False, sort_keys=True, default=str)
    except (TypeError, ValueError):
        canonical = str(arguments)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:8]


def redact_prompt(text: str, max_len: int = 100) -> str:
    """把使用者 prompt 截為前 100 字；完整內容不進 log。"""
    if not text:
        return ""
    text = text.replace("\n", " ").strip()
    return text[:max_len] + ("…" if len(text) > max_len else "")
