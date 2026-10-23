import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

from src.config import PROJECT_ROOT

DB_PATH = PROJECT_ROOT / "m365_audit.sqlite3"


class AuditDB:
    """Base de datos SQLite local para evitar procesar correos duplicados y auditar acciones."""

    def __init__(self, db_path: Path = DB_PATH) -> None:
        self.db_path = db_path
        self._init_db()

    def _init_db(self) -> None:
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS processed_emails (
                    message_id TEXT PRIMARY KEY,
                    sender_email TEXT,
                    subject TEXT,
                    category TEXT,
                    confidence REAL,
                    final_action TEXT,
                    reason TEXT,
                    draft_markdown TEXT,
                    processed_at TEXT
                )
                """
            )
            conn.commit()

    def is_processed(self, message_id: str) -> bool:
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute(
                "SELECT 1 FROM processed_emails WHERE message_id = ?",
                (message_id,),
            )
            return cur.fetchone() is not None

    def record_action(
        self,
        message_id: str,
        sender_email: str,
        subject: str,
        category: str,
        confidence: float,
        final_action: str,
        reason: str,
        draft_markdown: str,
    ) -> None:
        now_iso = datetime.now(timezone.utc).isoformat()
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT OR REPLACE INTO processed_emails (
                    message_id, sender_email, subject, category,
                    confidence, final_action, reason, draft_markdown, processed_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    message_id,
                    sender_email,
                    subject,
                    category,
                    confidence,
                    final_action,
                    reason,
                    draft_markdown,
                    now_iso,
                ),
            )
            conn.commit()

    def get_recent_logs(self, limit: int = 20) -> List[Dict[str, Any]]:
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.execute(
                "SELECT * FROM processed_emails ORDER BY processed_at DESC LIMIT ?",
                (limit,),
            )
            return [dict(row) for row in cur.fetchall()]
