"""
qb_queue.py — Shared SQLite job queue between Flask app and QB SOAP server.

The Flask app WRITES jobs here when a user pushes an order to QB.
The SOAP server READS pending jobs and returns them to QBWC as QBXML.
QBWC delivers QB's response back to the SOAP server, which UPDATES the job.
"""

import sqlite3
import json
import uuid
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "qb_jobs.db"


def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Create the jobs table if it doesn't exist. Call once on startup."""
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS qb_jobs (
                id          TEXT PRIMARY KEY,
                type        TEXT NOT NULL,
                payload     TEXT NOT NULL,
                status      TEXT NOT NULL DEFAULT 'pending',
                result      TEXT,
                error       TEXT,
                created_at  TEXT NOT NULL,
                updated_at  TEXT NOT NULL
            )
        """)
        conn.commit()


def enqueue(job_type: str, payload: dict) -> str:
    """
    Add a new job to the queue. Returns the job ID.

    job_type: 'sales_order' | 'estimate' | 'inventory_query'
    payload:  dict with all data needed to build the QBXML request
    """
    job_id = str(uuid.uuid4())
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        conn.execute(
            """INSERT INTO qb_jobs (id, type, payload, status, created_at, updated_at)
               VALUES (?, ?, ?, 'pending', ?, ?)""",
            (job_id, job_type, json.dumps(payload), now, now)
        )
        conn.commit()
    return job_id


def get_next_pending() -> dict | None:
    """
    Return the oldest pending job and mark it as 'processing'.
    Returns None if queue is empty.
    Called by the SOAP server when QBWC asks for work.
    """
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        row = conn.execute(
            "SELECT * FROM qb_jobs WHERE status = 'pending' ORDER BY created_at ASC LIMIT 1"
        ).fetchone()
        if row is None:
            return None
        conn.execute(
            "UPDATE qb_jobs SET status = 'processing', updated_at = ? WHERE id = ?",
            (now, row["id"])
        )
        conn.commit()
        return dict(row)


def complete_job(job_id: str, result: dict):
    """Mark a job as done and store QB's response."""
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        conn.execute(
            "UPDATE qb_jobs SET status = 'done', result = ?, updated_at = ? WHERE id = ?",
            (json.dumps(result), now, job_id)
        )
        conn.commit()


def fail_job(job_id: str, error: str):
    """Mark a job as failed and store the error message."""
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        conn.execute(
            "UPDATE qb_jobs SET status = 'error', error = ?, updated_at = ? WHERE id = ?",
            (error, now, job_id)
        )
        conn.commit()


def get_job(job_id: str) -> dict | None:
    """Fetch a single job by ID. Used by Flask to poll for results."""
    with _conn() as conn:
        row = conn.execute(
            "SELECT * FROM qb_jobs WHERE id = ?", (job_id,)
        ).fetchone()
        return dict(row) if row else None


def get_recent_jobs(limit: int = 50) -> list[dict]:
    """Return the most recent jobs for the admin status view."""
    with _conn() as conn:
        rows = conn.execute(
            "SELECT * FROM qb_jobs ORDER BY created_at DESC LIMIT ?", (limit,)
        ).fetchall()
        return [dict(r) for r in rows]
