"""Penyimpanan riwayat penjadwalan (SQLite). Satu file database + file Excel input asli per run.

Lokasi: env DATA_DIR (default ../storage). Di Railway/Render, arahkan DATA_DIR ke folder
yang di-mount sebagai Volume/Disk (misal /data), kalau tidak riwayat hilang tiap deploy ulang.
"""
from __future__ import annotations
import gzip, json, os, sqlite3, uuid
from datetime import datetime, timezone
from pathlib import Path

DATA_DIR = Path(os.getenv("DATA_DIR") or Path(__file__).resolve().parent.parent / "storage")
DB = DATA_DIR / "jadwal.db"


def _conn():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(DB, timeout=30)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")
    return c


def init():
    with _conn() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS runs(
            id TEXT PRIMARY KEY, created_at TEXT, label TEXT, filename TEXT, method TEXT,
            n_so INT, n_ops INT, n_late INT, makespan_min INT, runtime_sec REAL,
            overrides TEXT, result BLOB, input_xlsx BLOB)""")


def new_id() -> str:
    return uuid.uuid4().hex[:10]


def now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def save_run(payload: dict, xlsx_bytes: bytes):
    m = payload["metrics"]
    with _conn() as c:
        c.execute("INSERT INTO runs VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)", (
            payload["run_id"], payload["created_at"], payload.get("label", ""), payload.get("filename", ""),
            payload["method"], m["n_so"], m["n_operations"], m["n_late"], m["makespan_min"], payload["runtime_sec"],
            json.dumps(payload["overrides"]), gzip.compress(json.dumps(payload).encode()), gzip.compress(xlsx_bytes)))


def list_runs(limit: int = 200):
    with _conn() as c:
        rows = c.execute("""SELECT id,created_at,label,filename,method,n_so,n_ops,n_late,makespan_min,runtime_sec,overrides
                            FROM runs ORDER BY created_at DESC LIMIT ?""", (limit,)).fetchall()
    out = []
    for r in rows:
        d = dict(r); d["overrides"] = json.loads(d["overrides"] or "{}"); out.append(d)
    return out


def get_run(run_id: str):
    with _conn() as c:
        r = c.execute("SELECT result FROM runs WHERE id=?", (run_id,)).fetchone()
    return json.loads(gzip.decompress(r["result"])) if r else None


def get_input(run_id: str):
    with _conn() as c:
        r = c.execute("SELECT input_xlsx, filename FROM runs WHERE id=?", (run_id,)).fetchone()
    return (gzip.decompress(r["input_xlsx"]), r["filename"]) if r else None


def delete_run(run_id: str) -> bool:
    with _conn() as c:
        return c.execute("DELETE FROM runs WHERE id=?", (run_id,)).rowcount > 0
