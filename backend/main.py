"""
Backend Penjadwalan Produksi Bengkel (FastAPI).

Satu service ini melakukan DUA hal:
  1. API  : /api/methods, /api/schedule  -> memanggil final_colab_scheduler.py (kode Dinda)
  2. Web  : melayani ../frontend/index.html di alamat "/"

Jalankan lokal:
    pip install -r requirements.txt
    uvicorn main:app --reload --port 8000
    -> buka http://localhost:8000

Environment variable (semuanya opsional):
    APP_USER, APP_PASSWORD   kalau diisi, seluruh situs minta login (HTTP Basic)
    MAX_UPLOAD_MB            batas ukuran Excel (default 15)
    ALLOWED_ORIGINS          pisahkan koma; hanya perlu jika frontend di domain lain
"""
from __future__ import annotations

import base64
import hmac
import json
import math
import os
import tempfile
import threading
import time
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, Response

import final_colab_scheduler as sched
import storage

# Kode Dinda mengunci jumlah operasi = 694 (cocok untuk data skripsi). Untuk aplikasi,
# kunci itu dimatikan supaya Excel dengan jumlah SO/operasi lain tetap bisa diproses.
sched.EXPECTED_OPERATION_COUNT = None

BASE = Path(__file__).resolve().parent
FRONTEND = BASE.parent / "frontend"
MAX_UPLOAD = int(os.getenv("MAX_UPLOAD_MB", "15")) * 1024 * 1024
APP_USER = os.getenv("APP_USER", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "")

METHODS = {
    "EDD": "EDD + ECT (due date terdekat dulu)",
    "FCFS": "FCFS + ECT (siapa datang duluan)",
    "SPT": "SPT + ECT (operasi tercepat dulu)",
    "CR": "CR + ECT (critical ratio)",
    "GA": "Genetic Algorithm (1 seed, ~10-60 detik)",
}

# Metode yang dijalankan app. Dinda masih memfinalkan metode -> atur lewat environment variable:
#   SCHEDULE_METHOD = EDD | FCFS | SPT | CR | GA   (default EDD)
#   ALLOW_METHOD_CHOICE = 1  -> tampilkan pilihan metode di web (mode uji). Atau buka alamat dengan ?uji=1
DEFAULT_METHOD = os.getenv("SCHEDULE_METHOD", "EDD").upper()
if DEFAULT_METHOD not in METHODS:
    DEFAULT_METHOD = "EDD"
ALLOW_CHOICE = os.getenv("ALLOW_METHOD_CHOICE", "0") == "1"

app = FastAPI(title="Penjadwalan Produksi Bengkel")
storage.init()

origins = [o.strip() for o in os.getenv("ALLOWED_ORIGINS", "").split(",") if o.strip()]
if origins:
    app.add_middleware(CORSMiddleware, allow_origins=origins, allow_methods=["*"], allow_headers=["*"])

# GA + numba berat di CPU: izinkan satu penjadwalan pada satu waktu supaya server kecil tidak tumbang.
_run_lock = threading.Semaphore(1)


@app.middleware("http")
async def basic_auth(request: Request, call_next):
    if APP_USER and APP_PASSWORD and request.url.path != "/api/health":
        ok = False
        h = request.headers.get("authorization", "")
        if h.lower().startswith("basic "):
            try:
                u, _, p = base64.b64decode(h[6:]).decode().partition(":")
                ok = hmac.compare_digest(u, APP_USER) and hmac.compare_digest(p, APP_PASSWORD)
            except Exception:
                ok = False
        if not ok:
            return Response("Login diperlukan", 401, headers={"WWW-Authenticate": 'Basic realm="Jadwal Produksi"'})
    return await call_next(request)


def _clean(df: pd.DataFrame) -> list[dict]:
    """DataFrame -> list dict yang aman untuk JSON (tanggal ISO, NaN -> null)."""
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = out[c].dt.strftime("%Y-%m-%dT%H:%M:%S")
    out = out.astype(object).where(pd.notna(out), None)
    rows = out.to_dict(orient="records")
    for r in rows:
        for k, v in r.items():
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                r[k] = None
            elif hasattr(v, "item"):  # numpy scalar
                r[k] = v.item()
    return rows


def _utilization(P, ops: pd.DataFrame) -> list[dict]:
    """Utilisasi mesin internal = menit proses / (unit x menit efektif hari kerja selama jadwal berjalan)."""
    internal = ops[ops["Resource_Type"] == "Internal"]
    if internal.empty:
        return []
    t0, t1 = internal["Start"].min().normalize(), internal["Finish"].max().normalize()
    cal = P["CAL"].drop_duplicates("Date")
    cal = cal[(cal["Is_Working_Day"] == 1) & (cal["Date"] >= t0) & (cal["Date"] <= t1)]
    avail_per_unit = float(cal["Effective_Minutes"].sum())
    units = dict(zip(P["RES"]["Resource"], P["RES"]["Capacity_Units"]))
    busy = internal.groupby("Resource")["Duration_Productive_Min"].sum()
    res = []
    for name, b in busy.items():
        cap = float(units.get(name, 1) or 1) * avail_per_unit
        res.append({"resource": name, "busy_min": float(b), "avail_min": cap,
                    "pct": round(100 * float(b) / cap, 1) if cap else 0.0})
    return sorted(res, key=lambda x: -x["pct"])


def _apply_overrides(src: Path, dst: Path, workers: dict, machines: dict) -> None:
    """Tulis ulang Excel dengan jumlah pekerja / unit mesin yang diubah user (patokan Excel dibiarkan untuk yang tidak diubah)."""
    x = pd.ExcelFile(src)
    sheets = {n: pd.read_excel(x, n) for n in x.sheet_names}
    hc, res = sheets["Human_Capacity"], sheets["Resources"]
    for g, v in workers.items():
        if g not in sched.HUMAN_GROUPS:
            raise HTTPException(422, f"Grup pekerja '{g}' tidak dikenal. Pilihan: {sched.HUMAN_GROUPS}")
        if not (1 <= int(v) <= 99):
            raise HTTPException(422, f"Jumlah pekerja {g} harus 1-99")
        hc[g] = int(v)  # berlaku untuk semua periode bulan di sheet Human_Capacity
    names = set(res["Resource"].astype(str))
    unknown = [m for m in machines if m not in names]
    if unknown:
        raise HTTPException(422, f"Mesin tidak ada di sheet Resources: {unknown}")
    for m, v in machines.items():
        if not (1 <= int(v) <= 20):
            raise HTTPException(422, f"Jumlah unit {m} harus 1-20")
        res.loc[res["Resource"].astype(str) == m, "Capacity_Units"] = int(v)
    with pd.ExcelWriter(dst) as w:
        for n, d in sheets.items():
            d.to_excel(w, sheet_name=n, index=False)


@app.get("/api/health")
def health():
    return {"ok": True}


@app.get("/api/config")
def config():
    return {"default_method": DEFAULT_METHOD, "method_label": METHODS[DEFAULT_METHOD],
            "allow_choice": ALLOW_CHOICE, "methods": [{"id": k, "label": v} for k, v in METHODS.items()]}


@app.get("/api/runs")
def runs():
    return {"runs": storage.list_runs()}


@app.get("/api/runs/{run_id}")
def run_detail(run_id: str):
    r = storage.get_run(run_id)
    if r is None:
        raise HTTPException(404, "Riwayat tidak ditemukan")
    return JSONResponse(r)


@app.get("/api/runs/{run_id}/input")
def run_input(run_id: str):
    r = storage.get_input(run_id)
    if r is None:
        raise HTTPException(404, "Riwayat tidak ditemukan")
    data, name = r
    return Response(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{(name or "input.xlsx").replace(chr(34), "")}"'})


@app.delete("/api/runs/{run_id}")
def run_delete(run_id: str):
    if not storage.delete_run(run_id):
        raise HTTPException(404, "Riwayat tidak ditemukan")
    return {"deleted": run_id}


@app.get("/api/methods")
def methods():
    return {"methods": [{"id": k, "label": v} for k, v in METHODS.items()]}


@app.post("/api/schedule")
def schedule(
    file: UploadFile = File(...),
    method: str = Form(""),
    label: str = Form(""),
    due_time: str = Form("17:00"),
    seed: int = Form(101),
    generations: int = Form(70),
    workers: str = Form(""),    # JSON, contoh {"Mech":14,"QC":3}; kosong = pakai patokan Excel
    machines: str = Form(""),   # JSON, contoh {"CNC 1":2}
):
    method = (method or DEFAULT_METHOD).upper()
    if method not in METHODS:
        raise HTTPException(400, f"Metode '{method}' tidak dikenal. Pilihan: {list(METHODS)}")
    if not (file.filename or "").lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "File harus .xlsx / .xlsm")
    data = file.file.read(MAX_UPLOAD + 1)
    if len(data) > MAX_UPLOAD:
        raise HTTPException(413, f"File terlalu besar (maks {MAX_UPLOAD // 1024 // 1024} MB)")
    generations = max(10, min(int(generations), 300))
    try:
        w_over = json.loads(workers) if workers.strip() else {}
        m_over = json.loads(machines) if machines.strip() else {}
        assert isinstance(w_over, dict) and isinstance(m_over, dict)
    except Exception:
        raise HTTPException(400, "Format perubahan pekerja/mesin tidak valid")

    if not _run_lock.acquire(timeout=120):
        raise HTTPException(503, "Server sedang menjalankan penjadwalan lain, coba lagi sebentar.")
    try:
        with tempfile.TemporaryDirectory() as td:
            xlsx = Path(td) / "input.xlsx"
            xlsx.write_bytes(data)
            if w_over or m_over:
                try:
                    xlsx2 = Path(td) / "input_override.xlsx"
                    _apply_overrides(xlsx, xlsx2, w_over, m_over)
                    xlsx = xlsx2
                except HTTPException:
                    raise
                except Exception as e:
                    raise HTTPException(422, f"Gagal menerapkan perubahan kapasitas: {e}")
            try:
                P = sched.load_problem(str(xlsx), due_time=due_time)
            except Exception as e:  # kolom/sheet kurang, format salah, dll.
                raise HTTPException(422, f"Excel tidak sesuai format 7 sheet penjadwalan: {e}")

            t0 = time.perf_counter()
            try:
                if method == "GA":
                    params = dict(sched.BASELINE_PARAMS, generations=generations)
                    res = sched.run_ga_checkpointed(P, Path(td), "api", seed=int(seed), params=params)
                    ops = sched.ga_schedule_dataframe(P, res, method="GA")
                    so = sched.build_so_summary(P, res["finish"], "GA")
                    nlate, tot, mx, mk = res["fit"]
                else:
                    res = sched.run_heuristic(P, method)
                    ops, so = res["schedule"], res["so_summary"]
                    nlate, tot, mx, mk = res["metrics"]
            except HTTPException:
                raise
            except Exception as e:
                raise HTTPException(500, f"Penjadwalan gagal: {e}")
            runtime = time.perf_counter() - t0

            extra = [c for c in P["SO"].columns if c not in ("SO_ID", "Release_Date", "Due_Date")
                     and P["SO"][c].notna().any()][:8]
            payload = {
                "run_id": storage.new_id(), "created_at": storage.now(),
                "label": label.strip()[:80], "filename": file.filename,
                "method": method,
                "method_label": METHODS[method],
                "runtime_sec": round(runtime, 2),
                "metrics": {"n_late": int(nlate), "total_tardiness_min": int(tot),
                            "max_tardiness_min": int(mx), "makespan_min": int(mk),
                            "n_so": int(len(so)), "n_operations": int(len(ops))},
                "overrides": {"workers": w_over, "machines": m_over},
                "so_summary": _clean(so),
                "utilization": _utilization(P, ops),
                "so_extra_cols": extra,
                "so_extra": {str(r["SO_ID"]): r for r in _clean(P["SO"][["SO_ID"] + extra])},
                "operations": _clean(ops),
            }
            storage.save_run(payload, data)
            return JSONResponse(payload)
    finally:
        _run_lock.release()


# ---- Frontend statis (satu service = satu deploy) ----
@app.get("/")
def index():
    return FileResponse(FRONTEND / "index.html")
