from __future__ import annotations

import argparse
import gzip
import json
import math
import os
import pickle
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Any

import numpy as np
import pandas as pd
from numba import njit

# ============================================================
# FINAL SCHEDULING PIPELINE - PT XYZ
# DRCFJSP + branching precedence + pooled human capacity
# Methods: FCFS+ECT, EDD+ECT, SPT+ECT, CR+ECT, Genetic Algorithm
# ============================================================
# Important modelling rules used by this script:
# 1) Service Order = job; operation = scheduling entity.
# 2) Ready set is precedence-based.
# 3) Internal operation may pause over nonproductive periods / end-of-day,
#    but resumes on the SAME physical resource.
# 4) Physical + human capacity must be feasible simultaneously.
# 5) Component lead time starts after the SO incoming-inspection operation.
#    It gates operations belonging to that component; precedence propagates it.
# 6) Vendor duration uses LT_Plan_Days (calendar days), not Total_Time.
#    Vendor does not consume internal physical/human capacity.
# 7) Final Assembly is NOT given an artificial max-component-LT gate;
#    it waits through its actual precedence relations.
# 8) Due_Date in the workbook is date-only. This script treats the contractual
#    due timestamp as DUE_TIME on the due date. Default = 17:00 (end of shift).
# 9) Makespan is reported as elapsed calendar minutes from earliest release
#    to the latest Service Order Final Assembly completion.
# ============================================================

DUE_TIME_DEFAULT = "17:00"
EXPECTED_OPERATION_COUNT = 694
HUMAN_GROUPS = ["Mech", "Mach", "Welder", "QC"]
FINAL_SEEDS = list(range(101, 121))
TUNING_SEEDS = [41, 42, 43]

BASELINE_PARAMS = {
    "population_size": 50,
    "generations": 70,
    "pc": 0.90,
    "pm_seq": 0.20,
    "pm_res": 0.10,
    "elite": 2,
    "tournament": 3,
}

RATE_SETS = {
    "A": {"pc": 0.85, "pm_seq": 0.15, "pm_res": 0.05},
    "B": {"pc": 0.90, "pm_seq": 0.20, "pm_res": 0.10},
    "C": {"pc": 0.95, "pm_seq": 0.25, "pm_res": 0.15},
}


# ---------------------------
# Utility I/O
# ---------------------------

def atomic_pickle_gz(obj, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with gzip.open(tmp, "wb") as f:
        pickle.dump(obj, f, protocol=pickle.HIGHEST_PROTOCOL)
    os.replace(tmp, path)


def load_pickle_gz(path: Path):
    with gzip.open(path, "rb") as f:
        return pickle.load(f)


def atomic_csv(df: pd.DataFrame, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    df.to_csv(tmp, index=False)
    os.replace(tmp, path)


def write_json(obj, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2, ensure_ascii=False, default=str)
    os.replace(tmp, path)


def timestamp_min(ts) -> int:
    return int(pd.Timestamp(ts).value // 60_000_000_000)


def min_to_timestamp(x: int) -> pd.Timestamp:
    return pd.Timestamp(int(x), unit="m")


def hm_to_min(s: str) -> int:
    h, m = map(int, str(s).split(":"))
    return h * 60 + m


def parse_due_timestamp(date_value, due_time: str) -> pd.Timestamp:
    h, m = map(int, due_time.split(":"))
    return pd.Timestamp(date_value).normalize() + pd.Timedelta(hours=h, minutes=m)


def fit_tuple_to_key(fit):
    return tuple(float(x) for x in fit)


# ---------------------------
# Calendar and problem loader
# ---------------------------

def build_productive_intervals(calendar_df: pd.DataFrame) -> List[Tuple[int, int]]:
    required = ["Activity", "Start", "End", "Available_for_Process"]
    for c in required:
        if c not in calendar_df.columns:
            raise ValueError(f"Calendar missing required column: {c}")

    act = calendar_df[required].dropna(subset=["Activity"]).copy()
    shift_rows = act[act["Activity"].astype(str).str.strip().str.lower().eq("shift")]
    if len(shift_rows) != 1:
        raise ValueError(f"Expected exactly one Shift row in Calendar activity block; found {len(shift_rows)}")
    shift = shift_rows.iloc[0]
    intervals = [(hm_to_min(shift["Start"]), hm_to_min(shift["End"]))]

    blocked = act[(act["Available_for_Process"].fillna(0).astype(int) == 0)
                  & act["Start"].notna() & act["End"].notna()]
    for _, r in blocked.iterrows():
        try:
            b0, b1 = hm_to_min(r["Start"]), hm_to_min(r["End"])
        except Exception:
            continue
        nxt = []
        for a, b in intervals:
            if b1 <= a or b0 >= b:
                nxt.append((a, b))
            else:
                if a < b0:
                    nxt.append((a, b0))
                if b1 < b:
                    nxt.append((b1, b))
        intervals = nxt
    return sorted(intervals)


def _require_columns(df: pd.DataFrame, sheet: str, cols: List[str]):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"Sheet {sheet} missing columns: {missing}")


def load_problem(input_path: str, due_time: str = DUE_TIME_DEFAULT) -> Dict[str, Any]:
    xls = pd.ExcelFile(input_path)
    required_sheets = ["SO", "Operations", "Eligible_Resources", "Resources", "Human_Capacity", "Calendar", "Lead_Time"]
    missing_sheets = [s for s in required_sheets if s not in xls.sheet_names]
    if missing_sheets:
        raise ValueError(f"Missing required sheets: {missing_sheets}")

    SO = pd.read_excel(xls, "SO")
    OPS = pd.read_excel(xls, "Operations")
    ER = pd.read_excel(xls, "Eligible_Resources")
    RES = pd.read_excel(xls, "Resources")
    HC = pd.read_excel(xls, "Human_Capacity")
    CAL = pd.read_excel(xls, "Calendar")
    LT = pd.read_excel(xls, "Lead_Time")

    _require_columns(SO, "SO", ["SO_ID", "Release_Date", "Due_Date"])
    _require_columns(OPS, "Operations", ["Operation_ID", "SO_ID", "Component_Code", "Op_Code", "Category", "Predecessor"])
    _require_columns(ER, "Eligible_Resources", ["Operation_ID", "Resource", "Alt_No", "Resource_Type", "Total_Time",
                                                "Manpower_Group", "Manpower_per_Resource", "Extra_Group", "Extra_Manpower"])
    _require_columns(RES, "Resources", ["Resource", "Resource_Type", "Capacity_Units"])
    _require_columns(HC, "Human_Capacity", ["Start_Date", "End_Date"] + HUMAN_GROUPS)
    _require_columns(CAL, "Calendar", ["Date", "Is_Working_Day", "Activity", "Start", "End", "Available_for_Process"])
    _require_columns(LT, "Lead_Time", ["LT_Type", "SO_ID", "Operation_ID", "Component_Code", "LT_Plan_Days"])

    SO = SO.copy(); OPS = OPS.copy(); ER = ER.copy(); RES = RES.copy(); HC = HC.copy(); CAL = CAL.copy(); LT = LT.copy()
    SO["SO_ID"] = SO["SO_ID"].astype(str)
    OPS["Operation_ID"] = OPS["Operation_ID"].astype(str)
    OPS["SO_ID"] = OPS["SO_ID"].astype(str)
    OPS["Component_Code"] = OPS["Component_Code"].astype(str)
    ER["Operation_ID"] = ER["Operation_ID"].astype(str)
    ER["Resource"] = ER["Resource"].astype(str)
    RES["Resource"] = RES["Resource"].astype(str)

    # Productive calendar
    intervals = build_productive_intervals(CAL)
    productive_min_day = sum(b - a for a, b in intervals)
    if productive_min_day != 400:
        raise ValueError(f"Productive calendar = {productive_min_day} min/day; expected 400 min/day.")

    work_dates = pd.to_datetime(CAL.loc[CAL["Is_Working_Day"].eq(1), "Date"]).dt.normalize().drop_duplicates().sort_values()
    if len(work_dates) == 0:
        raise ValueError("Calendar contains no working dates.")
    work_minutes = []
    for d in work_dates:
        for a, b in intervals:
            st = d + pd.Timedelta(minutes=a)
            work_minutes.extend(pd.date_range(st, periods=b-a, freq="min").to_pydatetime())
    work_minutes = np.array(work_minutes, dtype="datetime64[m]")
    work_min_int = work_minutes.astype(np.int64)
    wm_pd = pd.to_datetime(work_minutes)

    # Operation indexing
    op_ids = OPS["Operation_ID"].tolist()
    if EXPECTED_OPERATION_COUNT is not None and len(op_ids) != EXPECTED_OPERATION_COUNT:
        raise ValueError(f"Expected {EXPECTED_OPERATION_COUNT} operations; found {len(op_ids)}")
    if len(set(op_ids)) != len(op_ids):
        dup = OPS[OPS["Operation_ID"].duplicated(keep=False)]["Operation_ID"].tolist()
        raise ValueError(f"Duplicate Operation_ID found: {dup[:20]}")
    op_index = {o: i for i, o in enumerate(op_ids)}
    n = len(op_ids)

    # Predecessors
    pred_lists: List[List[int]] = []
    for x in OPS["Predecessor"]:
        if pd.isna(x) or str(x).strip() == "":
            pred_lists.append([])
        else:
            parts = [p.strip() for p in str(x).replace(",", ";").split(";") if p.strip()]
            missing = [p for p in parts if p not in op_index]
            if missing:
                raise ValueError(f"Missing predecessor(s): {missing}")
            pred_lists.append([op_index[p] for p in parts])

    pred_starts = np.zeros(n + 1, dtype=np.int32)
    pred_flat = []
    for i, ps in enumerate(pred_lists):
        pred_starts[i] = len(pred_flat)
        pred_flat.extend(ps)
    pred_starts[n] = len(pred_flat)
    pred_flat = np.array(pred_flat, dtype=np.int32)

    # Successors for fast indegree updates / graph checks
    succ_lists: List[List[int]] = [[] for _ in range(n)]
    for i, ps in enumerate(pred_lists):
        for p in ps:
            succ_lists[p].append(i)
    succ_starts = np.zeros(n + 1, dtype=np.int32)
    succ_flat = []
    for i, ss in enumerate(succ_lists):
        succ_starts[i] = len(succ_flat)
        succ_flat.extend(ss)
    succ_starts[n] = len(succ_flat)
    succ_flat = np.array(succ_flat, dtype=np.int32)

    # DAG check
    indeg_test = np.array([len(x) for x in pred_lists], dtype=int)
    q = [i for i in range(n) if indeg_test[i] == 0]
    seen = 0
    while q:
        i = q.pop()
        seen += 1
        for j in succ_lists[i]:
            indeg_test[j] -= 1
            if indeg_test[j] == 0:
                q.append(j)
    if seen != n:
        raise ValueError(f"Precedence graph contains cycle(s): only {seen}/{n} operations topologically reachable.")

    # SO mapping
    so_ids = SO["SO_ID"].tolist()
    if len(set(so_ids)) != len(so_ids):
        raise ValueError("Duplicate SO_ID found in SO sheet.")
    so_index = {s: i for i, s in enumerate(so_ids)}
    if any(s not in so_index for s in OPS["SO_ID"]):
        raise ValueError("Operations contains SO_ID not listed in SO sheet.")
    op_so_idx = np.array([so_index[s] for s in OPS["SO_ID"]], dtype=np.int16)
    op_comp = OPS["Component_Code"].astype(str).tolist()

    release_ts = pd.to_datetime(SO["Release_Date"]).dt.normalize()
    due_ts = pd.Series([parse_due_timestamp(x, due_time) for x in SO["Due_Date"]])
    if np.any(due_ts.values < release_ts.values):
        raise ValueError("Found Due_Date earlier than Release_Date.")
    release_min = np.array([timestamp_min(x) for x in release_ts], dtype=np.int64)
    due_min = np.array([timestamp_min(x) for x in due_ts], dtype=np.int64)

    # Exactly one inspection and one Final Assembly per SO
    inspection = {}
    final_assembly = {}
    for i, row in OPS.iterrows():
        s = str(row["SO_ID"])
        op_code = str(row["Op_Code"]).strip().upper()
        cat = str(row["Category"]).strip().lower()
        comp = str(row["Component_Code"]).strip().upper()
        if op_code == "INS" or cat == "inspeksi":
            if s in inspection:
                raise ValueError(f"More than one SO-level inspection candidate for SO {s}")
            inspection[s] = i
        if op_code == "FA" or comp == "FA" or cat == "final assembly":
            if s in final_assembly:
                raise ValueError(f"More than one Final Assembly candidate for SO {s}")
            final_assembly[s] = i
    for s in so_ids:
        if s not in inspection:
            raise ValueError(f"No incoming inspection operation found for SO {s}")
        if s not in final_assembly:
            raise ValueError(f"No Final Assembly operation found for SO {s}")
    insp_by_so = np.array([inspection[s] for s in so_ids], dtype=np.int32)
    fa_by_so = np.array([final_assembly[s] for s in so_ids], dtype=np.int32)

    # Ensure every operation in each SO can reach FA (FA is the terminal completion node)
    for si, s in enumerate(so_ids):
        fa = int(fa_by_so[si])
        ancestors = set([fa])
        stack = [fa]
        while stack:
            x = stack.pop()
            for p in pred_lists[x]:
                if p not in ancestors:
                    ancestors.add(p); stack.append(p)
        ops_so = set(np.where(op_so_idx == si)[0].tolist())
        missing_to_fa = ops_so - ancestors
        if missing_to_fa:
            examples = [op_ids[i] for i in list(missing_to_fa)[:10]]
            raise ValueError(f"SO {s}: operations not connected to Final Assembly: {examples}")

    # Lead times
    comp_lt: Dict[Tuple[str, str], int] = {}
    vendor_lt: Dict[str, int] = {}
    for _, r in LT.iterrows():
        typ = str(r["LT_Type"]).strip().lower()
        days = int(r["LT_Plan_Days"])
        if days < 0:
            raise ValueError("Negative lead time found.")
        if typ == "component":
            key = (str(r["SO_ID"]), str(r["Component_Code"]))
            comp_lt[key] = days
        elif typ == "vendor":
            vendor_lt[str(r["Operation_ID"])] = days

    comp_lt_days = np.zeros(n, dtype=np.int16)
    for i in range(n):
        comp_lt_days[i] = int(comp_lt.get((str(OPS.iloc[i]["SO_ID"]), op_comp[i]), 0))

    # Every component LT should map to an operation branch in the same SO.
    for (s, comp), d in comp_lt.items():
        mask = (OPS["SO_ID"].astype(str) == s) & (OPS["Component_Code"].astype(str) == comp)
        if not mask.any():
            raise ValueError(f"Component LT ({s}, {comp}) has no matching operation branch.")

    # Resource master
    internal_res = RES[RES["Resource_Type"].astype(str).str.strip().str.lower().eq("internal")].copy()
    internal_names = internal_res["Resource"].astype(str).tolist()
    res_id = {r: i for i, r in enumerate(internal_names)}
    if len(res_id) != len(internal_names):
        raise ValueError("Duplicate internal Resource name in Resources sheet.")
    res_cap = np.array([int(internal_res.loc[internal_res["Resource"].eq(r), "Capacity_Units"].iloc[0])
                        for r in internal_names], dtype=np.int16)
    if np.any(res_cap <= 0):
        raise ValueError("Internal resource has Capacity_Units <= 0.")

    # Human capacity per productive minute
    gid = {g: i for i, g in enumerate(HUMAN_GROUPS)}
    hcap = np.zeros((len(HUMAN_GROUPS), len(work_minutes)), dtype=np.int16)
    for gi, g in enumerate(HUMAN_GROUPS):
        for _, r in HC.iterrows():
            st = pd.Timestamp(r["Start_Date"]).normalize()
            en = pd.Timestamp(r["End_Date"]).normalize() + pd.Timedelta(days=1)
            mask = (wm_pd >= st) & (wm_pd < en)
            hcap[gi, mask] = int(r[g])
    if np.any(hcap <= 0):
        zero_groups = [HUMAN_GROUPS[g] for g in range(len(HUMAN_GROUPS)) if np.any(hcap[g] <= 0)]
        raise ValueError(f"Human capacity has zero/uncovered productive minutes for groups: {zero_groups}")

    # Eligible-resource arrays
    grouped = {oid: g.sort_values("Alt_No").reset_index(drop=True) for oid, g in ER.groupby("Operation_ID")}
    missing_er = [oid for oid in op_ids if oid not in grouped]
    if missing_er:
        raise ValueError(f"Operations without eligible resource: {missing_er[:20]}")

    alt_count = np.array([len(grouped[oid]) for oid in op_ids], dtype=np.int16)
    max_alt = int(alt_count.max())
    alt_res = np.full((n, max_alt), -1, dtype=np.int16)
    alt_type = np.full((n, max_alt), -1, dtype=np.int8)   # 0 internal, 1 vendor
    alt_dur = np.zeros((n, max_alt), dtype=np.int32)      # productive minutes for internal
    alt_vendor_days = np.zeros((n, max_alt), dtype=np.int16)
    alt_g1 = np.full((n, max_alt), -1, dtype=np.int8)
    alt_q1 = np.zeros((n, max_alt), dtype=np.int16)
    alt_g2 = np.full((n, max_alt), -1, dtype=np.int8)
    alt_q2 = np.zeros((n, max_alt), dtype=np.int16)
    alt_name = [[None] * max_alt for _ in range(n)]
    nominal_min = np.zeros(n, dtype=np.int32)

    for i, oid in enumerate(op_ids):
        g = grouped[oid]
        internal_times = []
        all_nominal = []
        for a, (_, r) in enumerate(g.iterrows()):
            resource = str(r["Resource"])
            typ = str(r["Resource_Type"]).strip().lower()
            alt_name[i][a] = resource
            if typ == "vendor" or resource.strip().lower() == "vendor":
                alt_type[i, a] = 1
                if oid not in vendor_lt:
                    raise ValueError(f"Vendor alternative without Vendor LT row: {oid}")
                alt_vendor_days[i, a] = int(vendor_lt[oid])
                all_nominal.append(int(vendor_lt[oid]) * 1440)
            else:
                alt_type[i, a] = 0
                if resource not in res_id:
                    raise ValueError(f"Eligible internal resource not in Resources master: {resource} for {oid}")
                alt_res[i, a] = res_id[resource]
                if pd.isna(r["Total_Time"]) or float(r["Total_Time"]) <= 0:
                    raise ValueError(f"Invalid Total_Time for {oid} / {resource}")
                mins = max(1, int(math.ceil(float(r["Total_Time"]) * 60 - 1e-9)))
                alt_dur[i, a] = mins
                internal_times.append(mins)
                all_nominal.append(mins)

            mg = r["Manpower_Group"]
            if pd.notna(mg) and str(mg).strip() not in ("", "Vendor"):
                mg = str(mg).strip()
                if mg not in gid:
                    raise ValueError(f"Unknown Manpower_Group '{mg}' for {oid}")
                q = int(r["Manpower_per_Resource"])
                alt_g1[i, a] = gid[mg]
                alt_q1[i, a] = q
                if q > int(hcap[gid[mg]].max()):
                    raise ValueError(f"Manpower requirement exceeds all capacity: {oid}, {mg}, q={q}")

            eg = r["Extra_Group"]
            if pd.notna(eg) and str(eg).strip() != "":
                eg = str(eg).strip()
                if eg not in gid:
                    raise ValueError(f"Unknown Extra_Group '{eg}' for {oid}")
                q = int(r["Extra_Manpower"])
                alt_g2[i, a] = gid[eg]
                alt_q2[i, a] = q
                if q > int(hcap[gid[eg]].max()):
                    raise ValueError(f"Extra manpower requirement exceeds all capacity: {oid}, {eg}, q={q}")

        nominal_min[i] = min(all_nominal)

    # Vendor LT rows must correspond to a vendor alternative
    for oid, days in vendor_lt.items():
        if oid not in grouped:
            raise ValueError(f"Vendor LT operation missing from Eligible_Resources: {oid}")
        gg = grouped[oid]
        if not ((gg["Resource_Type"].astype(str).str.lower() == "vendor") | (gg["Resource"].astype(str).str.lower() == "vendor")).any():
            raise ValueError(f"Vendor LT exists but operation has no Vendor eligible alternative: {oid}")

    return {
        "SO": SO, "OPS": OPS, "ER": ER, "RES": RES, "HC": HC, "CAL": CAL, "LT": LT,
        "op_ids": op_ids, "op_index": op_index, "so_ids": so_ids, "so_index": so_index,
        "op_so_idx": op_so_idx, "op_comp": op_comp,
        "release_min": release_min, "due_min": due_min,
        "release_ts": release_ts, "due_ts": due_ts,
        "pred_lists": pred_lists, "succ_lists": succ_lists,
        "pred_starts": pred_starts, "pred_flat": pred_flat,
        "succ_starts": succ_starts, "succ_flat": succ_flat,
        "insp_by_so": insp_by_so, "fa_by_so": fa_by_so,
        "comp_lt_days": comp_lt_days,
        "work_min_int": work_min_int, "work_minutes": work_minutes,
        "intervals": intervals, "productive_min_day": productive_min_day,
        "alt_count": alt_count, "alt_res": alt_res, "alt_type": alt_type,
        "alt_dur": alt_dur, "alt_vendor_days": alt_vendor_days,
        "alt_g1": alt_g1, "alt_q1": alt_q1, "alt_g2": alt_g2, "alt_q2": alt_q2,
        "alt_name": alt_name, "res_cap": res_cap, "hcap": hcap,
        "nominal_min": nominal_min, "grouped": grouped, "internal_names": internal_names,
        "due_time": due_time,
    }


# ---------------------------
# Input validation report
# ---------------------------

def input_validation_report(P: Dict[str, Any]) -> pd.DataFrame:
    rows = [
        ("Sheets required", "PASS", "SO, Operations, Eligible_Resources, Resources, Human_Capacity, Calendar, Lead_Time"),
        ("Operation count", "PASS", str(len(P["op_ids"]))),
        ("Unique Operation_ID", "PASS", str(len(set(P["op_ids"])))),
        ("Predecessor references", "PASS", "No missing predecessor"),
        ("Precedence DAG", "PASS", "No cycle"),
        ("All operations connected to Final Assembly", "PASS", "Yes"),
        ("Eligible resources", "PASS", "Every operation has >=1 eligible alternative"),
        ("Internal resource master", "PASS", f"{len(P['internal_names'])} internal resources"),
        ("Productive time/day", "PASS", f"{P['productive_min_day']} minutes"),
        ("Human capacity coverage", "PASS", "All productive minutes covered"),
        ("Component lead time rows", "PASS", str(int((P["LT"]["LT_Type"].astype(str).str.lower()=="component").sum()))),
        ("Vendor lead time rows", "PASS", str(int((P["LT"]["LT_Type"].astype(str).str.lower()=="vendor").sum()))),
        ("Due-time convention", "PASS", f"Due date interpreted at {P['due_time']}"),
    ]
    return pd.DataFrame(rows, columns=["Check", "Status", "Detail"])


# ---------------------------
# Shared scheduling primitives
# ---------------------------

def base_time_for_operation(P, i: int, finish: np.ndarray, include_component_gate: bool = True) -> int:
    si = int(P["op_so_idx"][i])
    base = int(P["release_min"][si])
    for p in P["pred_lists"][i]:
        if finish[p] < 0:
            raise RuntimeError(f"Predecessor not scheduled for {P['op_ids'][i]}: {P['op_ids'][p]}")
        base = max(base, int(finish[p]))
    if include_component_gate and int(P["comp_lt_days"][i]) > 0:
        insp = int(P["insp_by_so"][si])
        if finish[insp] < 0:
            raise RuntimeError(f"Inspection not scheduled before component-LT operation {P['op_ids'][i]}")
        base = max(base, int(finish[insp]) + int(P["comp_lt_days"][i]) * 1440)
    return base


def precedence_ready_time(P, i: int, finish: np.ndarray) -> int:
    """Time operation entered precedence-ready state. Does not include component/resource availability."""
    si = int(P["op_so_idx"][i])
    t = int(P["release_min"][si])
    for p in P["pred_lists"][i]:
        t = max(t, int(finish[p]))
    return t


def find_internal_slot(P, i: int, a: int, base: int, ruse: np.ndarray, huse: np.ndarray):
    wm = P["work_min_int"]
    L = len(wm)
    rid = int(P["alt_res"][i, a])
    dur = int(P["alt_dur"][i, a])
    idx = int(np.searchsorted(wm, base))
    while idx + dur <= L:
        bad = -1
        for z in range(idx, idx + dur):
            if ruse[rid, z] + 1 > P["res_cap"][rid]:
                bad = z; break
            g1 = int(P["alt_g1"][i, a])
            if g1 >= 0 and huse[g1, z] + int(P["alt_q1"][i, a]) > P["hcap"][g1, z]:
                bad = z; break
            g2 = int(P["alt_g2"][i, a])
            if g2 >= 0 and huse[g2, z] + int(P["alt_q2"][i, a]) > P["hcap"][g2, z]:
                bad = z; break
        if bad < 0:
            st = int(wm[idx])
            fn = int(wm[idx + dur - 1]) + 1
            return idx, st, fn
        idx = bad + 1
    return None


def eval_alt(P, i: int, a: int, base: int, ruse: np.ndarray, huse: np.ndarray):
    if int(P["alt_type"][i, a]) == 1:
        st = base
        fn = base + int(P["alt_vendor_days"][i, a]) * 1440
        return {"alt": a, "idx": None, "start": st, "finish": fn, "resource": P["alt_name"][i][a], "type": "Vendor"}
    slot = find_internal_slot(P, i, a, base, ruse, huse)
    if slot is None:
        return None
    idx, st, fn = slot
    return {"alt": a, "idx": idx, "start": st, "finish": fn, "resource": P["alt_name"][i][a], "type": "Internal"}


def commit_alt(P, i: int, choice: dict, ruse: np.ndarray, huse: np.ndarray):
    a = int(choice["alt"])
    if int(P["alt_type"][i, a]) == 1:
        return
    idx = int(choice["idx"])
    dur = int(P["alt_dur"][i, a])
    rid = int(P["alt_res"][i, a])
    ruse[rid, idx:idx+dur] += 1
    g1 = int(P["alt_g1"][i, a])
    if g1 >= 0:
        huse[g1, idx:idx+dur] += int(P["alt_q1"][i, a])
    g2 = int(P["alt_g2"][i, a])
    if g2 >= 0:
        huse[g2, idx:idx+dur] += int(P["alt_q2"][i, a])


def compute_metrics(P, finish: np.ndarray) -> Tuple[int, int, int, int]:
    completions = np.array([int(finish[int(i)]) for i in P["fa_by_so"]], dtype=np.int64)
    if np.any(completions < 0):
        return (999, 10**18, 10**18, 10**18)
    tard = np.maximum(0, completions - P["due_min"])
    nlate = int(np.sum(tard > 0))
    total = int(np.sum(tard))
    mx = int(np.max(tard))
    makespan = int(np.max(completions) - np.min(P["release_min"]))
    return nlate, total, mx, makespan


def build_so_summary(P, finish: np.ndarray, method: str) -> pd.DataFrame:
    rows = []
    for si, s in enumerate(P["so_ids"]):
        fi = int(finish[int(P["fa_by_so"][si])])
        tard = max(0, fi - int(P["due_min"][si]))
        rows.append({
            "Method": method,
            "SO_ID": s,
            "Release": min_to_timestamp(int(P["release_min"][si])),
            "Due": min_to_timestamp(int(P["due_min"][si])),
            "Completion": min_to_timestamp(fi),
            "Tardiness_Min": int(tard),
            "Tardiness_Days": float(tard / 1440),
            "Status": "LATE" if tard > 0 else "ON_TIME",
        })
    return pd.DataFrame(rows)


# ---------------------------
# Heuristic schedulers
# ---------------------------

def heuristic_priority_key(P, rule: str, i: int, finish: np.ndarray, unscheduled: np.ndarray,
                           decision_time: int, ready_time: int):
    si = int(P["op_so_idx"][i])
    oid = P["op_ids"][i]
    due = int(P["due_min"][si])
    nominal = int(P["nominal_min"][i])

    if rule == "FCFS":
        return (ready_time, due, oid)
    if rule == "EDD":
        return (due, ready_time, oid)
    if rule == "SPT":
        return (nominal, due, ready_time, oid)
    if rule == "CR":
        mask = (P["op_so_idx"] == si) & unscheduled
        rpt = int(np.sum(P["nominal_min"][mask]))
        rpt = max(rpt, 1)
        cr = (due - decision_time) / rpt
        return (cr, due, ready_time, oid)
    raise ValueError(f"Unknown heuristic rule: {rule}")


def run_heuristic(P, rule: str):
    t0 = time.perf_counter()
    n = len(P["op_ids"])
    R = len(P["res_cap"])
    G = len(HUMAN_GROUPS)
    L = len(P["work_min_int"])
    ruse = np.zeros((R, L), dtype=np.int16)
    huse = np.zeros((G, L), dtype=np.int16)
    finish = np.full(n, -1, dtype=np.int64)
    start = np.full(n, -1, dtype=np.int64)
    selected_alt = np.full(n, -1, dtype=np.int16)
    indeg = np.array([len(x) for x in P["pred_lists"]], dtype=np.int32)
    unscheduled = np.ones(n, dtype=bool)
    rows = []

    for step in range(n):
        ready = np.where(unscheduled & (indeg == 0))[0].tolist()
        if not ready:
            raise RuntimeError(f"{rule}: no ready operation at step {step}; precedence state invalid.")

        ready_times = {i: precedence_ready_time(P, i, finish) for i in ready}
        base_times = {i: base_time_for_operation(P, i, finish, include_component_gate=True) for i in ready}
        decision_time = min(base_times.values())
        selected = min(ready, key=lambda i: heuristic_priority_key(
            P, rule, i, finish, unscheduled, decision_time, ready_times[i]
        ))

        base = base_times[selected]
        candidates = []
        for a in range(int(P["alt_count"][selected])):
            ch = eval_alt(P, selected, a, base, ruse, huse)
            if ch is not None:
                candidates.append(ch)
        if not candidates:
            raise RuntimeError(f"{rule}: no feasible resource slot for {P['op_ids'][selected]} within Calendar horizon.")

        # ECT, then earliest start, then resource label, then alternative index.
        choice = min(candidates, key=lambda x: (x["finish"], x["start"], str(x["resource"]), x["alt"]))
        commit_alt(P, selected, choice, ruse, huse)
        start[selected] = int(choice["start"])
        finish[selected] = int(choice["finish"])
        selected_alt[selected] = int(choice["alt"])
        unscheduled[selected] = False
        for j in P["succ_lists"][selected]:
            indeg[j] -= 1

        a = int(choice["alt"])
        rows.append({
            "Method": f"{rule}+ECT",
            "Schedule_Order": step + 1,
            "Operation_ID": P["op_ids"][selected],
            "SO_ID": P["so_ids"][int(P["op_so_idx"][selected])],
            "Component_Code": P["op_comp"][selected],
            "Alt_No": a + 1,
            "Resource": choice["resource"],
            "Resource_Type": choice["type"],
            "Start": min_to_timestamp(int(choice["start"])),
            "Finish": min_to_timestamp(int(choice["finish"])),
            "Duration_Productive_Min": int(P["alt_dur"][selected, a]) if choice["type"] == "Internal" else np.nan,
            "Vendor_LT_Days": int(P["alt_vendor_days"][selected, a]) if choice["type"] == "Vendor" else np.nan,
            "Manpower_Group": HUMAN_GROUPS[int(P["alt_g1"][selected, a])] if int(P["alt_g1"][selected, a]) >= 0 else None,
            "Manpower_Qty": int(P["alt_q1"][selected, a]),
            "Extra_Group": HUMAN_GROUPS[int(P["alt_g2"][selected, a])] if int(P["alt_g2"][selected, a]) >= 0 else None,
            "Extra_Qty": int(P["alt_q2"][selected, a]),
        })

    runtime = time.perf_counter() - t0
    metrics = compute_metrics(P, finish)
    return {
        "method": f"{rule}+ECT",
        "metrics": metrics,
        "runtime_sec": runtime,
        "finish": finish,
        "start": start,
        "selected_alt": selected_alt,
        "schedule": pd.DataFrame(rows),
        "so_summary": build_so_summary(P, finish, f"{rule}+ECT"),
    }


# ---------------------------
# Fast GA decoder (Numba)
# ---------------------------

@njit
def decode_ga_numba(priority_perm, genes, op_so_idx, release_min, due_min, work_min_int,
                    pred_starts, pred_flat, succ_starts, succ_flat, insp_by_so, fa_by_so, comp_lt_days,
                    alt_count, alt_res, alt_type, alt_dur, alt_vendor_days,
                    alt_g1, alt_q1, alt_g2, alt_q2, res_cap, hcap):
    n = len(op_so_idx)
    L = len(work_min_int)
    R = len(res_cap)
    G = hcap.shape[0]

    rank = np.empty(n, np.int32)
    for pos in range(n):
        rank[priority_perm[pos]] = pos

    indeg = np.empty(n, np.int16)
    for i in range(n):
        indeg[i] = pred_starts[i+1] - pred_starts[i]

    done = np.zeros(n, np.uint8)
    finish = np.full(n, -1, np.int64)
    start = np.full(n, -1, np.int64)
    selected_alt = np.full(n, -1, np.int16)
    ruse = np.zeros((R, L), np.uint8)
    huse = np.zeros((G, L), np.int16)
    count = 0

    while count < n:
        best = -1
        best_rank = 1_000_000_000
        for i in range(n):
            if done[i] == 0 and indeg[i] == 0 and rank[i] < best_rank:
                best_rank = rank[i]
                best = i
        if best < 0:
            return (999, 1e18, 1e18, 1e18), finish, start, selected_alt

        i = best
        soidx = op_so_idx[i]
        base = release_min[soidx]
        for k in range(pred_starts[i], pred_starts[i+1]):
            p = pred_flat[k]
            if finish[p] > base:
                base = finish[p]

        d = comp_lt_days[i]
        if d > 0:
            insp = insp_by_so[soidx]
            if finish[insp] < 0:
                return (996, 1e18, 1e18, 1e18), finish, start, selected_alt
            ca = finish[insp] + np.int64(d) * 1440
            if ca > base:
                base = ca

        a = int(genes[i] % alt_count[i])
        selected_alt[i] = a

        if alt_type[i, a] == 1:
            st = base
            fn = base + np.int64(alt_vendor_days[i, a]) * 1440
        else:
            rid = alt_res[i, a]
            dur = alt_dur[i, a]
            idx = np.searchsorted(work_min_int, base)
            found = False
            while idx + dur <= L:
                bad = -1
                for z in range(idx, idx + dur):
                    if ruse[rid, z] + 1 > res_cap[rid]:
                        bad = z; break
                    g1 = alt_g1[i, a]
                    if g1 >= 0 and huse[g1, z] + alt_q1[i, a] > hcap[g1, z]:
                        bad = z; break
                    g2 = alt_g2[i, a]
                    if g2 >= 0 and huse[g2, z] + alt_q2[i, a] > hcap[g2, z]:
                        bad = z; break
                if bad < 0:
                    found = True; break
                idx = bad + 1
            if not found:
                return (998, 1e18, 1e18, 1e18), finish, start, selected_alt

            st = work_min_int[idx]
            fn = work_min_int[idx + dur - 1] + 1
            for z in range(idx, idx + dur):
                ruse[rid, z] += 1
                g1 = alt_g1[i, a]
                if g1 >= 0:
                    huse[g1, z] += alt_q1[i, a]
                g2 = alt_g2[i, a]
                if g2 >= 0:
                    huse[g2, z] += alt_q2[i, a]

        start[i] = st
        finish[i] = fn
        done[i] = 1
        count += 1

        for k in range(succ_starts[i], succ_starts[i+1]):
            j = succ_flat[k]
            indeg[j] -= 1

    nlate = 0
    total = 0.0
    mx = 0.0
    cmax = -1
    relmin = release_min[0]
    for s in range(len(due_min)):
        fi = finish[fa_by_so[s]]
        if fi > cmax:
            cmax = fi
        if release_min[s] < relmin:
            relmin = release_min[s]
        t = fi - due_min[s]
        if t > 0:
            nlate += 1
            total += t
            if t > mx:
                mx = t
    return (nlate, total, mx, float(cmax - relmin)), finish, start, selected_alt


def decode_ga(P, ps, rs):
    return decode_ga_numba(
        ps, rs,
        P["op_so_idx"], P["release_min"], P["due_min"], P["work_min_int"],
        P["pred_starts"], P["pred_flat"], P["succ_starts"], P["succ_flat"],
        P["insp_by_so"], P["fa_by_so"], P["comp_lt_days"],
        P["alt_count"], P["alt_res"], P["alt_type"], P["alt_dur"], P["alt_vendor_days"],
        P["alt_g1"], P["alt_q1"], P["alt_g2"], P["alt_q2"], P["res_cap"], P["hcap"]
    )


def ox(p1, p2, rng):
    n = len(p1)
    a, b = sorted(rng.choice(n, 2, replace=False))
    if a == b:
        b = min(n, a + 1)
    child = np.full(n, -1, dtype=np.int32)
    child[a:b] = p1[a:b]
    used = set(child[a:b].tolist())
    fill = [int(x) for x in p2 if int(x) not in used]
    pos = list(range(0, a)) + list(range(b, n))
    for q, x in zip(pos, fill):
        child[q] = x
    return child


def tournament(fits, k, rng):
    ids = rng.integers(0, len(fits), size=k)
    best = int(ids[0])
    for z in ids[1:]:
        z = int(z)
        if tuple(fits[z]) < tuple(fits[best]):
            best = z
    return best


def checkpoint_path(workdir: Path, run_key: str) -> Path:
    return workdir / "checkpoints" / f"{run_key}.pkl.gz"


def run_ga_checkpointed(P, workdir: Path, run_key: str, seed: int, params: dict):
    ck = checkpoint_path(workdir, run_key)
    n = len(P["op_ids"])
    alt_count = P["alt_count"]

    if ck.exists():
        state = load_pickle_gz(ck)
        if state.get("params") != params or int(state.get("seed")) != int(seed):
            ck.unlink()
            return run_ga_checkpointed(P, workdir, run_key, seed, params)
        generation_done = int(state["generation_done"])
        pps = state["pps"]
        prs = state["prs"]
        best_fit = tuple(state["best_fit"]) if state["best_fit"] is not None else None
        best_ps = state["best_ps"]
        best_rs = state["best_rs"]
        history = state["history"]
        rng = np.random.default_rng()
        rng.bit_generator.state = state["rng_state"]
        elapsed_before = float(state.get("elapsed_sec", 0.0))
        print(f"[RESUME] {run_key} from generation {generation_done + 1}")
    else:
        generation_done = 0
        rng = np.random.default_rng(seed)
        pop = int(params["population_size"])
        pps = np.stack([rng.permutation(n).astype(np.int32) for _ in range(pop)])
        prs = np.stack([
            np.array([rng.integers(int(c)) for c in alt_count], dtype=np.int16)
            for _ in range(pop)
        ])
        best_fit = None; best_ps = None; best_rs = None; history = []
        elapsed_before = 0.0

    t0 = time.perf_counter()
    start_gen = generation_done
    for gen in range(start_gen, int(params["generations"])):
        fits = []
        for i in range(int(params["population_size"])):
            fit, _, _, _ = decode_ga(P, pps[i], prs[i])
            fits.append(tuple(fit))

        order = sorted(range(len(fits)), key=lambda idx: fits[idx])
        gen_best = fits[order[0]]
        if best_fit is None or gen_best < best_fit:
            best_fit = tuple(gen_best)
            best_ps = pps[order[0]].copy()
            best_rs = prs[order[0]].copy()
        history.append({
            "Generation": gen + 1,
            "N_Late": int(best_fit[0]),
            "Total_Tardiness_Min": float(best_fit[1]),
            "Max_Tardiness_Min": float(best_fit[2]),
            "Makespan_Min": float(best_fit[3]),
        })

        # Elitism
        newps = [pps[i].copy() for i in order[:int(params["elite"])]]
        newrs = [prs[i].copy() for i in order[:int(params["elite"])]]

        while len(newps) < int(params["population_size"]):
            i1 = tournament(fits, int(params["tournament"]), rng)
            i2 = tournament(fits, int(params["tournament"]), rng)
            p1s, p2s = pps[i1], pps[i2]
            p1r, p2r = prs[i1], prs[i2]

            if rng.random() < float(params["pc"]):
                c1s = ox(p1s, p2s, rng)
                c2s = ox(p2s, p1s, rng)
                mask = rng.random(n) < 0.5
                c1r = np.where(mask, p1r, p2r).astype(np.int16)
                c2r = np.where(mask, p2r, p1r).astype(np.int16)
            else:
                c1s, c2s = p1s.copy(), p2s.copy()
                c1r, c2r = p1r.copy(), p2r.copy()

            for cs in (c1s, c2s):
                if rng.random() < float(params["pm_seq"]):
                    x, y = rng.choice(n, 2, replace=False)
                    cs[x], cs[y] = cs[y], cs[x]

            for cr in (c1r, c2r):
                if rng.random() < float(params["pm_res"]):
                    idx_candidates = np.where(alt_count > 1)[0]
                    if len(idx_candidates) > 0:
                        o = int(rng.choice(idx_candidates))
                        old = int(cr[o])
                        vals = [v for v in range(int(alt_count[o])) if v != old]
                        cr[o] = int(rng.choice(vals))

            newps.append(c1s); newrs.append(c1r)
            if len(newps) < int(params["population_size"]):
                newps.append(c2s); newrs.append(c2r)

        pps = np.stack(newps[:int(params["population_size"])]).astype(np.int32)
        prs = np.stack(newrs[:int(params["population_size"])]).astype(np.int16)

        elapsed_now = elapsed_before + (time.perf_counter() - t0)
        atomic_pickle_gz({
            "generation_done": gen + 1,
            "pps": pps, "prs": prs,
            "best_fit": best_fit, "best_ps": best_ps, "best_rs": best_rs,
            "history": history,
            "rng_state": rng.bit_generator.state,
            "elapsed_sec": elapsed_now,
            "seed": seed, "params": params,
        }, ck)

        if (gen + 1) % 10 == 0 or gen == 0 or gen + 1 == int(params["generations"]):
            print(f"[{run_key}] gen {gen+1}/{params['generations']} best={best_fit}")

    runtime = elapsed_before + (time.perf_counter() - t0)
    fit, finish, start, selected_alt = decode_ga(P, best_ps, best_rs)

    # first generation at which final-best tuple was reached
    final_tuple = tuple(fit)
    first_final = None
    for h in history:
        t = (h["N_Late"], h["Total_Tardiness_Min"], h["Max_Tardiness_Min"], h["Makespan_Min"])
        if tuple(float(x) for x in t) == tuple(float(x) for x in final_tuple):
            first_final = int(h["Generation"])
            break

    return {
        "fit": tuple(fit), "finish": finish, "start": start, "selected_alt": selected_alt,
        "best_ps": best_ps, "best_rs": best_rs, "history": history,
        "runtime_sec": runtime, "first_final_best_generation": first_final,
    }


def ga_schedule_dataframe(P, result: dict, method="GA") -> pd.DataFrame:
    rows = []
    order = np.argsort(result["start"])
    for seq, i in enumerate(order, start=1):
        a = int(result["selected_alt"][i])
        typ = "Vendor" if int(P["alt_type"][i, a]) == 1 else "Internal"
        rows.append({
            "Method": method,
            "Schedule_Order_By_Start": seq,
            "Operation_ID": P["op_ids"][i],
            "SO_ID": P["so_ids"][int(P["op_so_idx"][i])],
            "Component_Code": P["op_comp"][i],
            "Alt_No": a + 1,
            "Resource": P["alt_name"][i][a],
            "Resource_Type": typ,
            "Start": min_to_timestamp(int(result["start"][i])),
            "Finish": min_to_timestamp(int(result["finish"][i])),
            "Duration_Productive_Min": int(P["alt_dur"][i, a]) if typ == "Internal" else np.nan,
            "Vendor_LT_Days": int(P["alt_vendor_days"][i, a]) if typ == "Vendor" else np.nan,
            "Manpower_Group": HUMAN_GROUPS[int(P["alt_g1"][i, a])] if int(P["alt_g1"][i, a]) >= 0 else None,
            "Manpower_Qty": int(P["alt_q1"][i, a]),
            "Extra_Group": HUMAN_GROUPS[int(P["alt_g2"][i, a])] if int(P["alt_g2"][i, a]) >= 0 else None,
            "Extra_Qty": int(P["alt_q2"][i, a]),
        })
    return pd.DataFrame(rows)


# ---------------------------
# Independent feasibility validator
# ---------------------------

def validate_schedule(P, start: np.ndarray, finish: np.ndarray, selected_alt: np.ndarray, method: str):
    n = len(P["op_ids"])
    issues = []
    wm = P["work_min_int"]
    R = len(P["res_cap"]); G = len(HUMAN_GROUPS); L = len(wm)
    ruse = np.zeros((R, L), dtype=np.int16)
    huse = np.zeros((G, L), dtype=np.int16)

    if np.any(start < 0) or np.any(finish < 0) or np.any(selected_alt < 0):
        issues.append("Not all operations are scheduled/assigned")

    for i in range(n):
        if start[i] < 0 or finish[i] < 0 or selected_alt[i] < 0:
            continue
        si = int(P["op_so_idx"][i])
        oid = P["op_ids"][i]
        a = int(selected_alt[i])

        # release
        if int(start[i]) < int(P["release_min"][si]):
            issues.append(f"Release violation: {oid}")

        # precedence
        for p in P["pred_lists"][i]:
            if int(start[i]) < int(finish[p]):
                issues.append(f"Precedence violation: {P['op_ids'][p]} -> {oid}")

        # component availability
        d = int(P["comp_lt_days"][i])
        if d > 0:
            insp = int(P["insp_by_so"][si])
            ca = int(finish[insp]) + d * 1440
            if int(start[i]) < ca:
                issues.append(f"Component LT violation: {oid}")

        # alternative range
        if a < 0 or a >= int(P["alt_count"][i]):
            issues.append(f"Invalid selected alternative: {oid} alt={a}")
            continue

        if int(P["alt_type"][i, a]) == 1:
            expected = int(start[i]) + int(P["alt_vendor_days"][i, a]) * 1440
            if int(finish[i]) != expected:
                issues.append(f"Vendor LT finish mismatch: {oid}")
            continue

        # internal: start must correspond to productive minute and occupied duration
        idx = int(np.searchsorted(wm, int(start[i])))
        if idx >= L or int(wm[idx]) != int(start[i]):
            issues.append(f"Internal start not on productive minute: {oid}")
            continue
        dur = int(P["alt_dur"][i, a])
        if idx + dur > L:
            issues.append(f"Operation exceeds calendar horizon: {oid}")
            continue
        expected_finish = int(wm[idx + dur - 1]) + 1
        if int(finish[i]) != expected_finish:
            issues.append(f"Internal finish/duration mismatch: {oid}")

        rid = int(P["alt_res"][i, a])
        ruse[rid, idx:idx+dur] += 1
        g1 = int(P["alt_g1"][i, a])
        if g1 >= 0:
            huse[g1, idx:idx+dur] += int(P["alt_q1"][i, a])
        g2 = int(P["alt_g2"][i, a])
        if g2 >= 0:
            huse[g2, idx:idx+dur] += int(P["alt_q2"][i, a])

    for rid in range(R):
        if np.any(ruse[rid] > int(P["res_cap"][rid])):
            mx = int(np.max(ruse[rid] - int(P["res_cap"][rid])))
            issues.append(f"Physical capacity violation: {P['internal_names'][rid]} excess={mx}")

    for g in range(G):
        if np.any(huse[g] > P["hcap"][g]):
            excess = int(np.max(huse[g] - P["hcap"][g]))
            issues.append(f"Human capacity violation: {HUMAN_GROUPS[g]} excess={excess}")

    # All FA are terminal completion points and all SO operations finish no later than FA.
    for si, s in enumerate(P["so_ids"]):
        fa = int(P["fa_by_so"][si])
        fa_finish = int(finish[fa])
        ops = np.where(P["op_so_idx"] == si)[0]
        if np.any(finish[ops] > fa_finish):
            bad = [P["op_ids"][i] for i in ops if finish[i] > fa_finish][:10]
            issues.append(f"SO {s}: operation finishes after Final Assembly: {bad}")

    checks = [
        ("All operations scheduled", not any("Not all operations" in x for x in issues)),
        ("Release date", not any("Release violation" in x for x in issues)),
        ("Precedence", not any("Precedence violation" in x for x in issues)),
        ("Component availability", not any("Component LT violation" in x for x in issues)),
        ("Vendor lead time", not any("Vendor LT" in x for x in issues)),
        ("Productive calendar / duration", not any("productive minute" in x or "duration mismatch" in x or "calendar horizon" in x for x in issues)),
        ("Physical capacity", not any("Physical capacity violation" in x for x in issues)),
        ("Human capacity", not any("Human capacity violation" in x for x in issues)),
        ("Final Assembly terminal", not any("finishes after Final Assembly" in x for x in issues)),
    ]
    summary = pd.DataFrame([{"Method": method, "Constraint": c, "Status": "VALID" if ok else "INVALID"} for c, ok in checks])
    return len(issues) == 0, summary, issues


# ---------------------------
# GA tuning
# ---------------------------

def aggregate_tuning_results(df: pd.DataFrame) -> dict:
    # Quality first: medians in the same lexicographic objective order.
    return {
        "Median_N_Late": float(df["N_Late"].median()),
        "Median_Total_Tardiness_Min": float(df["Total_Tardiness_Min"].median()),
        "Median_Max_Tardiness_Min": float(df["Max_Tardiness_Min"].median()),
        "Median_Makespan_Min": float(df["Makespan_Min"].median()),
        "Std_Makespan_Min": float(df["Makespan_Min"].std(ddof=0)),
        "Mean_Runtime_Sec": float(df["Runtime_Sec"].mean()),
    }


def tuning_selection_key(summary: dict):
    return (
        summary["Median_N_Late"],
        summary["Median_Total_Tardiness_Min"],
        summary["Median_Max_Tardiness_Min"],
        summary["Median_Makespan_Min"],
        summary["Std_Makespan_Min"],
        summary["Mean_Runtime_Sec"],
    )


def run_candidate(P, workdir: Path, factor: str, label: str, params: dict, seeds=TUNING_SEEDS):
    rows = []
    safe_label = str(label).replace(" ", "_").replace("/", "-")
    for seed in seeds:
        key = f"tune_{factor}_{safe_label}_seed{seed}"
        result = run_ga_checkpointed(P, workdir, key, seed, params)
        fit = result["fit"]
        rows.append({
            "Factor": factor, "Candidate": label, "Seed": seed,
            "N_Late": int(fit[0]),
            "Total_Tardiness_Min": int(fit[1]),
            "Max_Tardiness_Min": int(fit[2]),
            "Makespan_Min": int(fit[3]),
            "Runtime_Sec": float(result["runtime_sec"]),
            "First_Final_Best_Generation": result["first_final_best_generation"],
            **{f"Param_{k}": v for k, v in params.items()},
        })
    return pd.DataFrame(rows)


def tune_ga(P, workdir: Path):
    selected = dict(BASELINE_PARAMS)
    all_run_rows = []
    summary_rows = []

    stages = [
        ("Population", [("30", {"population_size": 30}), ("50", {"population_size": 50}), ("70", {"population_size": 70})]),
        ("Generations", [("50", {"generations": 50}), ("70", {"generations": 70}), ("100", {"generations": 100})]),
        ("Rates", [(k, v) for k, v in RATE_SETS.items()]),
        ("Elite", [("1", {"elite": 1}), ("2", {"elite": 2}), ("4", {"elite": 4})]),
        ("Tournament", [("2", {"tournament": 2}), ("3", {"tournament": 3}), ("4", {"tournament": 4})]),
    ]

    for factor, candidates in stages:
        print(f"\n=== TUNING {factor} ===")
        candidate_summaries = []
        for label, patch in candidates:
            params = dict(selected)
            params.update(patch)
            df = run_candidate(P, workdir, factor, label, params)
            all_run_rows.append(df)
            agg = aggregate_tuning_results(df)
            candidate_summaries.append((label, patch, params, agg))
            summary_rows.append({"Factor": factor, "Candidate": label, **params, **agg})
            print(label, agg)

        winner = min(candidate_summaries, key=lambda x: tuning_selection_key(x[3]))
        label, patch, params, agg = winner
        selected.update(patch)
        print(f"SELECTED {factor}: {label} -> {selected}")

    runs_df = pd.concat(all_run_rows, ignore_index=True)
    summary_df = pd.DataFrame(summary_rows)
    summary_df["Selected_At_Stage"] = False
    # Mark the winner within each factor using the same selection key.
    for factor, g in summary_df.groupby("Factor"):
        idx = min(g.index, key=lambda ix: tuning_selection_key(g.loc[ix].to_dict()))
        summary_df.loc[idx, "Selected_At_Stage"] = True

    atomic_csv(runs_df, workdir / "GA_tuning_runs.csv")
    atomic_csv(summary_df, workdir / "GA_tuning_summary.csv")
    write_json(selected, workdir / "GA_selected_params.json")
    return selected, summary_df, runs_df


# ---------------------------
# Plotting helpers
# ---------------------------

def save_convergence_plot(history_df: pd.DataFrame, out_path: Path, title: str):
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(history_df["Generation"], history_df["Makespan_Min"])
    ax.set_xlabel("Generation")
    ax.set_ylabel("Best Makespan (minutes)")
    ax.set_title(title)
    ax.grid(True, alpha=0.25)
    fig.tight_layout()
    fig.savefig(out_path, dpi=180)
    plt.close(fig)


# ---------------------------
# Pipeline stages
# ---------------------------

def run_heuristics_stage(P, workdir: Path):
    summaries = []
    schedules = []
    sos = []
    val_summaries = []
    issue_rows = []
    results = {}

    for rule in ["FCFS", "EDD", "SPT", "CR"]:
        print(f"\n=== HEURISTIC {rule}+ECT ===")
        res = run_heuristic(P, rule)
        method = res["method"]
        valid, vdf, issues = validate_schedule(P, res["start"], res["finish"], res["selected_alt"], method)
        if not valid:
            for x in issues:
                issue_rows.append({"Method": method, "Issue": x})
            raise RuntimeError(f"{method} failed feasibility validation. See issues: {issues[:10]}")
        fit = res["metrics"]
        summaries.append({
            "Method": method,
            "N_Late": int(fit[0]),
            "Total_Tardiness_Min": int(fit[1]),
            "Max_Tardiness_Min": int(fit[2]),
            "Makespan_Min": int(fit[3]),
            "Runtime_Sec": float(res["runtime_sec"]),
            "Total_Tardiness_Days": float(fit[1] / 1440),
            "Max_Tardiness_Days": float(fit[2] / 1440),
            "Makespan_Days": float(fit[3] / 1440),
        })
        schedules.append(res["schedule"])
        sos.append(res["so_summary"])
        val_summaries.append(vdf)
        results[method] = res
        print(f"{method}: {fit} runtime={res['runtime_sec']:.3f}s VALID")

    summary_df = pd.DataFrame(summaries)
    schedule_df = pd.concat(schedules, ignore_index=True)
    so_df = pd.concat(sos, ignore_index=True)
    val_df = pd.concat(val_summaries, ignore_index=True)
    issue_df = pd.DataFrame(issue_rows, columns=["Method", "Issue"])

    atomic_csv(summary_df, workdir / "heuristic_summary.csv")
    atomic_csv(schedule_df, workdir / "heuristic_schedules.csv")
    atomic_csv(so_df, workdir / "heuristic_so_summary.csv")
    atomic_csv(val_df, workdir / "feasibility_heuristics.csv")
    atomic_csv(issue_df, workdir / "feasibility_heuristic_issues.csv")
    return results, summary_df, val_df


def run_final_ga_stage(P, workdir: Path, params: dict):
    rows = []
    all_history = []
    best = None

    # Warm-up compilation before timing thesis runs.
    rng = np.random.default_rng(999)
    ps0 = rng.permutation(len(P["op_ids"])).astype(np.int32)
    rs0 = np.array([rng.integers(int(c)) for c in P["alt_count"]], dtype=np.int16)
    _ = decode_ga(P, ps0, rs0)

    for seed in FINAL_SEEDS:
        print(f"\n=== FINAL GA seed {seed} ===")
        key = f"final_seed{seed}"
        res = run_ga_checkpointed(P, workdir, key, seed, params)
        fit = res["fit"]
        rows.append({
            "Seed": seed,
            "N_Late": int(fit[0]),
            "Total_Tardiness_Min": int(fit[1]),
            "Max_Tardiness_Min": int(fit[2]),
            "Makespan_Min": int(fit[3]),
            "Runtime_Sec": float(res["runtime_sec"]),
            "First_Final_Best_Generation": res["first_final_best_generation"],
            "Total_Tardiness_Days": float(fit[1] / 1440),
            "Max_Tardiness_Days": float(fit[2] / 1440),
            "Makespan_Days": float(fit[3] / 1440),
        })
        h = pd.DataFrame(res["history"])
        h.insert(0, "Seed", seed)
        all_history.append(h)
        if best is None or tuple(res["fit"]) < tuple(best["fit"]):
            best = res
            best["seed"] = seed
        print(f"seed {seed}: {fit} runtime={res['runtime_sec']:.3f}s")

    runs_df = pd.DataFrame(rows)
    hist_df = pd.concat(all_history, ignore_index=True)
    atomic_csv(runs_df, workdir / "GA_final_20_runs.csv")
    atomic_csv(hist_df, workdir / "GA_convergence_history.csv")

    # Validate best schedule independently.
    valid, vdf, issues = validate_schedule(P, best["start"], best["finish"], best["selected_alt"], "GA")
    atomic_csv(vdf, workdir / "feasibility_GA.csv")
    atomic_csv(pd.DataFrame([{"Method": "GA", "Issue": x} for x in issues]), workdir / "feasibility_GA_issues.csv")
    if not valid:
        raise RuntimeError(f"Best GA schedule failed validator: {issues[:20]}")

    sched_df = ga_schedule_dataframe(P, best, method="GA")
    so_df = build_so_summary(P, best["finish"], "GA")
    atomic_csv(sched_df, workdir / "GA_best_schedule.csv")
    atomic_csv(so_df, workdir / "GA_best_so_summary.csv")

    best_hist = hist_df[hist_df["Seed"].eq(best["seed"])].copy()
    save_convergence_plot(best_hist, workdir / "GA_best_convergence.png", f"GA Convergence - Best Run Seed {best['seed']}")
    return best, runs_df, hist_df, vdf


def compare_methods(P, workdir: Path, heuristic_summary: pd.DataFrame, ga_best: dict):
    ga_fit = ga_best["fit"]
    rows = heuristic_summary[["Method", "N_Late", "Total_Tardiness_Min", "Max_Tardiness_Min", "Makespan_Min", "Runtime_Sec"]].to_dict("records")
    rows.append({
        "Method": "GA (best of 20)",
        "N_Late": int(ga_fit[0]),
        "Total_Tardiness_Min": int(ga_fit[1]),
        "Max_Tardiness_Min": int(ga_fit[2]),
        "Makespan_Min": int(ga_fit[3]),
        "Runtime_Sec": float(ga_best["runtime_sec"]),
    })
    df = pd.DataFrame(rows)
    df["Total_Tardiness_Days"] = df["Total_Tardiness_Min"] / 1440
    df["Max_Tardiness_Days"] = df["Max_Tardiness_Min"] / 1440
    df["Makespan_Days"] = df["Makespan_Min"] / 1440
    df["Lexicographic_Key"] = df.apply(lambda r: f"({int(r.N_Late)}, {int(r.Total_Tardiness_Min)}, {int(r.Max_Tardiness_Min)}, {int(r.Makespan_Min)})", axis=1)
    best_key = min((int(r.N_Late), int(r.Total_Tardiness_Min), int(r.Max_Tardiness_Min), int(r.Makespan_Min)) for _, r in df.iterrows())
    df["Best_Primary_Criteria"] = df.apply(lambda r: (int(r.N_Late), int(r.Total_Tardiness_Min), int(r.Max_Tardiness_Min), int(r.Makespan_Min)) == best_key, axis=1)
    atomic_csv(df, workdir / "Method_comparison.csv")
    return df


def run_all(input_path: str, workdir: str, due_time: str, stage: str):
    workdir = Path(workdir)
    workdir.mkdir(parents=True, exist_ok=True)

    print("Loading and validating input...")
    P = load_problem(input_path, due_time=due_time)
    val_input = input_validation_report(P)
    atomic_csv(val_input, workdir / "input_validation.csv")
    print(val_input.to_string(index=False))

    manifest = {
        "Input_File": str(input_path),
        "Operation_Count": len(P["op_ids"]),
        "SO_Count": len(P["so_ids"]),
        "Due_Time_Convention": due_time,
        "Productive_Minutes_Per_Day": P["productive_min_day"],
        "Methods": ["FCFS+ECT", "EDD+ECT", "SPT+ECT", "CR+ECT", "GA"],
        "Lexicographic_Objectives": ["N_Late", "Total_Tardiness", "Max_Tardiness", "Makespan"],
        "Component_LT_Rule": "SO incoming inspection finish + LT_Plan_Days (calendar days); gates component operations",
        "Vendor_Rule": "Vendor uses LT_Plan_Days (calendar days), ignores Total_Time, consumes no internal capacity",
        "Resume_Rule": "Internal operation may pause over nonproductive calendar gaps and resumes on same resource",
    }
    write_json(manifest, workdir / "run_manifest.json")

    if stage == "validate":
        print("Validation-only stage complete.")
        return

    heuristic_results = None; heuristic_summary = None
    if stage in ("heuristics", "all"):
        heuristic_results, heuristic_summary, _ = run_heuristics_stage(P, workdir)
        if stage == "heuristics":
            print("Heuristic stage complete.")
            return

    selected_params = None
    if stage in ("tuning", "all"):
        selected_params, _, _ = tune_ga(P, workdir)
        if stage == "tuning":
            print("Tuning stage complete. Selected params:", selected_params)
            return

    if stage == "final_ga":
        param_file = workdir / "GA_selected_params.json"
        if param_file.exists():
            with open(param_file, "r", encoding="utf-8") as f:
                selected_params = json.load(f)
        else:
            selected_params = dict(BASELINE_PARAMS)
            print("WARNING: GA_selected_params.json not found; using BASELINE_PARAMS:", selected_params)
        best, runs, hist, _ = run_final_ga_stage(P, workdir, selected_params)
        print("Final GA complete. Best:", best["fit"], "seed", best["seed"])
        return

    if stage == "all":
        if heuristic_summary is None:
            _, heuristic_summary, _ = run_heuristics_stage(P, workdir)
        if selected_params is None:
            selected_params = dict(BASELINE_PARAMS)
        ga_best, runs, hist, _ = run_final_ga_stage(P, workdir, selected_params)
        comparison = compare_methods(P, workdir, heuristic_summary, ga_best)

        # Combined feasibility table
        fhs = pd.read_csv(workdir / "feasibility_heuristics.csv")
        fga = pd.read_csv(workdir / "feasibility_GA.csv")
        atomic_csv(pd.concat([fhs, fga], ignore_index=True), workdir / "feasibility_summary.csv")

        print("\n=== FINAL METHOD COMPARISON ===")
        print(comparison.to_string(index=False))
        print("\nALL STAGES COMPLETE.")
        print(f"Outputs saved to: {workdir}")


def main():
    ap = argparse.ArgumentParser(description="PT XYZ final DRCFJSP scheduling pipeline")
    ap.add_argument("--input", required=True, help="Path to Input Penjadwalan_Python TA.xlsx")
    ap.add_argument("--workdir", required=True, help="Persistent output/checkpoint directory, preferably Google Drive")
    ap.add_argument("--due-time", default=DUE_TIME_DEFAULT, help="Interpret date-only due dates at this local clock time; default 17:00")
    ap.add_argument("--stage", choices=["validate", "heuristics", "tuning", "final_ga", "all"], default="all")
    args = ap.parse_args()
    run_all(args.input, args.workdir, args.due_time, args.stage)


if __name__ == "__main__":
    main()
