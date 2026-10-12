# -*- coding: utf-8 -*-
"""Inject RBA data into the dashboard HTML template -> PassThroughGap_Dashboard.html

October 2026 refresh of the July dashboard (CLAUDE/RBA_Passthrough_Stata/scripts/build_dashboard.py):
same method, plus RBA F6 actual housing lending rates and the 30 Sep 2026 hike.
Run from anywhere: python work/build_dashboard.py
"""
import json
import sys
import re
from pathlib import Path

import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
PROJECT = Path(__file__).resolve().parents[1]
RAW = PROJECT / "data" / "raw"
TEMPLATE = PROJECT / "work" / "dashboard_template.html"
NEXT_DECISION = {"label": "Nov 3 2026", "at": "2026-11-03T14:30:00+11:00"}  # Monetary Policy Board meets 2-3 Nov


def load_rba(path, series):
    raw = pd.read_excel(path, sheet_name="Data", header=None)
    sid_row = raw[raw.iloc[:, 0].astype(str).str.strip() == "Series ID"].index[0]
    cols = raw.iloc[sid_row].tolist()
    df = raw.iloc[sid_row + 1:].copy()
    df.columns = ["Date"] + cols[1:]
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"]).set_index("Date")
    return df[[s for s in series if s in df.columns]].apply(pd.to_numeric, errors="coerce")


f1d = load_rba(RAW / "f01d.xlsx", ["FIRMMCRTD"])
f1m = load_rba(RAW / "f01hist.xlsx", ["FIRMMCRT"])
f4 = load_rba(RAW / "f04hist.xlsx", ["FRDIRSAO10K", "FRDIRSAB10K", "FRDIRBTD10K1Y"])
f5 = load_rba(RAW / "f05hist.xlsx", ["FILRHLBVS", "FILRHLBVD"])
f6 = load_rba(RAW / "f06hist.xlsx", ["FLRHOOVA", "FLRHOFVA"])

m = pd.concat([f1m, f4, f5, f6], axis=1, sort=True)
m.index = m.index.to_period("M")
cr_d = f1d["FIRMMCRTD"].resample("ME").last()
cr_d.index = cr_d.index.to_period("M")
m["CASH"] = cr_d.combine_first((m["FIRMMCRT"] / 0.25).round() * 0.25)
m = m.rename(columns={"FILRHLBVS": "SVR", "FILRHLBVD": "DISC", "FRDIRSAO10K": "SAVER",
                      "FRDIRSAB10K": "BONUS", "FRDIRBTD10K1Y": "TD1Y",
                      "FLRHOOVA": "EXIST", "FLRHOFVA": "NEWL"})
last = m["SVR"].last_valid_index()
last_f6 = m["EXIST"].last_valid_index()

# Actual target changes, from the daily table.
daily = f1d["FIRMMCRTD"].dropna().round(4)
chg = daily[daily.diff().fillna(0) != 0]
decisions = [(d, float(daily.shift(1).loc[d]), float(v)) for d, v in chg.items()]
latest_dec = decisions[-1]
# A decision in the last data month (e.g. 30 Sep) is in the month-end cash rate but not yet in retail
# rates, so measurement stops the month before it.
cap = last
if pd.Period(latest_dec[0], "M") >= last:
    cap = pd.Period(latest_dec[0], "M") - 1


def pt(t0m, t, col):
    t0 = pd.Period(t0m, "M") - 1
    t = min(pd.Period(t, "M"), cap)
    if pd.isna(m.loc[t, col]) or pd.isna(m.loc[t0, col]):
        return None
    dc = m.loc[t, "CASH"] - m.loc[t0, "CASH"]
    dr = m.loc[t, col] - m.loc[t0, col]
    v = dr / dc * 100 if dc else np.nan
    return None if pd.isna(v) else round(float(v), 1)


COLS = [("svr", "SVR"), ("exist", "EXIST"), ("saver", "SAVER"), ("bonus", "BONUS")]
episodes = [
    ("2006–08", "2006-05", "2008-03"),
    ("2009–10", "2009-10", "2010-11"),
    ("2022–23", "2022-05", "2023-11"),
    ("2026 (live)", "2026-02", "2026-09"),
]
ep_out = []
for name, s, e in episodes:
    s_p, e_p = pd.Period(s, "M"), pd.Period(e, "M")
    base = s_p - 1
    to = min(e_p, cap)
    delta = m.loc[to, "CASH"] - m.loc[base, "CASH"]
    ep_out.append({
        "name": name,
        "delta": f"+{delta:.2f}pp",
        "bands": [s, str(e_p)],
        "fast": {c: pt(s, s_p + 3, col) for c, col in COLS},
        "full": {c: pt(s, e_p + 3, col) for c, col in COLS},
        "fullTo": str(min(e_p + 3, cap)),
    })
cuts = {"name": "2025 cuts", "delta": "−0.75pp",
        "fast": {c: pt("2025-02", pd.Period("2025-05"), col) for c, col in COLS},
        "full": {c: pt("2025-02", pd.Period("2025-11"), col) for c, col in COLS}}


def series(df_slice, cols):
    out = []
    for p, row in df_slice.iterrows():
        rec = {"d": str(p)}
        for k, c in cols.items():
            v = row[c]
            rec[k] = None if pd.isna(v) else round(float(v), 2)
        out.append(rec)
    return out


def gap(r, a):
    return None if pd.isna(r[a]) or pd.isna(r["SAVER"]) else round(float(r[a] - r["SAVER"]), 2)


hikes26 = [(d, old, new) for d, old, new in decisions if d.year == 2026 and new > old]
e22 = next(e for e in ep_out if e["name"].startswith("2022"))
payload = {
    "updated": str(pd.Timestamp.today().date()),
    "dataTo": last.strftime("%b %Y"),
    "f6To": last_f6.strftime("%b %Y"),
    "capTo": cap.strftime("%b %Y"),
    "next": NEXT_DECISION,
    "overlay": series(m.loc["2021-01":last], {"cash": "CASH", "svr": "SVR", "disc": "DISC", "exist": "EXIST",
                                              "newl": "NEWL", "saver": "SAVER"}),
    "wedge": [{"d": str(p), "w1": gap(r, "DISC"), "w2": gap(r, "SVR"), "w3": gap(r, "EXIST")}
              for p, r in m.loc["2004-01":last].iterrows()],
    "episodes": ep_out,
    "cuts": cuts,
    "kpi": {
        "cash": float(latest_dec[2]),
        "cashFrom": f"{latest_dec[0].day} {latest_dec[0].strftime('%b')}",
        "hikes26": len(hikes26),
        "svr": float(m.loc[last, "SVR"]), "disc": float(m.loc[last, "DISC"]),
        "exist": float(m.loc[last_f6, "EXIST"]), "newl": float(m.loc[last_f6, "NEWL"]),
        "saver": float(m.loc[last, "SAVER"]), "bonus": float(m.loc[last, "BONUS"]),
        "wedge": round(float(m.loc[last, "DISC"] - m.loc[last, "SAVER"]), 2),
        "wedgePeak": round(float((m["DISC"] - m["SAVER"]).loc["2022-01":].max()), 2),
        "wedgeActual": round(float(m.loc[last_f6, "EXIST"] - m.loc[last_f6, "SAVER"]), 2),
        "asym": [e22["full"]["svr"], e22["full"]["exist"], e22["full"]["saver"]],
    },
    "hikes26": [[d.strftime("%b %d").replace(" 0", " "), round(new - old, 2)] for d, old, new in hikes26],
}

html = TEMPLATE.read_text(encoding="utf-8")
html = html.replace("/*__DATA__*/", "const DATA = " + json.dumps(payload) + ";")
(PROJECT / "PassThroughGap_Dashboard.html").write_text(html, encoding="utf-8")
# GitHub Pages serves index.html at the repo root.
(PROJECT / "index.html").write_text(html, encoding="utf-8")

# Artifact variant: the artifact host supplies the document skeleton.
body = re.sub(r"<!DOCTYPE html>\s*|<html[^>]*>\s*|</html>\s*|<head>\s*|</head>\s*|<body>\s*|</body>\s*|<meta [^>]*>\s*",
              "", html)
(PROJECT / "output" / "dashboard_artifact.html").write_text(body, encoding="utf-8")

print("saved PassThroughGap_Dashboard.html | retail data to", last, "| F6 to", last_f6, "| measured to", cap)
for e in ep_out + [cuts]:
    print(f"  {e['name']:<12} {e['delta']:>9}  fast {e['fast']}  full {e['full']}")
print("  KPI", payload["kpi"])
