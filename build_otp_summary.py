#!/usr/bin/env python3
"""
build_otp_summary.py
Produces data/ey_otp.json — a small, deployable OTP reliability dataset for the
EY analysis page.

The full scraper dump (output/otp_latest.json, ~200 KB) is gitignored and never
deployed. This script distills it to EY overall metrics + per-route D0/D15/A15,
which the page fetches to show "reliability beside the network".

Run from the network-site folder after an OTP scrape:
    python build_otp_summary.py
"""

import json
import os
import math
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
SRC  = os.path.join(HERE, "output", "otp_latest.json")
OUT  = os.path.join(HERE, "data", "ey_otp.json")


def num(x):
    """Round to 1dp; turn NaN/None/non-numeric into JSON null."""
    if x is None:
        return None
    if isinstance(x, float) and math.isnan(x):
        return None
    try:
        return round(float(x), 1)
    except (TypeError, ValueError):
        return None


def main():
    if not os.path.exists(SRC):
        raise SystemExit(f"Source not found: {SRC} (run the OTP scraper first)")

    with open(SRC, encoding="utf-8") as f:
        d = json.load(f)

    s = d.get("summary", [])
    overall = next((x for x in s if x.get("scope_type") == "OVERALL"), {})

    routes = []
    for x in s:
        if x.get("scope_type") != "ROUTE":
            continue
        if x.get("Date") not in ("ALL_DATES", None):
            continue
        rv = x.get("scope_value", "")
        parts = rv.split("-")
        routes.append({
            "r":    rv,
            "from": parts[0] if len(parts) == 2 else "",
            "to":   parts[1] if len(parts) == 2 else "",
            "n":    int(x.get("flights", 0) or 0),
            "d0":   num(x.get("D0_%")),
            "d15":  num(x.get("D15_%")),
            "a15":  num(x.get("A15_%")),
            "dep":  num(x.get("avg_dep_delay_min")),
            "arr":  num(x.get("avg_arr_delay_min")),
        })

    out = {
        "generated":      date.today().isoformat(),
        "run_date":       d.get("run_date"),
        "sample_date":    next((x.get("scope_value") for x in s if x.get("scope_type") == "DATE"), None),
        "sample_flights": d.get("ok") or d.get("total"),
        "airline":        d.get("airline", "EY"),
        "overall": {
            "d0":           num(overall.get("D0_%")),
            "d15":          num(overall.get("D15_%")),
            "a15":          num(overall.get("A15_%")),
            "within_block": num(overall.get("within_block_%")),
            "avg_dep":      num(overall.get("avg_dep_delay_min")),
            "avg_arr":      num(overall.get("avg_arr_delay_min")),
            "completed":    num(overall.get("completed_%")),
            "cancelled":    num(overall.get("cancelled_%")),
        },
        "routes": routes,
    }

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
        f.write("\n")

    o = out["overall"]
    print(f"ey_otp.json: {len(routes)} routes | sample {out['sample_flights']} flights on {out['sample_date']}")
    print(f"  overall  D0 {o['d0']}% · D15 {o['d15']}% · A15 {o['a15']}% · within-block {o['within_block']}% · avg dep {o['avg_dep']} min")


if __name__ == "__main__":
    main()
