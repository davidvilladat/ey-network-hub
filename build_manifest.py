#!/usr/bin/env python3
"""
build_manifest.py
Regenerates data/manifest.json — the single source of truth for the site's
"as-of" date labels and headline counts.

Counts (carriers / airports / agreements) are derived live from the data files,
so they can never drift. The editorial date strings (ZED date, FR24 window, etc.)
are preserved from the existing manifest unless overridden on the command line.

Run from the network-site folder after any data refresh:
    python build_manifest.py
    python build_manifest.py --zed "15 Jun 2026" --fr24 "12-14 Jun 2026"
"""

import json
import os
import argparse
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")

NETWORKS = os.path.join(DATA, "networks.json")
AIRPORTS = os.path.join(DATA, "airports.json")
AGREEMENTS = os.path.join(DATA, "zed_agreements.json")
MANIFEST = os.path.join(DATA, "manifest.json")

# Editorial defaults — used only if no existing manifest and no CLI override.
DEFAULTS = {
    "zed_agreements_updated": "30 Apr 2026",
    "fr24_sample_window": "1–3 May 2026",
    "ey_financials": "FY 2025 Annual Report",
    "cirium_window": "Q4 2025 vs Q4 2026",
}

NOTE = ("Single source of truth for the 'as-of' date labels shown on the site. "
        "Counts are derived live from the data files by build_manifest.py — do not hand-edit.")


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--zed", help="ZED agreements 'as-of' date label")
    ap.add_argument("--fr24", help="FR24 sample window label")
    ap.add_argument("--ey", help="EY financials label")
    ap.add_argument("--cirium", help="Cirium window label")
    args = ap.parse_args()

    # Start from existing manifest (preserve editorial fields) then DEFAULTS.
    editorial = dict(DEFAULTS)
    if os.path.exists(MANIFEST):
        try:
            prev = load_json(MANIFEST)
            for k in editorial:
                if k in prev:
                    editorial[k] = prev[k]
        except (ValueError, OSError):
            pass

    # CLI overrides
    if args.zed:    editorial["zed_agreements_updated"] = args.zed
    if args.fr24:   editorial["fr24_sample_window"] = args.fr24
    if args.ey:     editorial["ey_financials"] = args.ey
    if args.cirium: editorial["cirium_window"] = args.cirium

    networks = load_json(NETWORKS)
    airports = load_json(AIRPORTS)
    agreements = load_json(AGREEMENTS)

    # Derive the FR24 data-window date from the most common `win` in networks.json
    wins = [n.get("win") for n in networks.values() if isinstance(n, dict) and n.get("win")]
    fr24_win = max(set(wins), key=wins.count) if wins else None

    n_car = len(networks)
    n_air = len(airports)
    n_agr = len(agreements)
    covered = sum(1 for r in agreements if r.get("c") in networks)

    manifest = {
        "generated": date.today().isoformat(),
        "note": NOTE,
        "zed_agreements_updated": editorial["zed_agreements_updated"],
        "fr24_sample_window": editorial["fr24_sample_window"],
        "fr24_data_window": fr24_win,
        "ey_financials": editorial["ey_financials"],
        "cirium_window": editorial["cirium_window"],
        "counts": {
            "carriers_with_fr24": n_car,
            "airports": n_air,
            "zed_agreements": n_agr,
            "agreements_covered": covered,
            "agreements_data_blind": n_agr - covered,
            "coverage_pct": round(covered / n_agr * 100) if n_agr else 0,
        },
    }

    with open(MANIFEST, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
        f.write("\n")

    c = manifest["counts"]
    print(f"manifest.json regenerated ({manifest['generated']})")
    print(f"  carriers w/ FR24 : {n_car}")
    print(f"  airports         : {n_air}")
    print(f"  ZED agreements   : {n_agr}")
    print(f"  coverage         : {c['agreements_covered']}/{n_agr} ({c['coverage_pct']}%) · "
          f"{c['agreements_data_blind']} data-blind")
    print(f"  FR24 data window : {fr24_win}")


if __name__ == "__main__":
    main()
