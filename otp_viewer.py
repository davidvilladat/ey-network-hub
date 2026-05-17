#!/usr/bin/env python3
"""
OTP Viewer — Live progress monitor and results dashboard for Master Webscrapping CQ.

Run this in a separate terminal while the scraper is running, or after it finishes.
    python -X utf8 otp_viewer.py

Reads:
  output/scraper_status.json  — written every N flights by the scraper (live progress)
  output/otp_latest.json      — written at completion (full OTP results)
"""

import json
import os
import glob
import time
from datetime import datetime

from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.columns import Columns
from rich.text import Text
from rich.live import Live
from rich.layout import Layout
from rich.progress import Progress, BarColumn, TextColumn, MofNCompleteColumn
from rich import box

OUTPUT_DIR = "output"
STATUS_FILE = os.path.join(OUTPUT_DIR, "scraper_status.json")
OTP_FILE    = os.path.join(OUTPUT_DIR, "otp_latest.json")
REFRESH_SEC = 4

console = Console()


# ── Helpers ──────────────────────────────────────────────────────────────────

def _read_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _latest_excel():
    files = sorted(glob.glob(os.path.join(OUTPUT_DIR, "otp_report_*.xlsx")), reverse=True)
    return files[0] if files else None


def _fmt_pct(val, decimals=1):
    if val is None:
        return "[dim]—[/dim]"
    try:
        return f"{float(val):.{decimals}f}%"
    except Exception:
        return str(val)


def _color_pct(val, good=85, warn=70):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return "[dim]—[/dim]"
    color = "green" if v >= good else ("yellow" if v >= warn else "red")
    return f"[{color}]{v:.1f}%[/{color}]"


def _color_delay(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return "[dim]—[/dim]"
    color = "green" if v <= 0 else ("yellow" if v <= 15 else "red")
    sign  = "+" if v > 0 else ""
    return f"[{color}]{sign}{v:.1f} min[/{color}]"


# ── Live progress panel (shown while scraper is running) ──────────────────────

def build_progress_panel(st: dict) -> Panel:
    done    = st.get("done", 0)
    total   = st.get("total", 1)
    ok      = st.get("ok", 0)
    fail    = st.get("fail", 0)
    pct     = st.get("pct", 0.0)
    speed   = st.get("speed_per_min", 0.0)
    eta     = st.get("eta_min", 0.0)
    elapsed = st.get("elapsed_sec", 0.0)
    updated = st.get("last_updated", "")[:19].replace("T", " ")

    bar_filled = int(pct / 2)  # 50-char bar
    bar = "[green]" + "█" * bar_filled + "[/green][dim]" + "░" * (50 - bar_filled) + "[/dim]"

    ok_pct   = round(ok / max(done, 1) * 100, 1)
    fail_pct = round(fail / max(done, 1) * 100, 1)

    elapsed_str = f"{int(elapsed//60)}m {int(elapsed%60)}s"
    eta_str     = f"{int(eta)}m" if eta > 0 else "—"

    lines = Text()
    lines.append(f"\n  {bar}\n\n", style="")
    lines.append(f"  Flights    ", style="dim")
    lines.append(f"{done} / {total}   ", style="bold white")
    lines.append(f"({pct:.1f}%)\n", style="cyan")
    lines.append(f"  Success    ", style="dim")
    lines.append(f"{ok}  ", style="bold green")
    lines.append(f"({ok_pct}%)   ", style="green")
    lines.append(f"Failed  ", style="dim")
    lines.append(f"{fail}  ", style="bold red")
    lines.append(f"({fail_pct}%)\n\n", style="red")
    lines.append(f"  Speed      ", style="dim")
    lines.append(f"{speed:.1f} flights/min\n", style="yellow")
    lines.append(f"  Elapsed    ", style="dim")
    lines.append(f"{elapsed_str}\n", style="white")
    lines.append(f"  ETA        ", style="dim")
    lines.append(f"{eta_str}\n", style="white")
    lines.append(f"\n  Last update  {updated}", style="dim")

    return Panel(lines, title="[bold cyan]⏳  SCRAPING IN PROGRESS[/bold cyan]",
                 border_style="cyan", padding=(0, 1))


# ── OTP results panel (shown after run completes) ─────────────────────────────

def build_results_panels(otp: dict) -> list:
    panels = []
    summary_rows = otp.get("summary", [])

    # ── Run info ──
    run_text = Text()
    run_text.append(f"\n  Airline     ", style="dim")
    run_text.append(f"{otp.get('airline','?')}\n", style="bold white")
    run_text.append(f"  Flights OK  ", style="dim")
    run_text.append(f"{otp.get('ok','?')} / {otp.get('total','?')}\n", style="bold green")
    run_text.append(f"  Failed      ", style="dim")
    run_text.append(f"{otp.get('fail','?')}\n", style="bold red")
    run_text.append(f"  Duration    ", style="dim")
    run_text.append(f"{otp.get('elapsed_human','?')}\n", style="white")
    run_text.append(f"  Output      ", style="dim")
    fname = os.path.basename(otp.get("output_file",""))
    run_text.append(f"{fname}\n", style="dim yellow")
    panels.append(Panel(run_text, title="[bold green]✅  RUN COMPLETE[/bold green]",
                        border_style="green", padding=(0, 1)))

    if not summary_rows:
        return panels

    # Pull OVERALL row
    overall = next((r for r in summary_rows if r.get("scope_type") == "OVERALL"), None)
    if overall:
        # ── OTP metrics ──
        m = Table(box=box.SIMPLE, show_header=True, header_style="bold dim",
                  padding=(0, 2))
        m.add_column("Metric", style="dim", width=22)
        m.add_column("Value",  justify="right", width=10)
        m.add_column("",       width=2)
        m.add_column("Metric", style="dim", width=22)
        m.add_column("Value",  justify="right", width=10)

        dep_rows = [
            ("D+0  on time",  "D0_%"),
            ("D+15 <15 min",  "D15_%"),
            ("D+30 <30 min",  "D30_%"),
            ("D+60 <60 min",  "D60_%"),
            ("Avg dep delay", "avg_dep_delay_min"),
        ]
        arr_rows = [
            ("A+0  on time",  "A0_%"),
            ("A+15 <15 min",  "A15_%"),
            ("A+30 <30 min",  "A30_%"),
            ("A+60 <60 min",  "A60_%"),
            ("Avg arr delay", "avg_arr_delay_min"),
        ]
        for (dl, dk), (al, ak) in zip(dep_rows, arr_rows):
            dv = overall.get(dk)
            av = overall.get(ak)
            if "delay" in dk:
                dval = _color_delay(dv)
                aval = _color_delay(av)
            else:
                dval = _color_pct(dv)
                aval = _color_pct(av)
            m.add_row(dl, dval, "", al, aval)

        m.add_row("", "", "", "", "")
        wb = overall.get("within_block_%")
        m.add_row("Within block",  _color_pct(wb), "", "P90 arr delay",
                  _color_delay(overall.get("p90_arr_delay_min")))

        panels.append(Panel(m, title="[bold yellow]📊  OVERALL OTP — EY 2026-05-16[/bold yellow]",
                            border_style="yellow"))

    # ── Per-route worst delays ──
    route_rows = [r for r in summary_rows if r.get("scope_type") == "ROUTE"]
    if route_rows:
        route_rows.sort(key=lambda r: r.get("avg_arr_delay_min") or 0, reverse=True)
        rt = Table(box=box.SIMPLE, show_header=True, header_style="bold dim",
                   padding=(0, 1))
        rt.add_column("Route",      width=12)
        rt.add_column("Flights",    justify="right", width=7)
        rt.add_column("D+15%",      justify="right", width=7)
        rt.add_column("A+15%",      justify="right", width=7)
        rt.add_column("Avg dep",    justify="right", width=9)
        rt.add_column("Avg arr",    justify="right", width=9)
        for r in route_rows[:15]:
            rt.add_row(
                str(r.get("scope_value", "")),
                str(int(r.get("flights", 0))),
                _color_pct(r.get("D15_%")),
                _color_pct(r.get("A15_%")),
                _color_delay(r.get("avg_dep_delay_min")),
                _color_delay(r.get("avg_arr_delay_min")),
            )
        panels.append(Panel(rt, title="[bold magenta]🗺  TOP 15 ROUTES by Avg Arrival Delay[/bold magenta]",
                            border_style="magenta"))

    return panels


# ── Main loop ─────────────────────────────────────────────────────────────────

def main():
    console.clear()
    console.print(Panel(
        "[bold white]Etihad Airways OTP Dashboard[/bold white]\n"
        "[dim]Reads output/scraper_status.json and output/otp_latest.json[/dim]\n"
        "[dim]Refreshes every 4 seconds — Ctrl+C to exit[/dim]",
        border_style="blue", padding=(0, 2)
    ))

    try:
        while True:
            st  = _read_json(STATUS_FILE)
            otp = _read_json(OTP_FILE)

            console.clear()
            console.print(
                f"[bold blue]EY OTP Dashboard[/bold blue]  "
                f"[dim]{datetime.now().strftime('%H:%M:%S')}[/dim]"
            )
            console.print()

            if st is None and otp is None:
                console.print(Panel(
                    "[yellow]No scraper data found yet.[/yellow]\n\n"
                    f"[dim]Looking for: {os.path.abspath(STATUS_FILE)}\n"
                    "Start the scraper and this viewer will update automatically.[/dim]",
                    border_style="yellow"
                ))
            elif st and st.get("status") == "running":
                console.print(build_progress_panel(st))
            else:
                # Done or post-run
                if otp:
                    for panel in build_results_panels(otp):
                        console.print(panel)
                elif st:
                    console.print(Panel(
                        f"[green]Run complete[/green]  "
                        f"{st.get('ok','?')} OK / {st.get('fail','?')} failed\n"
                        f"[dim]Load {_latest_excel()} for full details[/dim]",
                        border_style="green"
                    ))
                if st:
                    console.print(f"[dim]Last run: {st.get('elapsed_human','?')} · "
                                  f"Output: {os.path.basename(st.get('output_file',''))}[/dim]")

            time.sleep(REFRESH_SEC)

    except KeyboardInterrupt:
        console.print("\n[dim]Viewer closed.[/dim]")


if __name__ == "__main__":
    main()
