#!/usr/bin/env python3
"""
================================================================================
 Master Webscrapping CQ — v1.0
 FlightAware Airline OTP Scraper
================================================================================

 Author   : Camilo Quiroga (AEG Fuels Revenue Management)
 Purpose  : Extract on-time performance (OTP) data from FlightAware for any
            airline in a Cirium schedule file, and produce an analysis-ready
            Excel report with D+0/D+15/D+30/D+60/D+360 and A+0/A+14/A+15/A+30/
            A+60/A+360 metrics at multiple aggregation levels.

 HIGH-LEVEL OVERVIEW
 -------------------
 1. Load airports.csv (ourairports.com) → IATA↔ICAO mapping + longitude-based
    UTC offset estimate (used to build Zulu URLs).
 2. Load the Cirium schedule XLSX → list of flights operating the target date(s).
 3. For each flight+date, build a FlightAware history URL WITH pre-computed Zulu
    time so FlightAware returns the exact historical flight on the first try:
      https://www.flightaware.com/live/flight/ETD101/history/20260416/1020Z/OMAA/EGLL
    Pattern: /live/flight/{ICAO_AIRLINE}{FLIGHT_NUM}/history/{YYYYMMDD}/{HHMM}Z/{ORIGIN_ICAO}/{DEST_ICAO}
 4. Open each URL in parallel headless browsers, parse the rendered HTML.
 5. Compute OTP metrics per flight (with Cirium Block_Mins as primary source
    for within-block calculations).
 6. Output a single Excel file with four sheets:
      - Summary       : aggregated metrics (scope_type × scope_value × Date)
      - All_Flights   : every flight with full data
      - Run_Info      : runtime stats (OTP Agents, time, throughput)
      - Failed_Scrapes: diagnostics for any flights that didn't return data

 WHAT MAKES THE ZULU URL IMPORTANT
 ---------------------------------
 FlightAware's URL without a Zulu time (/history/20260416/OMAA/EGLL) sometimes
 returns a generic view for the NEXT scheduled flight instead of the one that
 actually operated — reporting 0 delays because the flight hasn't happened yet.
 We pre-compute the Zulu time from Cirium's local Dep_Time and the origin
 airport's longitude-derived UTC offset, so the URL is unambiguous from the
 first request. A thin-data detection + history-picker fallback covers the
 ~1% of cases where the computed Zulu is off by a slot.

 USAGE
 -----
   - Place SummerS_EY_QR_EK.xlsx (or equivalent) next to this script.
   - Run from Spyder: %runfile "Master Webscrapping CQ.py"
   - Run from shell: python "Master Webscrapping CQ.py"
   - The script is fully interactive: no command-line flags.

   On startup the script will prompt you for:
     Reference date (YYYY-MM-DD):  2026-04-17
     Days back (0=same day):       1
     Number of OTP Agents:         (recommended based on CPU/RAM)
     Flights per day (ENTER=all):  50

   With date=2026-04-17 and days_back=1, it scrapes April 17 AND April 16.
   With date=2026-04-17 and days_back=0, it scrapes only April 17.

 REQUIREMENTS
 ------------
   pip install -r requirements.txt
   (pandas, openpyxl, beautifulsoup4, selenium, webdriver-manager, psutil)

 OUTPUTS
 -------
   output/otp_report_YYYYMMDD_HHMM.xlsx   Single Excel workbook with 4 sheets.
================================================================================
"""

import os
import sys
import re
import time
import random
import logging
import subprocess
from datetime import datetime, timedelta
from typing import Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
import threading


# Silence webdriver-manager verbose logs ("WebDriver manager" banners per worker).
# Must be set BEFORE webdriver_manager is imported anywhere.
os.environ.setdefault("WDM_LOG", "0")
os.environ.setdefault("WDM_LOG_LEVEL", "0")
os.environ.setdefault("WDM_PRINT_FIRST_LINE", "False")

# ──────────────────────────────────────────────────────────────────────────────
# DEPENDENCY CHECK — Verify all required libraries are installed
# ──────────────────────────────────────────────────────────────────────────────
# Before importing any third-party library, we check if it's available.
# If something is missing, we attempt to install it from requirements.txt.
# This prevents confusing ImportError messages for users who haven't set up
# their environment yet.

REQUIRED_PACKAGES = {
    # { import_name: pip_package_name }
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "bs4": "beautifulsoup4",
    "selenium": "selenium",
}

# webdriver-manager auto-downloads the correct chromedriver for your Chrome version.
# It's not strictly required (Selenium can find chromedriver in PATH) but it makes
# setup much easier on Windows.
OPTIONAL_PACKAGES = {
    "webdriver_manager": "webdriver-manager",
    "psutil": "psutil",   # Used for smart OTP Agents count suggestion
}

def check_dependencies():
    """
    Verify that all required Python packages are importable.
    If any are missing, attempt to install them via pip using requirements.txt.
    Selenium is the required browser engine (works everywhere including Spyder).
    Playwright is optional (faster but doesn't work in all environments).
    """
    missing = []

    # Check required packages (includes Selenium)
    for import_name, pip_name in REQUIRED_PACKAGES.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)

    # Check optional packages
    for import_name, pip_name in OPTIONAL_PACKAGES.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)

    if not missing:
        return  # All good

    # ── Attempt auto-install ──
    print(f"\n  Missing packages: {', '.join(missing)}")
    print(f"  Attempting to install...\n")

    req_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "requirements.txt")

    if os.path.exists(req_file):
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", req_file],
            capture_output=True, text=True
        )
    else:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install"] + missing,
            capture_output=True, text=True
        )

    if result.returncode != 0:
        print(f"  Auto-install failed. Please install manually:")
        print(f"  pip install -r requirements.txt")
        print(f"\n  Error: {result.stderr[:500]}")
        sys.exit(1)

    print("  All dependencies installed successfully.\n")

    # Re-verify required packages
    for import_name in REQUIRED_PACKAGES:
        try:
            __import__(import_name)
        except ImportError:
            print(f"  Failed to import {import_name} after install. Check your Python environment.")
            sys.exit(1)


# Run the check immediately — before any third-party imports
check_dependencies()

# ── Now safe to import third-party libraries ──
import pandas as pd
from bs4 import BeautifulSoup


# ──────────────────────────────────────────────────────────────────────────────
# LOGGING — Clean, minimal, informative
# ──────────────────────────────────────────────────────────────────────────────
# We use a single logger. Progress updates (flight count, ETA) go to INFO.
# Individual flight results are logged at DEBUG to avoid flooding the console.
# To see per-flight detail, run with --verbose.

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s │ %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("scraper")


# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────────────────────────────────────

@dataclass
class Config:
    max_workers: int = 5
    min_delay_sec: float = 2.0
    max_delay_sec: float = 5.0
    page_timeout_sec: int = 45     # FlightAware + Cloudflare challenge can take up to 40s
    max_retries: int = 3           # Bumped from 2 — gives timeout failures another chance
    headless: bool = True
    output_dir: str = "output"
    user_agent: str = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
    # Filters
    airline_filter: str = ""           # e.g. "AV" — empty means all airlines
    routes_filter: list = field(default_factory=list)

# ICAO airline code mapping
# FlightAware URLs use ICAO airline codes (3 letters), not IATA (2 letters).
# Example: Etihad is "EY" in IATA but "ETD" in ICAO.
# This map converts the 2-letter code from Cirium to the 3-letter code for URLs.
IATA_TO_ICAO_AIRLINE = {
    "EY": "ETD",   # Etihad Airways
    "EK": "UAE",   # Emirates
    "QR": "QTR",   # Qatar Airways
}


# ──────────────────────────────────────────────────────────────────────────────
# PROGRESS TRACKER
# ──────────────────────────────────────────────────────────────────────────────
# Shared across all worker threads. Provides real-time counts and ETA.

class ProgressTracker:
    """
    Thread-safe progress tracker that prints a status line showing:
      - How many flights have been scraped so far
      - How many succeeded vs failed
      - INSTANT speed (last 30-flight rolling window, not cumulative average)
      - ETA based on that instant speed

    Why rolling window instead of cumulative average:
    -------------------------------------------------
    The first ~50 flights are slow because the offset cache is empty — every
    flight does 2 page loads instead of 1. After the cache warms up, speed
    doubles. Reporting cumulative average would under-estimate real speed
    most of the run and over-estimate at the end if things slow down.

    A 30-flight sliding window gives you a "what's happening RIGHT NOW" speed.
    """

    def __init__(self, total: int, update_every: int = 10, status_file: str = None):
        self.total = total
        self.update_every = update_every
        self._status_file = status_file
        self._lock = threading.Lock()
        self._done = 0
        self._ok = 0
        self._fail = 0
        self._start_time = time.time()
        # Rolling window of completion timestamps — used for instant speed
        self._recent_timestamps = []
        self._window_size = 30

    def record(self, success: bool):
        """Called by each worker after scraping one flight."""
        now = time.time()
        with self._lock:
            self._done += 1
            if success:
                self._ok += 1
            else:
                self._fail += 1

            # Keep only the last N timestamps for rolling speed calculation
            self._recent_timestamps.append(now)
            if len(self._recent_timestamps) > self._window_size:
                self._recent_timestamps.pop(0)

            if self._done % self.update_every == 0 or self._done == self.total:
                self._print_status()

    def _instant_rate(self) -> float:
        """Flights per second over the rolling window. Returns 0 if too few samples."""
        if len(self._recent_timestamps) < 2:
            return 0
        window_sec = self._recent_timestamps[-1] - self._recent_timestamps[0]
        if window_sec <= 0:
            return 0
        return (len(self._recent_timestamps) - 1) / window_sec

    def _print_status(self):
        import json as _json
        elapsed = time.time() - self._start_time
        rate = self._instant_rate()
        remaining = (self.total - self._done) / rate if rate > 0 else 0

        pct = self._done / self.total * 100
        eta_min = remaining / 60

        log.info(
            f"Progress: {self._done}/{self.total} ({pct:.0f}%) │ "
            f"OK: {self._ok}  Fail: {self._fail} │ "
            f"Speed: {rate:.1f} flights/sec │ "
            f"ETA: {eta_min:.1f} min"
        )

        if self._status_file:
            try:
                with open(self._status_file, "w") as _f:
                    _json.dump({
                        "status": "running",
                        "total": self.total,
                        "done": self._done,
                        "ok": self._ok,
                        "fail": self._fail,
                        "pct": round(pct, 1),
                        "speed_per_min": round(rate * 60, 1),
                        "eta_min": round(eta_min, 1),
                        "elapsed_sec": round(elapsed, 1),
                        "last_updated": datetime.now().isoformat(),
                    }, _f)
            except Exception:
                pass

    def summary(self) -> dict:
        elapsed = time.time() - self._start_time
        return {
            "total": self.total,
            "done": self._done,
            "ok": self._ok,
            "fail": self._fail,
            "elapsed_sec": round(elapsed, 1),
        }


# ──────────────────────────────────────────────────────────────────────────────
# AIRPORTS LOOKUP — Build IATA↔ICAO mapping from airports.csv
# ──────────────────────────────────────────────────────────────────────────────

def load_airport_lookup(filepath: str) -> dict:
    """
    Load airports.csv from ourairports.com and build a simple IATA ↔ ICAO
    lookup dictionary. No UTC offset is stored here — offsets come from
    FlightAware's own HTML at runtime (see OFFSET_CACHE in the scraper).

    Returns:
        dict with both directions: {"AUH": "OMAA", "OMAA": "AUH", ...}
    """
    log.info(f"Loading airport codes from {filepath}")

    df = pd.read_csv(filepath, usecols=["iata_code", "icao_code"], dtype=str)
    df = df.dropna(subset=["iata_code", "icao_code"])
    df["iata_code"] = df["iata_code"].str.strip().str.upper()
    df["icao_code"] = df["icao_code"].str.strip().str.upper()
    df = df[(df["iata_code"] != "") & (df["icao_code"] != "")]

    lookup = {}
    for _, row in df.iterrows():
        iata = row["iata_code"]
        icao = row["icao_code"]
        lookup[iata] = icao
        lookup[icao] = iata

    log.info(f"Loaded {len(df)} airport code pairs")
    return lookup


def load_airport_coord_offsets(filepath: str) -> dict:
    """
    Build a rough UTC offset lookup keyed by ICAO/IATA code using airport longitude.
    offset_hours = round(longitude / 15)
    Accuracy: ±1 hour for most airports (India/Iran/Nepal and other non-integer-hour
    zones may deviate by 30 min). Superseded by OFFSET_CACHE (scraped from FlightAware)
    and HUB_UTC_OFFSETS (exact values for Gulf hubs).
    Used by run_fr24_otp() to convert Cirium local scheduled times to UTC.
    """
    try:
        df = pd.read_csv(
            filepath,
            usecols=["icao_code", "iata_code", "longitude_deg"],
            dtype=str,
        )
        df["longitude_deg"] = pd.to_numeric(df["longitude_deg"], errors="coerce")
        df = df.dropna(subset=["longitude_deg"])
        result = {}
        for _, row in df.iterrows():
            offset = int(round(float(row["longitude_deg"]) / 15))
            if pd.notna(row.get("icao_code")) and str(row["icao_code"]).strip():
                result[str(row["icao_code"]).strip().upper()] = offset
            if pd.notna(row.get("iata_code")) and str(row["iata_code"]).strip():
                result[str(row["iata_code"]).strip().upper()] = offset
        return result
    except Exception as e:
        log.warning(f"Could not load coordinate offsets from {filepath}: {e}")
        return {}


def to_icao(code: str, lookup: dict) -> str:
    """Convert any airport code to ICAO (4-letter), or return as-is."""
    code = code.strip().upper()
    if len(code) == 4:
        return code
    return lookup.get(code, code)


# ──────────────────────────────────────────────────────────────────────────────
# UTC OFFSET CACHE & HELPERS
# ──────────────────────────────────────────────────────────────────────────────
# FlightAware shows the UTC offset next to each local time on the flight page
# (e.g. "05:20AM -05" for Bogotá). We extract that offset from the HTML once
# per origin airport and cache it here. All subsequent flights from the same
# airport build their Zulu URL directly using the cached offset, saving one
# round-trip per flight.
#
# The cache is a process-wide dict keyed by origin ICAO (e.g. "OMAA" -> +4).
# A lock protects it against concurrent writes from parallel workers.

OFFSET_CACHE = {}
OFFSET_CACHE_LOCK = threading.Lock()

# Pre-known UTC offsets for Gulf carrier hubs — no DST, stable year-round.
# Pre-warming these eliminates the double-page-load for all hub departures
# (the majority of EY/EK/QR legs originate at their respective hubs).
HUB_UTC_OFFSETS = {
    "OMAA": +4,   # AUH — Abu Dhabi  (Etihad hub)
    "OMDB": +4,   # DXB — Dubai      (Emirates hub)
    "OTHH": +3,   # DOH — Doha       (Qatar Airways hub)
}


def extract_utc_offset_from_html(html: str) -> Optional[int]:
    """
    Find the UTC offset of the ORIGIN airport in a FlightAware page.

    FlightAware renders each local time with a "±HH" suffix like "05:20AM -05".
    The page has MULTIPLE such times (departure and arrival), so we can't just
    pick the first regex match — the destination's offset would sometimes win.

    Strategy: look specifically for the "Scheduled" takeoff block. FlightAware
    puts the origin's scheduled takeoff time and offset together as:
        Scheduled 05:20AM -05
    We anchor on "Scheduled" to get the right one. If that fails, fall back to
    the first "AM/PM ±HH" match (which is usually the origin's gate departure).

    Returns an integer hour offset (e.g. -5, +2) or None if nothing was found.
    """
    if not html:
        return None

    # Primary: "Scheduled HH:MM(AM|PM) ±HH" — this is the origin's scheduled takeoff
    m = re.search(r"Scheduled\s*\d{1,2}:\d{2}\s*[AP]M\s*([+-]\d{1,2})(?![0-9])", html, re.I)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass

    # Fallback: first "HH:MM(AM|PM) ±HH" on the page (usually origin departure)
    m = re.search(r"\d{1,2}:\d{2}\s*[AP]M\s*([+-]\d{1,2})(?![0-9])", html)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass

    return None


def local_to_zulu(local_hhmm, utc_offset_hours: Optional[int]) -> Optional[str]:
    """
    Convert local HHMM + UTC offset to 4-digit Zulu string "HHMM".

    Example:
      local=1710 (5:10 PM in Bogotá), offset=-5
      → 17:10 - (-5) = 22:10 = "2210"

    Matches FlightAware's URL pattern exactly. Returns None if inputs invalid.
    """
    if local_hhmm is None or utc_offset_hours is None:
        return None
    try:
        t = int(local_hhmm)
    except (ValueError, TypeError):
        return None
    hh, mm = t // 100, t % 100
    if not (0 <= hh <= 23 and 0 <= mm <= 59):
        return None
    # UTC = Local - Offset (so Bogotá -5 → add 5 hours)
    total_min = (hh * 60 + mm) - utc_offset_hours * 60
    total_min %= 1440
    return f"{total_min // 60:02d}{total_min % 60:02d}"


# ──────────────────────────────────────────────────────────────────────────────
# CLOUDFLARE CHALLENGE HANDLER
# ──────────────────────────────────────────────────────────────────────────────
# FlightAware sits behind Cloudflare's bot protection. When Cloudflare is
# suspicious of your traffic, it returns an interstitial page titled
# "Just a moment..." with JavaScript that must run for ~5 seconds before
# the real page loads.
#
# Symptoms if you don't handle this:
#   - Scraper hangs waiting for .flightStatusBig (never appears on the
#     challenge page)
#   - `parse_flight_html` returns None because the HTML is just the challenge
#   - Workers appear "stuck" for 30 seconds, then time out
#
# The fix: detect the challenge by its title, then wait for the real page
# title to appear (FlightAware's real pages have the flight number in the
# title). Give it up to 30 seconds — Cloudflare usually resolves in 5-8.

def is_cloudflare_challenge(title_or_html: str) -> bool:
    """Detect Cloudflare's 'Just a moment...' interstitial page."""
    if not title_or_html:
        return False
    lower = title_or_html[:500].lower()
    return (
        "just a moment" in lower or
        "checking your browser" in lower or
        "attention required | cloudflare" in lower or
        "cf-chl-" in lower     # Cloudflare challenge DOM markers
    )


def wait_past_cloudflare_selenium(driver, max_wait_sec: int = 30) -> bool:
    """
    If the current page is a Cloudflare challenge, wait for it to resolve.
    Returns True if we're past the challenge, False if still stuck after max_wait.
    """
    start = time.time()
    while time.time() - start < max_wait_sec:
        title = driver.title or ""
        if not is_cloudflare_challenge(title):
            return True   # Clean page — we're past the challenge
        time.sleep(1)     # Poll once per second
    return False          # Still on challenge page after max_wait


def wait_past_cloudflare_playwright(page, max_wait_sec: int = 30) -> bool:
    """Same as above but for Playwright's page object."""
    start = time.time()
    while time.time() - start < max_wait_sec:
        try:
            title = page.title() or ""
        except Exception:
            title = ""
        if not is_cloudflare_challenge(title):
            return True
        time.sleep(1)
    return False


# ──────────────────────────────────────────────────────────────────────────────
# CIRIUM SCHEDULE LOADER
# ──────────────────────────────────────────────────────────────────────────────

def load_cirium_schedule(filepath: str, dates: list[str], config: Config) -> pd.DataFrame:
    """
    Load the Cirium schedule XLSX and build the list of flights to scrape.

    LOGIC:
      1. Load the full Cirium file (all months of the season).
      2. Filter to the target dates.
      3. Each row IS a flight-segment that operates that day. Keep them 1-to-1.

    FILE FORMAT (SummerS_EY_QR_EK.xlsx):
      The file has a proper header row (row 0). Key columns used:
        Travel Date, Op Airline Code, Origin Code, Destination Code,
        Equipment Code, Dep Time, Arr Time, Dep Term, Arr Term,
        Flight, Block Mins, Seats, Total Kilometers, Orig WAC, Dest WAC,
        Arr Flag, Alliance, Seats - Business Class, Seats - Economy, etc.

      For EY/EK/QR, Op Airline Code is both the marketing and operating carrier
      (no subsidiary operators), so Mkt_Al = Op_Al = Op Airline Code.

    Returns:
        DataFrame with one row per (flight × date) to scrape
    """
    log.info(f"Loading Cirium schedule: {filepath}")

    df = pd.read_excel(filepath, header=0)

    # ── Rename to internal column names used throughout the script ──
    df = df.rename(columns={
        "Travel Date":        "Date",
        "Op Airline Code":    "Mkt_Al",
        "Alliance":           "Alliance",
        "Origin Code":        "Orig",
        "Destination Code":   "Dest",
        "Equipment Code":     "Equip",
        "Dep Time":           "Dep_Time",
        "Arr Time":           "Arr_Time",
        "Dep Term":           "Dep_Term",
        "Arr Term":           "Arr_Term",
        "Flight":             "Flight",
        "Block Mins":         "Block_Mins",
        "Seats":              "Seats",
        "Total Kilometers":   "Kilometers",
        "Orig WAC":           "Orig_WAC",
        "Dest WAC":           "Dest_WAC",
        "Arr Flag":           "Arr_Flag",
        # Cabin detail columns — kept in output for analysis
        "Seats - First Class":    "Seats_F",
        "Seats - Business Class": "Seats_C",
        "Seats - Prem Econ":      "Seats_W",
        "Seats - Economy":        "Seats_Y",
        "ASKs":                   "ASKs",
    })

    # EY/EK/QR are direct operators — no subsidiary relationships.
    # Create Op_Al as an alias of Mkt_Al so the rest of the code
    # (which expects both columns) works without modification.
    df["Op_Al"] = df["Mkt_Al"]

    # ── Clean dates ──
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])

    # ── Normalize string columns ──
    for col in ["Mkt_Al", "Op_Al", "Orig", "Dest", "Equip"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()

    # ── Numeric columns ──
    df["Flight"]     = pd.to_numeric(df["Flight"],     errors="coerce")
    df["Block_Mins"] = pd.to_numeric(df["Block_Mins"], errors="coerce")
    df = df.dropna(subset=["Flight"])

    log.info(f"Schedule loaded: {len(df):,} total schedule rows")

    # ── Filter to target dates ──
    target_dates = set()
    for d in dates:
        target_dates.add(pd.Timestamp(f"{d[:4]}-{d[4:6]}-{d[6:8]}"))
    df = df[df["Date"].isin(target_dates)].copy()
    log.info(f"After date filter ({len(dates)} date{'s' if len(dates)>1 else ''}): {len(df):,} flight-segments")

    # ── Apply airline filter (optional — empty = all three carriers) ──
    if config.airline_filter:
        al = config.airline_filter.upper()
        df = df[df["Mkt_Al"] == al].copy()
        log.info(f"After airline filter (Mkt_Al={al}): {len(df):,} flight-segments")

    # ── Apply route filter ──
    if config.routes_filter:
        df["_route"] = df["Orig"] + "-" + df["Dest"]
        df = df[df["_route"].isin(config.routes_filter)].copy()
        df.drop(columns=["_route"], inplace=True)
        log.info(f"After route filter: {len(df):,} flight-segments")

    if df.empty:
        log.warning("No flights match the filters.")
        return df

    # ── Build FlightAware flight identifier ──
    # FlightAware URLs use ICAO airline codes (3 letters).
    # EY→ETD, EK→UAE, QR→QTR (see IATA_TO_ICAO_AIRLINE above).
    df["icao_airline"] = df["Mkt_Al"].map(IATA_TO_ICAO_AIRLINE).fillna(df["Mkt_Al"])
    df["flight_fa"]    = df["icao_airline"] + df["Flight"].astype(int).astype(str)
    df["date_str"]     = df["Date"].dt.strftime("%Y%m%d")

    # ── Summary log ──
    log.info(f"Ready to scrape: {len(df):,} flight-segments across {df['Mkt_Al'].nunique()} airline(s)")
    for al, count in df["Mkt_Al"].value_counts().items():
        log.info(f"  {al}: {count:,} flights")

    return df


# ──────────────────────────────────────────────────────────────────────────────
# URL BUILDER
# ──────────────────────────────────────────────────────────────────────────────

def build_url(flight_fa: str, date_str: str, orig_icao: str, dest_icao: str) -> str:
    """
    Construct a FlightAware flight history URL.

    The URL pattern is:
      https://www.flightaware.com/live/flight/{ICAO_FLIGHT}/history/{YYYYMMDD}/{ORIG_ICAO}/{DEST_ICAO}

    Example:
      flight_fa  = "ETD101"
      date_str   = "20260418"
      orig_icao  = "OMAA"
      dest_icao  = "EGLL"
      → https://www.flightaware.com/live/flight/ETD101/history/20260418/OMAA/EGLL
    """
    return (
        f"https://www.flightaware.com/live/flight/"
        f"{flight_fa}/history/{date_str}/{orig_icao}/{dest_icao}"
    )


def _select_best_zulu(zulu_times: list, task: dict) -> Optional[str]:
    """
    Pick the best Zulu time from a list of candidates on the FlightAware page.

    Two comparison strategies, tried in order:
      1. If we pre-computed a Zulu time for this task (from Cirium local time
         + airport UTC offset), compare candidates directly in UTC space —
         this is accurate.
      2. Otherwise, fall back to comparing against the Cirium local Dep_Time
         as minutes-of-day (ignoring UTC offset) — less accurate but still
         picks between e.g. a 06:00 and a 20:00 flight correctly.

    Logic:
      - 0 candidates → None
      - 1 candidate  → return it
      - 2+ candidates → pick the one closest to our reference time
      - Ties broken by original order (first match wins)

    Deduplicates the input list first.
    """
    if not zulu_times:
        return None

    unique = list(dict.fromkeys(zulu_times))
    if len(unique) == 1:
        return unique[0]

    def hhmm_to_min(s):
        """Parse '1020' or '1020Z' → 620 minutes (10*60 + 20). Returns None on fail."""
        try:
            s = s[:4] if s[-1:].upper() == "Z" and len(s) == 5 else s[:4]
            return int(s[:2]) * 60 + int(s[2:4])
        except (ValueError, IndexError, AttributeError):
            return None

    # Prefer our pre-computed Zulu (accurate UTC-space comparison)
    ref_min = None
    computed = task.get("zulu_time_computed")   # "1020Z" or None
    if computed:
        ref_min = hhmm_to_min(computed)

    # Fallback: compare against Cirium local dep_time (treating it as UTC —
    # not accurate for distant longitudes but robust for picking between
    # morning vs afternoon flights)
    if ref_min is None:
        sched = task.get("sched_dep_local")
        if sched is not None:
            try:
                sched_str = str(int(sched)).zfill(4)
                ref_min = int(sched_str[:2]) * 60 + int(sched_str[2:])
            except (ValueError, TypeError):
                pass

    if ref_min is None:
        return unique[0]   # No reference → first candidate

    def minute_distance(zulu):
        zm = hhmm_to_min(zulu)
        if zm is None:
            return 10**9
        diff = abs(zm - ref_min)
        return min(diff, 1440 - diff)   # Wrap around midnight

    return min(unique, key=minute_distance)


# ──────────────────────────────────────────────────────────────────────────────
# HTML PARSER — Extract flight data from a FlightAware page
# ──────────────────────────────────────────────────────────────────────────────

def parse_flight_html(html: str) -> Optional[dict]:
    """
    Parse a FlightAware flight page and extract all timing data.

    Uses TWO strategies in order of reliability:

    STRATEGY 1 — JavaScript data object (most reliable):
      FlightAware embeds flight data in a JS variable called 'trackpollBootstrap'.
      This contains epoch timestamps which are timezone-unambiguous.
      We extract these with regex on the script tag contents.

    STRATEGY 2 — Visible HTML text (fallback):
      If the JS object is missing or incomplete, we fall back to parsing the
      visible text on the page. The "Flight Details" panel shows times like:
        Gate Departure: 09:24AM -05 / Scheduled 09:00AM -05
        Takeoff: 09:44AM -05 / Scheduled 09:10AM -05
        Landing: 02:36PM EDT / Scheduled 02:19PM EDT
        Gate Arrival: 02:48PM EDT / Scheduled 02:20PM EDT

    Returns:
      dict with all extracted fields, or None if the page has no flight data.
    """
    soup = BeautifulSoup(html, "html.parser")
    page_text = soup.get_text("\n", strip=True)
    data = {}

    # ── Quick exit: no data on page ──
    if any(phrase in page_text for phrase in [
        "not in our database", "flight not found", "Flight Not Found"
    ]):
        return None

    # ── Flight status ──
    # We look for the status ONLY in the prominent status banner at the top
    # of the page (element with class flightStatusBig/statusBig), not in
    # general page text — the word "cancel" appears in footer menus,
    # "Cancel alerts" links, cookie banners, etc., and those would create
    # false positives if we searched the whole page.
    status_el = soup.select_one("span.flightStatusBig, span.statusBig, .flightPageStatus, .flightPageSummary-title")
    status_text = status_el.get_text(" ", strip=True) if status_el else ""
    data["status_raw"] = status_text

    status_upper = status_text.upper()
    if "CANCEL" in status_upper:
        data["flight_status"] = "CANCELLED"
    elif "DIVERT" in status_upper:
        data["flight_status"] = "DIVERTED"
    elif "EN ROUTE" in status_upper or "IN FLIGHT" in status_upper:
        data["flight_status"] = "EN_ROUTE"
    elif "ARRIVED" in status_upper or "LANDED" in status_upper:
        data["flight_status"] = "ARRIVED"
    elif "SCHEDULED" in status_upper:
        data["flight_status"] = "SCHEDULED"
    elif status_text:
        data["flight_status"] = "UNKNOWN"
    else:
        # If we can't find a status banner but the page has times, assume completed
        data["flight_status"] = "UNKNOWN"

    # ═══════════════════════════════════════════════════════════════════════
    # STRATEGY 1: JavaScript trackpollBootstrap object
    # ═══════════════════════════════════════════════════════════════════════
    # FlightAware renders the page with React/JS and embeds the raw data in
    # a script tag. Epoch timestamps are integers (seconds since 1970-01-01).
    # This is the best source because it's timezone-unambiguous.

    for script in soup.find_all("script"):
        txt = script.string or ""
        if "trackpollBootstrap" not in txt:
            continue

        data["_parse_source"] = "js"

        def js_val(field_name: str) -> Optional[str]:
            """Extract a named value from the JS object using regex."""
            for pat in [
                rf'"{field_name}"\s*:\s*"([^"]*)"',    # string value
                rf'"{field_name}"\s*:\s*(-?\d+\.?\d*)',  # numeric value
            ]:
                m = re.search(pat, txt)
                if m:
                    return m.group(1)
            return None

        # Departure timestamps (epoch seconds)
        data["dep_gate_actual_epoch"] = js_val("gateDepartureTime") or js_val("actualdeparturetime")
        data["dep_gate_sched_epoch"] = js_val("scheduledDepartureTime") or js_val("filed_departuretime")
        data["dep_takeoff_actual_epoch"] = js_val("takeoffTime")
        data["dep_takeoff_sched_epoch"] = js_val("scheduledTakeoffTime")

        # Arrival timestamps (epoch seconds)
        data["arr_landing_actual_epoch"] = js_val("landingTime")
        data["arr_landing_sched_epoch"] = js_val("scheduledLandingTime")
        data["arr_gate_actual_epoch"] = js_val("gateArrivalTime") or js_val("actualarrivaltime")
        data["arr_gate_sched_epoch"] = js_val("scheduledArrivalTime") or js_val("filed_arrivaltime")

        # Aircraft & route info
        data["aircraft_type_fa"] = js_val("aircrafttype")
        data["tail_number"] = js_val("registration")
        data["distance_sm"] = js_val("distance")

        break

    # ═══════════════════════════════════════════════════════════════════════
    # STRATEGY 2: Visible HTML text parsing
    # ═══════════════════════════════════════════════════════════════════════
    # The Flight Details panel on the right side of the page shows times in
    # human-readable format. We use regex to extract them as a fallback or
    # supplement to the JS data.

    if "_parse_source" not in data:
        data["_parse_source"] = "html"

    # Each block follows the pattern:
    #   Label\n ActualTime\n ... Scheduled ScheduledTime
    time_patterns = {
        "dep_gate":    r"Gate Departure\s*\n\s*([\d:]+\s*[AP]M\s*[\w\-\+\d]*)\s*\n.*?Scheduled\s+([\d:]+\s*[AP]M\s*[\w\-\+\d]*)",
        "dep_takeoff": r"Takeoff\s*\n\s*([\d:]+\s*[AP]M\s*[\w\-\+\d]*)\s*\n.*?Scheduled\s+([\d:]+\s*[AP]M\s*[\w\-\+\d]*)",
        "arr_landing": r"Landing\s*\n\s*([\d:]+\s*[AP]M\s*[\w\-\+\d]*)\s*\n.*?Scheduled\s+([\d:]+\s*[AP]M\s*[\w\-\+\d]*)",
        "arr_gate":    r"Gate Arrival\s*\n\s*([\d:]+\s*[AP]M\s*[\w\-\+\d]*)\s*\n.*?Scheduled\s+([\d:]+\s*[AP]M\s*[\w\-\+\d]*)",
    }

    for key, pattern in time_patterns.items():
        m = re.search(pattern, page_text, re.DOTALL)
        if m:
            data.setdefault(f"{key}_actual_text", m.group(1).strip())
            data.setdefault(f"{key}_sched_text", m.group(2).strip())

    # Arrival delay from the hero section (e.g. "(28 minutes late)" or "(on time)")
    delay_m = re.search(r"\((\d+)\s*minutes?\s*late\)", page_text, re.I)
    if delay_m:
        data["arrival_delay_text_min"] = int(delay_m.group(1))
    elif re.search(r"\(on time\)", page_text, re.I):
        data["arrival_delay_text_min"] = 0
    early_m = re.search(r"\((\d+)\s*minutes?\s*early\)", page_text, re.I)
    if early_m:
        data["arrival_delay_text_min"] = -int(early_m.group(1))

    # Total travel time (e.g. "4h 24m total travel time")
    m = re.search(r"(\d+)h\s*(\d+)m\s*total travel time", page_text, re.I)
    if m:
        data["actual_travel_min"] = int(m.group(1)) * 60 + int(m.group(2))

    # Taxi times
    taxi_matches = re.findall(r"Taxi Time:\s*(\d+)\s*minutes", page_text)
    if len(taxi_matches) >= 1:
        data["taxi_out_min"] = int(taxi_matches[0])
    if len(taxi_matches) >= 2:
        data["taxi_in_min"] = int(taxi_matches[1])

    # Departure gate
    gate_m = re.search(r"(?:left\s+)?GATE\s+([A-Z0-9]+)", page_text)
    if gate_m:
        data["departure_gate"] = gate_m.group(1)

    return data


# ──────────────────────────────────────────────────────────────────────────────
# SCRAPER WORKER — Processes one flight at a time
# ──────────────────────────────────────────────────────────────────────────────

def scrape_one_flight(page, task: dict, config: Config) -> dict:
    """
    Scrape a single flight using a Playwright browser page.
    Uses the shared OFFSET_CACHE to build Zulu URLs directly when possible.
    See scrape_one_flight_selenium for the full strategy description.
    """
    orig_icao = task["orig_icao"]

    # Check cache → build Zulu URL directly if offset is known
    with OFFSET_CACHE_LOCK:
        cached_offset = OFFSET_CACHE.get(orig_icao)

    zulu_hhmm = local_to_zulu(task.get("sched_dep_local"), cached_offset) if cached_offset is not None else None

    if zulu_hhmm:
        url = (
            f"https://www.flightaware.com/live/flight/"
            f"{task['flight_fa']}/history/{task['date_str']}/"
            f"{zulu_hhmm}Z/{orig_icao}/{task['dest_icao']}"
        )
    else:
        url = task["url"]   # Plain URL — first flight for this airport

    result = {
        "flight_fa": task["flight_fa"],
        "mkt_al": task.get("mkt_al"),
        "op_al": task["op_al"],
        "orig": task["orig"],
        "dest": task["dest"],
        "orig_icao": orig_icao,
        "dest_icao": task["dest_icao"],
        "route": task["route"],
        "date": task["date_str"],
        "url": url,
        "zulu_time": f"{zulu_hhmm}Z" if zulu_hhmm else None,
        "utc_offset_used": cached_offset,
        "sched_block_mins": task.get("sched_block_mins"),
        "sched_dep_local": task.get("sched_dep_local"),
        "sched_arr_local": task.get("sched_arr_local"),
        "equip_sched": task.get("equip"),
        "seats_sched": task.get("seats"),
        "distance_km": task.get("distance_km"),
    }

    for attempt in range(config.max_retries + 1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=config.page_timeout_sec * 1000)

            # Wait past Cloudflare challenge if present
            if not wait_past_cloudflare_playwright(page, max_wait_sec=30):
                result["scrape_status"] = "ERROR:cloudflare_challenge"
                if attempt < config.max_retries:
                    time.sleep(5)
                    continue
                break

            page.wait_for_selector(
                "span.flightStatusBig, div.track-panel, .flightPageSummary, .flightPageNotAvailable",
                timeout=10000,
            )
            time.sleep(1.5)

            html = page.content()

            # ── Cache miss path: figure out the Zulu URL and redirect to it ──
            if cached_offset is None:
                # Re-check cache — another worker may have populated it
                with OFFSET_CACHE_LOCK:
                    cached_offset = OFFSET_CACHE.get(orig_icao)

                if cached_offset is None:
                    discovered = extract_utc_offset_from_html(html)
                    if discovered is not None:
                        with OFFSET_CACHE_LOCK:
                            OFFSET_CACHE[orig_icao] = discovered
                        cached_offset = discovered

                zulu_hhmm = local_to_zulu(task.get("sched_dep_local"), cached_offset) if cached_offset is not None else None

                corrected_url = None
                if zulu_hhmm:
                    corrected_url = (
                        f"https://www.flightaware.com/live/flight/"
                        f"{task['flight_fa']}/history/{task['date_str']}/"
                        f"{zulu_hhmm}Z/{orig_icao}/{task['dest_icao']}"
                    )
                else:
                    # Fallback: pick a Zulu link from the history picker
                    zulu_pattern = (
                        rf"/live/flight/{task['flight_fa']}/history/"
                        rf"{task['date_str']}/(\d{{4}}Z)/"
                        rf"{orig_icao}/{task['dest_icao']}"
                    )
                    zulu_times = re.findall(zulu_pattern, html)
                    chosen_zulu = _select_best_zulu(zulu_times, task)
                    if chosen_zulu:
                        corrected_url = (
                            f"https://www.flightaware.com/live/flight/"
                            f"{task['flight_fa']}/history/{task['date_str']}/"
                            f"{chosen_zulu}/{orig_icao}/{task['dest_icao']}"
                        )
                        result["zulu_time"] = chosen_zulu

                if corrected_url:
                    result["utc_offset_used"] = cached_offset
                    result["url_original"] = url
                    result["url"] = corrected_url
                    if zulu_hhmm:
                        result["zulu_time"] = f"{zulu_hhmm}Z"
                    page.goto(corrected_url, wait_until="domcontentloaded",
                              timeout=config.page_timeout_sec * 1000)
                    wait_past_cloudflare_playwright(page, max_wait_sec=30)
                    page.wait_for_selector(
                        "span.flightStatusBig, div.track-panel, .flightPageSummary",
                        timeout=10000,
                    )
                    time.sleep(1.5)
                    html = page.content()

            parsed = parse_flight_html(html)
            if parsed:
                result.update(parsed)
                result["scrape_status"] = "OK"
                break
            else:
                result["scrape_status"] = "NO_DATA"

        except Exception as e:
            if attempt < config.max_retries:
                time.sleep(2)
                continue
            result["scrape_status"] = f"ERROR:{str(e)[:80]}"

    return result


def scrape_one_flight_selenium(driver, task: dict, config: Config) -> dict:
    """
    Scrape one flight using Selenium.

    Strategy:
      1. If we already know the origin airport's UTC offset (cached from a
         previous flight), build the Zulu URL directly and scrape it. One hit.
      2. If not cached, hit the generic URL (no Zulu). Parse out the offset
         FlightAware displays (e.g. "-05" next to local times), cache it for
         future flights, then redirect to the Zulu URL and scrape again.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    orig_icao = task["orig_icao"]

    # ── Step 1: Check cache, build Zulu URL directly if we already know offset ──
    with OFFSET_CACHE_LOCK:
        cached_offset = OFFSET_CACHE.get(orig_icao)

    zulu_hhmm = local_to_zulu(task.get("sched_dep_local"), cached_offset) if cached_offset is not None else None

    if zulu_hhmm:
        url = (
            f"https://www.flightaware.com/live/flight/"
            f"{task['flight_fa']}/history/{task['date_str']}/"
            f"{zulu_hhmm}Z/{orig_icao}/{task['dest_icao']}"
        )
    else:
        url = task["url"]   # Plain URL (no Zulu) — first flight for this airport

    result = {
        "flight_fa": task["flight_fa"],
        "mkt_al": task.get("mkt_al"),
        "op_al": task["op_al"],
        "orig": task["orig"],
        "dest": task["dest"],
        "orig_icao": orig_icao,
        "dest_icao": task["dest_icao"],
        "route": task["route"],
        "date": task["date_str"],
        "url": url,
        "zulu_time": f"{zulu_hhmm}Z" if zulu_hhmm else None,
        "utc_offset_used": cached_offset,
        "sched_block_mins": task.get("sched_block_mins"),
        "sched_dep_local": task.get("sched_dep_local"),
        "sched_arr_local": task.get("sched_arr_local"),
        "equip_sched": task.get("equip"),
        "seats_sched": task.get("seats"),
        "distance_km": task.get("distance_km"),
    }

    for attempt in range(config.max_retries + 1):
        try:
            driver.get(url)

            # ── Wait past Cloudflare's challenge page if it appears ──
            # FlightAware is behind Cloudflare. When Cloudflare is suspicious
            # (e.g. 12 parallel requests from same IP), it serves an
            # interstitial "Just a moment..." page that runs JS for 5+ seconds
            # before the real page loads. Without this wait, the scraper
            # would time out looking for .flightStatusBig on the challenge page.
            if not wait_past_cloudflare_selenium(driver, max_wait_sec=30):
                # Still stuck on challenge after 30s → mark as error and retry
                result["scrape_status"] = "ERROR:cloudflare_challenge"
                if attempt < config.max_retries:
                    time.sleep(5)
                    continue
                break

            WebDriverWait(driver, config.page_timeout_sec).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    "span.flightStatusBig, div.track-panel, .flightPageSummary, .flightPageNotAvailable"))
            )
            time.sleep(1.5)

            html = driver.page_source

            # ── Cache miss path: figure out the Zulu URL and redirect to it ──
            if cached_offset is None:
                # First check: another worker might have populated the cache
                # for this airport while we were loading the generic page.
                with OFFSET_CACHE_LOCK:
                    cached_offset = OFFSET_CACHE.get(orig_icao)

                # Still no cache → try to read the offset from THIS page's HTML
                if cached_offset is None:
                    discovered = extract_utc_offset_from_html(html)
                    if discovered is not None:
                        with OFFSET_CACHE_LOCK:
                            OFFSET_CACHE[orig_icao] = discovered
                        cached_offset = discovered

                # Build Zulu URL from whatever offset we now have
                zulu_hhmm = local_to_zulu(task.get("sched_dep_local"), cached_offset) if cached_offset is not None else None

                corrected_url = None
                if zulu_hhmm:
                    corrected_url = (
                        f"https://www.flightaware.com/live/flight/"
                        f"{task['flight_fa']}/history/{task['date_str']}/"
                        f"{zulu_hhmm}Z/{orig_icao}/{task['dest_icao']}"
                    )
                else:
                    # ── FALLBACK: offset regex failed (thin page, no visible times).
                    # Scan the history picker for OTHER Zulu links for this same
                    # flight+date+route. Pick the one closest to our Cirium
                    # scheduled dep_time (interpreted as local). This is the
                    # rescue path for airports whose "Scheduled" block doesn't
                    # render a recognizable time+offset.
                    zulu_pattern = (
                        rf"/live/flight/{task['flight_fa']}/history/"
                        rf"{task['date_str']}/(\d{{4}}Z)/"
                        rf"{orig_icao}/{task['dest_icao']}"
                    )
                    zulu_times = re.findall(zulu_pattern, html)
                    chosen_zulu = _select_best_zulu(zulu_times, task)
                    if chosen_zulu:
                        corrected_url = (
                            f"https://www.flightaware.com/live/flight/"
                            f"{task['flight_fa']}/history/{task['date_str']}/"
                            f"{chosen_zulu}/{orig_icao}/{task['dest_icao']}"
                        )
                        result["zulu_time"] = chosen_zulu

                # Re-navigate if we got a better URL
                if corrected_url:
                    result["utc_offset_used"] = cached_offset
                    result["url_original"] = url
                    result["url"] = corrected_url
                    if zulu_hhmm:
                        result["zulu_time"] = f"{zulu_hhmm}Z"
                    driver.get(corrected_url)
                    # Same Cloudflare wait for the redirect
                    wait_past_cloudflare_selenium(driver, max_wait_sec=30)
                    WebDriverWait(driver, config.page_timeout_sec).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR,
                            "span.flightStatusBig, div.track-panel, .flightPageSummary"))
                    )
                    time.sleep(1.5)
                    html = driver.page_source

            parsed = parse_flight_html(html)
            if parsed:
                result.update(parsed)
                result["scrape_status"] = "OK"
                break
            else:
                result["scrape_status"] = "NO_DATA"
        except Exception as e:
            if attempt < config.max_retries:
                time.sleep(2)
                continue
            result["scrape_status"] = f"ERROR:{str(e)[:80]}"

    return result


# ──────────────────────────────────────────────────────────────────────────────
# WORKER THREADS — Each worker manages its own browser instance
# ──────────────────────────────────────────────────────────────────────────────

def worker_playwright(task_queue, config: Config, worker_id: int,
                      progress: ProgressTracker) -> list[dict]:
    """
    Worker thread for Playwright engine. Pulls tasks from shared queue
    (same pattern as worker_selenium — see that docstring for rationale).
    """
    from playwright.sync_api import sync_playwright
    import queue as _queue

    results = []
    pw = sync_playwright().start()
    browser = pw.chromium.launch(
        headless=config.headless,
        args=["--no-sandbox", "--disable-dev-shm-usage"],
    )
    context = browser.new_context(
        user_agent=config.user_agent,
        viewport={"width": 1920, "height": 1080},
    )
    page = context.new_page()

    while True:
        try:
            task = task_queue.get_nowait()
        except _queue.Empty:
            break

        result = scrape_one_flight(page, task, config)
        results.append(result)

        success = result["scrape_status"] == "OK"
        progress.record(success)

        time.sleep(random.uniform(config.min_delay_sec, config.max_delay_sec))

    page.close()
    context.close()
    browser.close()
    pw.stop()

    return results


def worker_selenium(task_queue, config: Config, worker_id: int,
                    progress: ProgressTracker, chromedriver_path: str) -> list[dict]:
    """
    Worker thread using Selenium. Pulls tasks from a SHARED QUEUE — not a
    pre-assigned chunk. This prevents the "last worker left holding the bag"
    problem where one slow worker drags out the final minutes while others
    sit idle.

    Flow:
      - Each worker spins up its own Chrome (staggered, +0.8s per worker).
      - Loops: pop a task from the queue → scrape it → record → repeat.
      - When the queue is empty, the worker shuts down its browser and exits.

    Resilience:
      - Driver rebuild on crash
      - Auto-refresh after 8 consecutive failures (rate-limit recovery)
    """
    import queue as _queue

    # Prefer undetected-chromedriver (bypasses Cloudflare bot detection).
    # Falls back to stock Selenium if uc is not installed.
    try:
        import undetected_chromedriver as uc
        _use_uc = True
    except ImportError:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
        _use_uc = False

    # Stagger startup so 4-6 Chromes don't fight for resources at the same ms
    time.sleep(worker_id * 1.5)

    def build_driver():
        """
        Factory that creates a fresh Chrome driver — used on startup and after crashes.

        Uses undetected-chromedriver (uc) when available. uc patches the ChromeDriver
        binary to remove the bot-detection signals that Cloudflare fingerprints:
          - navigator.webdriver property
          - Chrome automation extension markers
          - CDP Runtime.enable leak
        This is why standard headless Selenium gets blocked after ~10 requests while
        uc sustains scraping across hundreds of flights.
        """
        profile_dir = os.path.abspath(f"./chrome_profiles/worker_{worker_id}")
        os.makedirs(profile_dir, exist_ok=True)

        if _use_uc:
            options = uc.ChromeOptions()
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--log-level=3")
            options.add_argument("--disable-logging")
            options.add_argument(f"--user-data-dir={profile_dir}")
            d = uc.Chrome(
                options=options,
                headless=config.headless,
                version_main=None,      # auto-detect installed Chrome version
                use_subprocess=True,    # isolate each worker's Chrome process
            )
        else:
            # Stock Selenium fallback (may get blocked by Cloudflare)
            from selenium.webdriver.chrome.options import Options
            from selenium.webdriver.chrome.service import Service
            options = Options()
            if config.headless:
                options.add_argument("--headless")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--log-level=3")
            options.add_argument("--disable-logging")
            options.add_argument(f"--user-data-dir={profile_dir}")
            options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
            service = Service(chromedriver_path) if chromedriver_path else Service()
            service.log_output = subprocess.DEVNULL
            d = webdriver.Chrome(service=service, options=options)
            d.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            })

        d.set_page_load_timeout(config.page_timeout_sec)
        return d

    driver = build_driver()
    results = []
    consecutive_errors = 0

    # ── Main pull-loop: keep taking tasks until the queue is empty ──
    while True:
        try:
            task = task_queue.get_nowait()
        except _queue.Empty:
            break   # Queue drained — we're done

        try:
            result = scrape_one_flight_selenium(driver, task, config)
        except Exception as e:
            result = {
                "flight_fa": task["flight_fa"],
                "op_al": task["op_al"],
                "orig": task["orig"],
                "dest": task["dest"],
                "orig_icao": task["orig_icao"],
                "dest_icao": task["dest_icao"],
                "route": task["route"],
                "date": task["date_str"],
                "url": task["url"],
                "scrape_status": f"ERROR:driver_crash:{str(e)[:60]}",
            }
            try:
                driver.quit()
            except Exception:
                pass
            try:
                time.sleep(3)
                driver = build_driver()
            except Exception as build_err:
                log.error(f"OTP Agent {worker_id}: can't rebuild driver: {build_err}")
                # Put the current task back so another worker can try it
                task_queue.put(task)
                break

        results.append(result)
        success = result["scrape_status"] == "OK"
        progress.record(success)

        if success:
            consecutive_errors = 0
        else:
            consecutive_errors += 1
            if consecutive_errors >= 8:
                log.warning(f"OTP Agent {worker_id}: 8 consecutive failures, refreshing driver")
                try:
                    driver.quit()
                except Exception:
                    pass
                time.sleep(5)
                try:
                    driver = build_driver()
                    consecutive_errors = 0
                except Exception:
                    log.error(f"OTP Agent {worker_id}: driver refresh failed, stopping")
                    break

        time.sleep(random.uniform(config.min_delay_sec, config.max_delay_sec))

    try:
        driver.quit()
    except Exception:
        pass
    return results


# ──────────────────────────────────────────────────────────────────────────────
# PARALLEL ORCHESTRATOR
# ──────────────────────────────────────────────────────────────────────────────

def run_parallel_scrape(tasks: list[dict], config: Config) -> list[dict]:
    """
    Distribute scraping tasks across worker threads using a SHARED QUEUE.

    Why a shared queue instead of pre-assigned chunks:
    ---------------------------------------------------
    The old approach split tasks into 6 chunks up-front (one per worker).
    Problem: if worker A got assigned lots of slow flights (UIO, rare
    airports, Cloudflare-challenged routes) and worker B got lucky with
    fast flights (BOG domestic), worker B finishes its chunk in 10 minutes
    while worker A is still grinding for another 7 minutes. Net effect:
    run takes max(all_workers), not average. That's why every previous run
    "stuck" at 94%+ for 3-5 minutes.

    With a shared queue, workers take one task at a time and move on. The
    fast ones keep pulling more work, the slow ones take their time, but
    NO ONE SITS IDLE while others are still working. Finish time collapses
    to average(workers), not max(workers).

    Threads + queue works well here because scraping is I/O-bound; each
    worker spends most of its time waiting for pages to load.
    """
    import queue

    n_workers = min(config.max_workers, len(tasks))
    status_file = os.path.join(config.output_dir, "scraper_status.json")
    progress = ProgressTracker(
        total=len(tasks),
        update_every=max(10, len(tasks) // 20),
        status_file=status_file,
    )

    # ── Build the shared task queue ──
    # All tasks go in; workers race to pop them. queue.Queue is thread-safe.
    task_queue = queue.Queue()
    for t in tasks:
        task_queue.put(t)

    # ── Detect browser engine ──
    use_playwright = False
    try:
        from playwright.sync_api import sync_playwright
        pw_test = sync_playwright().start()
        pw_test.stop()
        use_playwright = True
        log.info(f"Engine: Playwright │ OTP Agents: {n_workers} │ Tasks: {len(tasks)}")
    except Exception:
        try:
            import selenium
            import importlib.util
            if importlib.util.find_spec("undetected_chromedriver") is not None:
                log.info(f"Engine: Selenium+UC (Cloudflare bypass) │ OTP Agents: {n_workers} │ Tasks: {len(tasks)}")
            else:
                log.info(f"Engine: Selenium (standard) │ OTP Agents: {n_workers} │ Tasks: {len(tasks)}")
                log.warning("undetected-chromedriver not found — Cloudflare may block requests.")
                log.warning("Install with: pip install undetected-chromedriver")
        except ImportError:
            log.error("Neither Playwright nor Selenium is available.")
            log.error("Install one: pip install selenium webdriver-manager")
            return []

    worker_fn = worker_playwright if use_playwright else worker_selenium

    # ── Chromedriver setup (Selenium only) ──
    # When undetected-chromedriver is available it manages its own chromedriver
    # binary internally — no webdriver-manager needed. Fall back to webdriver-manager
    # only if uc is not installed.
    import importlib.util as _importlib_util
    chromedriver_path = None
    if not use_playwright:
        if _importlib_util.find_spec("undetected_chromedriver") is not None:
            log.info("undetected-chromedriver will manage chromedriver automatically.")
        else:
            import logging as _logging
            _logging.getLogger("WDM").setLevel(_logging.CRITICAL)
            os.environ["WDM_LOG"] = "0"
            os.environ["WDM_LOG_LEVEL"] = "0"
            try:
                from webdriver_manager.chrome import ChromeDriverManager
                log.info("Preparing chromedriver (one-time download if needed)...")
                chromedriver_path = ChromeDriverManager().install()
                if chromedriver_path and not chromedriver_path.endswith((".exe", "chromedriver")):
                    import glob
                    candidate_dir = os.path.dirname(chromedriver_path)
                    exe_matches = glob.glob(os.path.join(candidate_dir, "chromedriver.exe"))
                    if exe_matches:
                        chromedriver_path = exe_matches[0]
                try:
                    os.chmod(chromedriver_path, 0o755)
                except Exception:
                    pass
                log.info(f"Chromedriver ready: {chromedriver_path}")
            except Exception as e:
                log.warning(f"webdriver-manager failed: {e}. Will try system chromedriver.")
                chromedriver_path = None

    # ── Launch workers ──
    # Each worker gets the same shared queue + its own worker_id.
    all_results = []

    with ThreadPoolExecutor(max_workers=n_workers) as executor:
        if use_playwright:
            futures = {
                executor.submit(worker_fn, task_queue, config, wid, progress): wid
                for wid in range(n_workers)
            }
        else:
            futures = {
                executor.submit(worker_fn, task_queue, config, wid, progress, chromedriver_path): wid
                for wid in range(n_workers)
            }

        for future in as_completed(futures):
            wid = futures[future]
            try:
                results = future.result()
                all_results.extend(results)
            except Exception as e:
                log.error(f"OTP Agent {wid} crashed: {e}")

    # ── Final summary ──
    s = progress.summary()
    elapsed_sec = s["elapsed_sec"]
    avg_per_flight = elapsed_sec / max(s["done"], 1)

    log.info(
        f"Scraping done │ {s['ok']} OK, {s['fail']} failed, "
        f"{elapsed_sec}s total ({avg_per_flight:.1f}s/flight)"
    )

    # Build run stats — returned alongside results so the caller can write
    # them to the Excel report (a "Run_Info" sheet) and the console summary.
    run_stats = {
        "otp_agents": n_workers,
        "engine": "Playwright" if use_playwright else "Selenium",
        "total_tasks": len(tasks),
        "ok_count": s["ok"],
        "fail_count": s["fail"],
        "elapsed_sec": elapsed_sec,
        "elapsed_human": _format_duration(elapsed_sec),
        "avg_sec_per_flight": round(avg_per_flight, 2),
    }

    return all_results, run_stats


def _format_duration(seconds: float) -> str:
    """Format a duration in seconds as 'Xm Ys' or 'Xh Ym Zs'."""
    seconds = int(seconds)
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    if h > 0:
        return f"{h}h {m}m {s}s"
    if m > 0:
        return f"{m}m {s}s"
    return f"{s}s"


# ──────────────────────────────────────────────────────────────────────────────
# FR24 API DATA SOURCE
# ──────────────────────────────────────────────────────────────────────────────
# Alternative to browser-based scraping. Uses the FlightRadar24 Business API
# (fr24api.flightradar24.com) to fetch actual departure/arrival times for an
# entire airline's daily schedule in a few API calls, instead of one browser
# session per flight.
#
# Data fields used from FR24 response:
#   datetime_takeoff   — wheels-off UTC (ISO 8601, e.g. "2026-05-16T05:20:00Z")
#   datetime_landed    — wheels-on  UTC
#   dest_icao_actual   — actual destination (differs from dest_icao if diverted)
#   flight_ended       — bool: False if flight still in progress
#   flight             — IATA flight number string e.g. "EY101"
#   type               — aircraft type e.g. "A350"
#   reg                — registration e.g. "A6-XWA"
#
# Rate limit: ~10 pages/burst; this code sleeps 0.5s between pages and uses
# exponential back-off on 429 errors.

FR24_API_BASE     = "https://fr24api.flightradar24.com/api"
FR24_PAGE_SIZE    = 100   # requested page size (API caps at 20 regardless)
FR24_SLEEP_SEC    = 1.2   # pause between paginated requests (keeps burst under rate limit)
FR24_RATELIMIT_WAIT = 65  # seconds to wait after a 429 before retrying
FR24_MAX_PAGES    = 30    # hard cap — prevents runaway fetches


def _fr24_get(url: str, api_key: str, timeout: int = 30) -> dict:
    """
    HTTP GET helper for the FR24 API. Tries `requests` first (handles SSL/headers
    better), falls back to `urllib.request`. Returns the parsed JSON dict.
    """
    import json as _json

    headers = {
        "Accept":         "application/json",
        "Accept-Version": "v1",
        "Authorization":  f"Bearer {api_key}",
        "User-Agent":     "Mozilla/5.0 (compatible; OTP-Scraper/1.0)",
    }
    try:
        import requests as _requests
        resp = _requests.get(url, headers=headers, timeout=timeout)
        resp.raise_for_status()
        return resp.json()
    except ImportError:
        pass
    # urllib fallback
    import urllib.request as _urllib_req
    req = _urllib_req.Request(url, headers=headers)
    with _urllib_req.urlopen(req, timeout=timeout) as resp:
        return _json.loads(resp.read())


FR24_BATCH_SIZE = 10   # IATA flight numbers per API call (>10 returns 400)


def fetch_fr24_flights_by_numbers(iata_numbers: list, date_str: str, api_key: str) -> list:
    """
    Fetch FR24 flight-summary data for a specific list of IATA flight numbers
    (e.g. ["EY247", "EY14", "EY42"]) on a given date.

    WHY THIS APPROACH:
    The `operating_as` filter has a broken pagination bug — every page returns
    the same first 20 records. Querying by explicit flight numbers sidesteps
    pagination entirely and gets precise per-flight results.

    TIME WINDOW:
    Uses a 24-hour UTC window (date T00:00Z → date T23:59Z). This matches how
    Cirium assigns the "Travel Date" field — by the UTC date of the actual
    departure, not the local calendar date at the origin city. Verified correct
    for EY14 ATL→AUH (departs 21:40 Atlanta local = 01:53 UTC next calendar day,
    and Cirium Travel Date = the UTC date = 2026-05-16).

    BATCHING:
    FR24_BATCH_SIZE flights per API call. For 249 EY flights: ~5 calls.
    With FR24_SLEEP_SEC between calls: fetch completes in ~6 seconds, well
    under the 10-request burst quota that triggers rate limiting.

    Args:
        iata_numbers: list of IATA flight strings e.g. ["EY247", "EY728"]
        date_str:     target date as "YYYY-MM-DD"
        api_key:      FR24 Business API key

    Returns:
        List of raw FR24 flight dicts with datetime_takeoff, datetime_landed, etc.
    """
    dt      = datetime.strptime(date_str, "%Y-%m-%d")
    dt_from = dt.strftime("%Y-%m-%dT00:00:00Z")
    dt_to   = (dt + timedelta(hours=27)).strftime("%Y-%m-%dT%H:%M:%SZ")  # +3h buffer for long-haul

    all_records    = []
    ratelimit_hits = 0
    n_batches      = (len(iata_numbers) + FR24_BATCH_SIZE - 1) // FR24_BATCH_SIZE

    for batch_idx in range(n_batches):
        batch      = iata_numbers[batch_idx * FR24_BATCH_SIZE:(batch_idx + 1) * FR24_BATCH_SIZE]
        flights_qs = ",".join(batch)
        url = (
            f"{FR24_API_BASE}/flight-summary/full"
            f"?flights={flights_qs}"
            f"&flight_datetime_from={dt_from}"
            f"&flight_datetime_to={dt_to}"
        )
        try:
            data    = _fr24_get(url, api_key)
            records = data.get("data", [])
            all_records.extend(records)
            log.info(f"FR24 batch {batch_idx+1}/{n_batches}: "
                     f"{len(records)}/{len(batch)} flights matched")
            ratelimit_hits = 0
            time.sleep(FR24_SLEEP_SEC)
        except Exception as _e:
            http_code = getattr(_e, "code", None) or getattr(
                getattr(_e, "response", None), "status_code", None
            )
            if http_code == 429:
                ratelimit_hits += 1
                if ratelimit_hits > 2:
                    log.warning(f"FR24 rate limit persists — "
                                f"{len(all_records)} records collected so far")
                    break
                log.warning(f"FR24 rate-limited on batch {batch_idx+1}, "
                            f"waiting {FR24_RATELIMIT_WAIT}s...")
                time.sleep(FR24_RATELIMIT_WAIT)
                # Retry this batch once
                try:
                    data    = _fr24_get(url, api_key)
                    records = data.get("data", [])
                    all_records.extend(records)
                    ratelimit_hits = 0
                except Exception as _e2:
                    log.error(f"FR24 batch {batch_idx+1} retry failed: {_e2}")
            else:
                log.error(f"FR24 batch {batch_idx+1} error: {_e}")

    log.info(f"FR24: {len(all_records)} records fetched for {len(iata_numbers)} flights on {date_str}")
    return all_records


def _iso_to_epoch(iso_str: Optional[str]) -> Optional[int]:
    """Parse a UTC ISO-8601 string ('2026-05-16T05:20:00Z') to Unix epoch (int)."""
    if not iso_str:
        return None
    try:
        import calendar as _cal
        s = iso_str.rstrip("Z").replace("T", " ")
        dt = datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
        return int(_cal.timegm(dt.timetuple()))
    except Exception:
        return None


def _local_hhmm_to_epoch(date_str: str, hhmm, utc_offset_hours: int) -> Optional[int]:
    """
    Convert a Cirium local-time HHMM integer + UTC offset to a Unix epoch.

    date_str: "YYYYMMDD" — the operating date in the origin's local calendar
    hhmm:     int/float e.g. 2130  →  21:30 local
    utc_offset_hours: integer e.g. +4 for AUH
    """
    try:
        import calendar as _cal
        t = int(float(hhmm))
        hh, mm = t // 100, t % 100
        if not (0 <= hh <= 23 and 0 <= mm <= 59):
            return None
        dt_local = datetime.strptime(date_str, "%Y%m%d").replace(hour=hh, minute=mm)
        dt_utc   = dt_local - timedelta(hours=utc_offset_hours)
        return int(_cal.timegm(dt_utc.timetuple()))
    except Exception:
        return None


def run_fr24_otp(tasks: list, config: Config, api_key: str,
                 coord_offsets: dict) -> tuple:
    """
    FR24 API alternative to run_parallel_scrape().

    1. Groups tasks by (airline_icao, date) — usually one group for EY.
    2. Fetches all flights for that group in a paginated batch call.
    3. Matches each Cirium task to an FR24 record by flight number + route.
    4. Converts FR24 timestamps to the epoch fields that compute_metrics() expects.

    Returns (results: list[dict], run_stats: dict) — same contract as
    run_parallel_scrape(), so the rest of main() is unchanged.

    Delay computation accuracy:
    - Hub departures (AUH/DXB/DOH): exact — offsets hardcoded in HUB_UTC_OFFSETS.
    - Other airports: ±1 hour — longitude-based estimate from airports.csv.
      For 15-minute OTP thresholds, errors >15 min are possible at non-integer-hour
      timezone airports (India UTC+5:30, Iran UTC+3:30, etc.). This is documented in
      the Run_Info sheet of the Excel output.
    """
    import json as _json

    start_time = time.time()
    os.makedirs(config.output_dir, exist_ok=True)
    status_file = os.path.join(config.output_dir, "scraper_status.json")

    # Merge lookup priority: OFFSET_CACHE (scraped, exact) > HUB_UTC_OFFSETS (hardcoded)
    # > coord_offsets (longitude-based, ±1h)
    all_offsets = {**coord_offsets, **HUB_UTC_OFFSETS, **OFFSET_CACHE}

    # ── Step 1: Group tasks by date, collect IATA flight numbers, batch-fetch ──
    tasks_by_date = {}
    for task in tasks:
        date_str = task.get("date_str", "")
        date_iso = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
        tasks_by_date.setdefault(date_iso, []).append(task)

    fr24_raw = {}   # {date_iso: [records]}
    for date_iso, date_tasks in sorted(tasks_by_date.items()):
        iata_numbers = []
        seen_flights: set = set()
        for task in date_tasks:
            airline_iata    = task.get("mkt_al", "")
            flight_fa       = task.get("flight_fa", "")
            flight_num      = re.sub(r"^[A-Z]{2,3}", "", flight_fa)
            fr24_flight_str = f"{airline_iata}{flight_num}"
            if fr24_flight_str and fr24_flight_str not in seen_flights:
                seen_flights.add(fr24_flight_str)
                iata_numbers.append(fr24_flight_str)
        log.info(f"FR24: fetching ETD flights for {date_iso}...")
        fr24_raw[date_iso] = fetch_fr24_flights_by_numbers(iata_numbers, date_iso, api_key)

    # ── Step 2: Build lookup {fr24_flight_str: [records]} ──
    # IATA flight string (e.g. "EY101") already encodes the airline, no need for airline_icao key
    fr24_lookup: dict = {}
    for _, records in fr24_raw.items():
        for rec in records:
            flight_str = rec.get("flight", "")
            fr24_lookup.setdefault(flight_str, []).append(rec)

    # ── Step 3: Match tasks → FR24 records ──
    results   = []
    ok_count  = 0
    fail_count = 0
    total     = len(tasks)

    for i, task in enumerate(tasks, 1):
        airline_iata = task.get("mkt_al", "")
        date_str     = task.get("date_str", "")
        flight_fa    = task.get("flight_fa", "")     # e.g. "ETD101"
        orig_icao    = task.get("orig_icao", "")
        dest_icao    = task.get("dest_icao", "")

        # Strip ICAO prefix from flight_fa to get numeric part → IATA flight string
        flight_num      = re.sub(r"^[A-Z]{2,3}", "", flight_fa)   # "101"
        fr24_flight_str = f"{airline_iata}{flight_num}"            # "EY101"

        row = {
            "flight_fa":        flight_fa,
            "mkt_al":           task.get("mkt_al"),
            "op_al":            task.get("op_al"),
            "orig":             task.get("orig"),
            "dest":             task.get("dest"),
            "route":            task.get("route"),
            "orig_icao":        orig_icao,
            "dest_icao":        dest_icao,
            "date_str":         date_str,
            "date":             date_str,
            "sched_dep_local":  task.get("sched_dep_local"),
            "sched_arr_local":  task.get("sched_arr_local"),
            "sched_block_mins": task.get("sched_block_mins"),
            "equip_sched":      task.get("equip"),
            "seats_sched":      task.get("seats"),
            "distance_km":      task.get("distance_km"),
            "data_source":      "FR24",
        }

        # Find best FR24 match: same flight_str → prefer exact orig+dest
        candidates = fr24_lookup.get(fr24_flight_str, [])
        match = None
        for c in candidates:
            if c.get("orig_icao") == orig_icao and c.get("dest_icao") == dest_icao:
                match = c
                break
        # Fallback: if multiple dates in window, pick closest to sched dep UTC
        if match is None and candidates:
            utc_off = all_offsets.get(orig_icao)
            if utc_off is not None and task.get("sched_dep_local") is not None:
                sched_epoch = _local_hhmm_to_epoch(date_str, task["sched_dep_local"], utc_off)
                if sched_epoch:
                    candidates.sort(
                        key=lambda c: abs((_iso_to_epoch(c.get("datetime_takeoff")) or 0) - sched_epoch)
                    )
            match = candidates[0]

        if match is None:
            row["scrape_status"] = "NOT_FOUND"
            row["flight_status"] = "CANCELLED"
            fail_count += 1
        else:
            dep_actual_epoch = _iso_to_epoch(match.get("datetime_takeoff"))
            arr_actual_epoch = _iso_to_epoch(match.get("datetime_landed"))

            row["dep_gate_actual_epoch"] = dep_actual_epoch
            row["arr_gate_actual_epoch"] = arr_actual_epoch
            row["fr24_id"]               = match.get("fr24_id")
            row["aircraft_type_fa"]      = match.get("type")
            row["tail_number"]           = match.get("reg")

            utc_off_orig = all_offsets.get(orig_icao)
            utc_off_dest = all_offsets.get(dest_icao)

            if utc_off_orig is not None and task.get("sched_dep_local") is not None:
                row["dep_gate_sched_epoch"] = _local_hhmm_to_epoch(
                    date_str, task["sched_dep_local"], utc_off_orig
                )
            if utc_off_dest is not None and task.get("sched_arr_local") is not None:
                row["arr_gate_sched_epoch"] = _local_hhmm_to_epoch(
                    date_str, task["sched_arr_local"], utc_off_dest
                )
                # Overnight correction: if sched arrival < sched departure, add one day
                dep_s = row.get("dep_gate_sched_epoch")
                arr_s = row.get("arr_gate_sched_epoch")
                if dep_s and arr_s and arr_s < dep_s:
                    row["arr_gate_sched_epoch"] = arr_s + 86400

            # Diversion detection
            dest_actual = match.get("dest_icao_actual")
            if dest_actual and dest_actual != match.get("dest_icao"):
                row["flight_status"] = "DIVERTED"
            elif not match.get("flight_ended", True):
                row["flight_status"] = "IN_FLIGHT"
            else:
                row["flight_status"] = "LANDED"

            row["scrape_status"] = "OK"
            ok_count += 1

        results.append(row)

        # Write live status JSON so otp_viewer.py can monitor FR24 runs too
        if i % 10 == 0 or i == total:
            elapsed = time.time() - start_time
            try:
                with open(status_file, "w") as _sf:
                    _json.dump({
                        "status":       "running",
                        "total":        total,
                        "done":         i,
                        "ok":           ok_count,
                        "fail":         fail_count,
                        "pct":          round(i / total * 100, 1),
                        "speed_per_min": round((i / max(elapsed, 1)) * 60, 1),
                        "eta_min":      0.0,
                        "elapsed_sec":  round(elapsed, 1),
                        "last_updated": datetime.now().isoformat(),
                    }, _sf)
            except Exception:
                pass

    elapsed = time.time() - start_time
    run_stats = {
        "otp_agents":         1,
        "engine":             "FR24 API",
        "total_tasks":        total,
        "ok_count":           ok_count,
        "fail_count":         fail_count,
        "elapsed_sec":        round(elapsed, 1),
        "elapsed_human":      _format_duration(elapsed),
        "avg_sec_per_flight": round(elapsed / max(total, 1), 2),
    }

    log.info(
        f"FR24 OTP complete: {ok_count} matched / {fail_count} not found "
        f"in {_format_duration(elapsed)}"
    )
    return results, run_stats


# ──────────────────────────────────────────────────────────────────────────────
# METRICS ENGINE — Calculate OTP metrics per flight
# ──────────────────────────────────────────────────────────────────────────────

def epoch_diff_minutes(sched_epoch: str, actual_epoch: str) -> Optional[float]:
    """
    Calculate delay in minutes from two epoch timestamps: actual - scheduled.
    Positive = late, negative = early.
    Returns None if either timestamp is missing or invalid.
    """
    try:
        t_sched = int(str(sched_epoch).strip())
        t_actual = int(str(actual_epoch).strip())
        return (t_actual - t_sched) / 60.0
    except (ValueError, TypeError, AttributeError):
        return None


def time_text_to_minutes(text: str) -> Optional[float]:
    """
    Convert a human-readable time like '09:24AM' to minutes since midnight.
    Used as fallback when epoch timestamps are unavailable.
    Ignores timezone suffixes (we only care about local time for OTP).
    """
    if not text or not isinstance(text, str):
        return None
    m = re.search(r"(\d{1,2}):(\d{2})\s*(AM|PM)", text.upper())
    if not m:
        return None
    h, mn, ap = int(m.group(1)), int(m.group(2)), m.group(3)
    if ap == "PM" and h != 12:
        h += 12
    elif ap == "AM" and h == 12:
        h = 0
    return h * 60 + mn


def _delay_from_text(sched_key: str, actual_key: str, row: dict) -> Optional[float]:
    """
    Calculate delay using text-based times (fallback when epoch not available).
    Handles midnight crossing with a ±720 minute threshold.
    """
    s = time_text_to_minutes(row.get(sched_key))
    a = time_text_to_minutes(row.get(actual_key))
    if s is None or a is None:
        return None
    diff = a - s
    if diff < -720:
        diff += 1440   # Crossed midnight forward
    elif diff > 720:
        diff -= 1440   # Crossed midnight backward
    return diff


def compute_metrics(row: dict) -> dict:
    """
    Calculate all OTP metrics for a single flight.

    METRIC DEFINITIONS:
    ───────────────────
    DEPARTURE:
      dep_delay_min  = actual gate departure − scheduled gate departure (minutes)
      D0   = 1 if dep_delay ≤ 0   (departed on time or early)
      D5   = 1 if dep_delay < 5   (departed within 5 minutes — boundary EXCLUDED)
      D15  = 1 if dep_delay < 15  (departed within 15 minutes — boundary EXCLUDED)
      D30  = 1 if dep_delay < 30
      D60  = 1 if dep_delay < 60

    ARRIVAL:
      arr_delay_min  = actual gate arrival − scheduled gate arrival (minutes)
      A0   = 1 if arr_delay ≤ 0   (arrived on time or early)
      A14  = 1 if arr_delay < 14  (DOT standard — boundary EXCLUDED)
      A15  = 1 if arr_delay < 15
      A30  = 1 if arr_delay < 30
      A60  = 1 if arr_delay < 60

    BLOCK TIME:
      actual_block_min   = actual gate-to-gate time
      sched_block_min    = scheduled gate-to-gate time
      within_block       = 1 if actual ≤ scheduled
      block_diff_min     = actual − scheduled

    COMPLETION:
      completed  = 1 if not cancelled
      cancelled  = 1 if status is CANCELLED
      diverted   = 1 if status is DIVERTED

    Data priority: epoch timestamps (from JS) > text times (from HTML) > page text delay
    """
    mx = {}

    # ── DEPARTURE DELAY ──
    # Priority 1: Epoch timestamps (timezone-safe)
    dep_delay = epoch_diff_minutes(
        row.get("dep_gate_sched_epoch"),
        row.get("dep_gate_actual_epoch"),
    )
    # Priority 2: Text-based times
    if dep_delay is None:
        dep_delay = _delay_from_text("dep_gate_sched_text", "dep_gate_actual_text", row)

    mx["dep_delay_min"] = round(dep_delay, 1) if dep_delay is not None else None

    if dep_delay is not None:
        mx["D0"]   = 1 if dep_delay <= 0 else 0     # On time or early (≤)
        mx["D5"]   = 1 if dep_delay < 5 else 0      # Strictly less than (<)
        mx["D15"]  = 1 if dep_delay < 15 else 0
        mx["D30"]  = 1 if dep_delay < 30 else 0
        mx["D60"]  = 1 if dep_delay < 60 else 0
        mx["D360"] = 1 if dep_delay < 360 else 0    # < 6 hours (major disruption threshold)

    # ── ARRIVAL DELAY ──
    # Priority 1: Gate arrival epochs
    arr_delay = epoch_diff_minutes(
        row.get("arr_gate_sched_epoch"),
        row.get("arr_gate_actual_epoch"),
    )
    # Priority 2: Landing epochs (if gate arrival not available)
    if arr_delay is None:
        arr_delay = epoch_diff_minutes(
            row.get("arr_landing_sched_epoch"),
            row.get("arr_landing_actual_epoch"),
        )
    # Priority 3: Gate arrival text
    if arr_delay is None:
        arr_delay = _delay_from_text("arr_gate_sched_text", "arr_gate_actual_text", row)
    # Priority 4: Landing text
    if arr_delay is None:
        arr_delay = _delay_from_text("arr_landing_sched_text", "arr_landing_actual_text", row)
    # Priority 5: Delay text from the page hero section
    if arr_delay is None:
        arr_delay = row.get("arrival_delay_text_min")

    mx["arr_delay_min"] = round(arr_delay, 1) if arr_delay is not None else None

    if arr_delay is not None:
        mx["A0"]   = 1 if arr_delay <= 0 else 0
        mx["A14"]  = 1 if arr_delay < 14 else 0     # DOT standard
        mx["A15"]  = 1 if arr_delay < 15 else 0
        mx["A30"]  = 1 if arr_delay < 30 else 0
        mx["A60"]  = 1 if arr_delay < 60 else 0
        mx["A360"] = 1 if arr_delay < 360 else 0    # < 6 hours

    # ── BLOCK TIME ──
    # Priority 1: Actual block from epochs (gate-to-gate)
    actual_block = epoch_diff_minutes(
        row.get("dep_gate_actual_epoch"),
        row.get("arr_gate_actual_epoch"),
    )
    # Priority 2: Actual block from text times
    if actual_block is None:
        dep_a = time_text_to_minutes(row.get("dep_gate_actual_text"))
        arr_a = time_text_to_minutes(row.get("arr_gate_actual_text"))
        if dep_a is not None and arr_a is not None:
            actual_block = arr_a - dep_a
            if actual_block < 0:
                actual_block += 1440  # Crossed midnight
    # Priority 3: Total travel time from page (e.g. "4h 24m total travel time")
    if actual_block is None:
        actual_block = row.get("actual_travel_min")

    # Scheduled block — three sources in order of reliability
    # Priority 1: Cirium Block_Mins (most reliable, part of the schedule row)
    sched_block = row.get("sched_block_mins")
    try:
        sched_block = float(sched_block) if sched_block is not None else None
    except (ValueError, TypeError):
        sched_block = None
    # Priority 2: FlightAware scheduled epochs
    if sched_block is None:
        sched_block = epoch_diff_minutes(
            row.get("dep_gate_sched_epoch"),
            row.get("arr_gate_sched_epoch"),
        )
    # Priority 3: FlightAware scheduled text times
    if sched_block is None:
        dep_s = time_text_to_minutes(row.get("dep_gate_sched_text"))
        arr_s = time_text_to_minutes(row.get("arr_gate_sched_text"))
        if dep_s is not None and arr_s is not None:
            sched_block = arr_s - dep_s
            if sched_block < 0:
                sched_block += 1440

    mx["actual_block_min"] = round(actual_block, 1) if actual_block is not None else None
    mx["sched_block_min"] = round(sched_block, 1) if sched_block is not None else None

    if actual_block is not None and sched_block is not None and sched_block > 0:
        mx["block_diff_min"] = round(actual_block - sched_block, 1)
        mx["within_block"] = 1 if actual_block <= sched_block else 0
        mx["block_pct"] = round(actual_block / sched_block * 100, 1)

    # ── COMPLETION ──
    status = str(row.get("flight_status", "")).upper()
    mx["completed"] = 0 if "CANCEL" in status else 1
    mx["cancelled"] = 1 if "CANCEL" in status else 0
    mx["diverted"]  = 1 if "DIVERT" in status else 0

    return mx


# ──────────────────────────────────────────────────────────────────────────────
# SUMMARY GENERATOR — Aggregate metrics by group
# ──────────────────────────────────────────────────────────────────────────────

def generate_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    Produce aggregated OTP summary statistics in an analysis-friendly shape.

    OUTPUT STRUCTURE:
    -----------------
    Each row represents one aggregation level. Instead of a single 'scope'
    column mixing different types ("OVERALL", "AIRLINE:EY", "ROUTE:AUH-LHR"),
    we split into TWO columns so pivot tables and filters work cleanly:

      scope_type    scope_value    Date          flights  D0_%  D15_%  ...
      OVERALL       ALL            ALL_DATES     850      84.3  94.9
      DATE          2026-04-16     2026-04-16    850      84.3  94.9
      AIRLINE       EY             ALL_DATES     280      84.2  95.0
      ROUTE         AUH-LHR        ALL_DATES     14       91.7  95.8
      FLIGHT        ETD101         ALL_DATES     1        100   100

    That way you can filter by scope_type='AIRLINE' in Excel/Power BI
    without regex, and group by Date for trend analysis.

    Groups the data by:
      - OVERALL: all flights combined
      - DATE: one row per target date
      - AIRLINE: one row per operating airline
      - ROUTE: one row per origin-destination pair
      - FLIGHT: one row per flight number

    For each group, calculates:
      - Percentage for each binary metric (D0%, D15%, A0%, A15%, etc.)
        Formula: SUM(metric) / COUNT(non-null values for that metric) × 100
        Cancelled flights with no delay data are excluded from delay percentages
        but included in completion rate.
      - Average, median, and P90 for delay minutes
    """
    metric_cols = [
        "D0", "D5", "D15", "D30", "D60", "D360",
        "A0", "A14", "A15", "A30", "A60", "A360",
        "within_block", "completed", "cancelled", "diverted",
    ]

    def summarize(group_df, scope_type, scope_value, date_value):
        """
        Build one summary row.
          scope_type  = 'OVERALL' | 'DATE' | 'AIRLINE' | 'ROUTE' | 'FLIGHT'
          scope_value = the actual value (e.g. 'EY', 'AUH-LHR', 'ETD101')
          date_value  = specific YYYY-MM-DD, or 'ALL_DATES' when aggregating across dates
        """
        n = len(group_df)
        row = {
            "scope_type": scope_type,
            "scope_value": scope_value,
            "Date": date_value,
            "flights": n,
        }

        for col in metric_cols:
            if col in group_df.columns:
                valid = group_df[col].dropna()
                if len(valid) > 0:
                    row[f"{col}_%"] = round(valid.sum() / len(valid) * 100, 1)

        for col in ["dep_delay_min", "arr_delay_min"]:
            if col in group_df.columns:
                valid = group_df[col].dropna()
                if len(valid) > 0:
                    row[f"avg_{col}"] = round(valid.mean(), 1)
                    row[f"med_{col}"] = round(valid.median(), 1)
                    row[f"p90_{col}"] = round(valid.quantile(0.9), 1)

        return row

    # Helper: format YYYYMMDD -> YYYY-MM-DD for readability in output
    def fmt_date(d):
        s = str(d)
        if len(s) == 8 and s.isdigit():
            return f"{s[:4]}-{s[4:6]}-{s[6:]}"
        return s

    rows = []

    # ── 1. OVERALL: one row combining everything ──
    rows.append(summarize(df, "OVERALL", "ALL", "ALL_DATES"))

    # ── 2. By date: one row per target date ──
    # This gives day-level trend visibility (e.g. Monday worse than Tuesday).
    if "date" in df.columns:
        for date, g in sorted(df.groupby("date")):
            rows.append(summarize(g, "DATE", fmt_date(date), fmt_date(date)))

    # ── 3. By MARKETING airline (aggregated across all dates) ──
    # For EY/EK/QR, mkt_al = op_al (no subsidiary operators), so this gives
    # a clean per-carrier breakdown: Etihad vs Emirates vs Qatar.
    if "mkt_al" in df.columns:
        for al, g in sorted(df.groupby("mkt_al")):
            rows.append(summarize(g, "MKT_AIRLINE", al, "ALL_DATES"))

    # ── 4. By OPERATING airline (aggregated across all dates) ──
    # For EY/EK/QR, mkt_al = op_al, so this mirrors the MKT_AIRLINE group.
    # Kept for structural consistency and future codeshare expansion.
    if "op_al" in df.columns:
        for al, g in sorted(df.groupby("op_al")):
            rows.append(summarize(g, "OP_AIRLINE", al, "ALL_DATES"))

    # ── 5. By marketing airline × date (daily brand-level trend) ──
    if "mkt_al" in df.columns and "date" in df.columns:
        for (al, date), g in sorted(df.groupby(["mkt_al", "date"])):
            rows.append(summarize(g, "MKT_AIRLINE", al, fmt_date(date)))

    # ── 6. By operating airline × date (daily carrier-level trend) ──
    if "op_al" in df.columns and "date" in df.columns:
        for (al, date), g in sorted(df.groupby(["op_al", "date"])):
            rows.append(summarize(g, "OP_AIRLINE", al, fmt_date(date)))

    # ── 7. By route (aggregated across all dates) ──
    if "route" in df.columns:
        for route, g in sorted(df.groupby("route")):
            rows.append(summarize(g, "ROUTE", route, "ALL_DATES"))

    # ── 8. By flight number (aggregated across all dates) ──
    if "flight_fa" in df.columns:
        for fn, g in sorted(df.groupby("flight_fa")):
            rows.append(summarize(g, "FLIGHT", fn, "ALL_DATES"))

    return pd.DataFrame(rows)


# ──────────────────────────────────────────────────────────────────────────────
# INTERACTIVE DATE INPUT
# ──────────────────────────────────────────────────────────────────────────────
# At the start of execution, the user is prompted for:
#   1. A reference date (YYYY-MM-DD) — the "anchor" date
#   2. Days back (integer) — how many additional days before the anchor
#
# The logic:
#   days_back = 0  → scrape only the reference date itself
#   days_back = 1  → scrape the reference date AND the day before (2 days)
#   days_back = 2  → scrape the reference date AND 2 days before (3 days)
#
# Example: date=2026-04-17, days_back=1 → scrapes [2026-04-17, 2026-04-16]

def suggest_workers() -> int:
    """
    Suggest an optimal number of OTP Agents (parallel browser instances).

    Since FlightAware sits behind Cloudflare, the bottleneck is NOT your
    machine — it's Cloudflare's bot detection. Too many parallel requests
    from the same IP trigger the "Just a moment..." challenge on EVERY
    request, which kills throughput.

    Empirical sweet spot for Cloudflare-protected sites: 4-6 workers.
    More than 8 and you start getting challenged on most requests.
    """
    try:
        cpu_count = os.cpu_count() or 4
    except Exception:
        cpu_count = 4

    ram_based_max = 8
    try:
        import psutil
        free_ram_gb = psutil.virtual_memory().available / (1024 ** 3)
        ram_based_max = int(free_ram_gb / 0.3)
    except ImportError:
        pass

    # Cap at 6 — higher just makes Cloudflare angrier without net speed gain
    suggested = min(cpu_count, ram_based_max, 6)
    return max(suggested, 2)


def prompt_user_inputs() -> dict:
    """
    Interactive startup: ask the user for date, days back, and worker count.
    Press Enter on any question to accept the default.
    Returns dict with 'dates' and 'workers'.
    """
    print()
    print("┌─────────────────────────────────────────────┐")
    print("│     Master Webscrapping CQ - v1.0            │")
    print("│     FlightAware Airline OTP Scraper          │")
    print("└─────────────────────────────────────────────┘")
    print()

    # ── Ask for reference date ──
    while True:
        date_input = input("  Reference date (YYYY-MM-DD) [default: yesterday]: ").strip()

        if date_input == "":
            ref_date = datetime.now() - timedelta(days=1)
            break

        try:
            ref_date = datetime.strptime(date_input, "%Y-%m-%d")
            break
        except ValueError:
            print("    Invalid format. Use YYYY-MM-DD (e.g. 2026-04-17)")

    # ── Ask for days back ──
    while True:
        days_input = input("  Days back from that date (0=same day, 1=+1 day back, ...) [default: 0]: ").strip()

        if days_input == "":
            days_back = 0
            break

        try:
            days_back = int(days_input)
            if days_back < 0:
                print("    Enter 0 or a positive number.")
                continue
            break
        except ValueError:
            print("    Enter a number (0, 1, 2, ...)")

    # ── Build the list of dates ──
    dates = []
    for i in range(days_back + 1):
        d = ref_date - timedelta(days=i)
        dates.append(d.strftime("%Y%m%d"))

    # Show confirmation
    date_strs = [f"{d[:4]}-{d[4:6]}-{d[6:]}" for d in dates]
    print(f"\n  -> Dates to scrape: {', '.join(date_strs)} ({len(dates)} day{'s' if len(dates)>1 else ''})")

    # ── Ask for OTP Agents (parallel workers) ──
    # Each agent is a Chrome instance that scrapes independently. More agents
    # = faster, but too many can overwhelm CPU/RAM or trigger rate limits.
    # Suggest an optimal count based on the machine's resources.
    suggested = suggest_workers()
    while True:
        w_input = input(f"  Number of OTP Agents (parallel workers) [default: {suggested} — recommended]: ").strip()

        if w_input == "":
            workers = suggested
            break

        try:
            workers = int(w_input)
            if workers < 1:
                print("    Enter a number >= 1.")
                continue
            if workers > 20:
                print(f"    Warning: {workers} agents is high — may cause instability. Continuing anyway.")
            break
        except ValueError:
            print("    Enter a number (1, 2, 3, ...)")

    print(f"  -> Launching {workers} OTP Agents")

    # ── Ask for flight limit ──
    # By default we scrape everything the Cirium schedule returns for the
    # target dates. But for quick tests / debugging / sampling, it's useful
    # to limit to the first N PER DAY. With multiple dates, the limit applies
    # separately to each date so you get coverage across all selected days.
    while True:
        limit_input = input("  How many flights PER DAY to scrape (number or ENTER for ALL): ").strip()

        if limit_input == "" or limit_input.lower() in ("all", "todos"):
            flight_limit = None   # None = no limit
            break

        try:
            flight_limit = int(limit_input)
            if flight_limit < 1:
                print("    Enter a positive number or ENTER for all.")
                continue
            break
        except ValueError:
            print("    Enter a number (e.g. 50) or press ENTER for all.")

    if flight_limit:
        print(f"  -> Will scrape up to {flight_limit} flights per day")
    else:
        print(f"  -> Will scrape ALL flights in the schedule")
    print()

    return {"dates": dates, "workers": workers, "flight_limit": flight_limit}


# ──────────────────────────────────────────────────────────────────────────────
# HTML REPORT GENERATOR
# ──────────────────────────────────────────────────────────────────────────────

def generate_html_report(df: pd.DataFrame, summary: pd.DataFrame, ok_df: pd.DataFrame,
                         fail_df: pd.DataFrame, run_stats: dict, config, ts: str) -> str:
    """Generate OTP HTML report matching the ey-analysis.html design system. Returns output path."""
    import html as _html_mod

    out_path    = os.path.join(config.output_dir, f"otp_report_{ts}.html")
    latest_path = os.path.join(config.output_dir, "otp_latest.html")

    # ── formatting helpers ─────────────────────────────────────────────────────
    def fv(v, d=1):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "—"
        try:
            return f"{float(v):.{d}f}"
        except (TypeError, ValueError):
            return str(v)

    def fp(v):
        try:
            return f"{float(v):.1f}%"
        except (TypeError, ValueError):
            return "—"

    def fd(v):
        try:
            f = float(v); sign = "+" if f > 0 else ""
            return f"{sign}{f:.1f}"
        except (TypeError, ValueError):
            return "—"

    def pclass(v, good=85.0, warn=70.0):
        try:
            f = float(v)
            return "c-teal" if f >= good else ("c-gold" if f >= warn else "c-crimson")
        except (TypeError, ValueError):
            return "c-mute"

    def dclass(v):
        try:
            f = float(v)
            return "c-teal" if f <= 0 else ("c-gold" if f <= 15 else "c-crimson")
        except (TypeError, ValueError):
            return "c-mute"

    def esc(s):
        return _html_mod.escape(str(s) if s is not None else "")

    # ── run metadata ───────────────────────────────────────────────────────────
    airline   = config.airline_filter or "ALL"
    dates_str = (", ".join(sorted(df["Date"].dropna().astype(str).unique()))
                 if "Date" in df.columns else "—")
    n_ok    = run_stats["ok_count"]
    n_fail  = run_stats["fail_count"]
    n_total = run_stats["total_tasks"]

    ov: dict = {}
    if not summary.empty:
        _ov = summary[summary["scope_type"] == "OVERALL"]
        if not _ov.empty:
            ov = _ov.iloc[0].to_dict()

    route_rows = pd.DataFrame()
    flight_rows = pd.DataFrame()
    if not summary.empty:
        _r = summary[summary["scope_type"] == "ROUTE"].copy()
        if not _r.empty:
            route_rows = _r.sort_values("avg_arr_delay_min", ascending=False,
                                        na_position="last").head(25)
        _f = summary[summary["scope_type"] == "FLIGHT"].copy()
        if not _f.empty:
            flight_rows = _f.sort_values("avg_arr_delay_min", ascending=False,
                                         na_position="last")

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── CSS — matches ey-analysis.html design system ───────────────────────────
    css = """
:root{
  --bg-deep:#0a0d12;--bg-panel:#11151c;--bg-card:#161b24;--bg-elev:#1d2330;
  --line:#232a39;--line-soft:#1a2030;--ink:#e8ecf2;--ink-dim:#8a93a6;--ink-mute:#5a6478;
  --copper:#d97f4a;--gold:#e8b94a;--teal:#4fb8a8;--crimson:#e54e4e;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{background:var(--bg-deep);color:var(--ink);font-family:'Inter Tight',sans-serif;-webkit-font-smoothing:antialiased;line-height:1.5}
.shell{max-width:1480px;margin:0 auto;padding:0 32px}
.eyebrow{font-family:'JetBrains Mono',monospace;font-size:11px;letter-spacing:.2em;text-transform:uppercase;color:var(--copper);font-weight:600}
.c-teal{color:var(--teal)}.c-gold{color:var(--gold)}.c-crimson{color:var(--crimson)}.c-mute{color:var(--ink-mute)}
.nav-back{background:var(--bg-panel);border-bottom:1px solid var(--line);padding:10px 0}
.nav-back a{font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:var(--ink-mute);text-decoration:none;display:inline-flex;align-items:center;gap:8px;transition:color .15s}
.nav-back a:hover{color:var(--teal)}
.hero{background:radial-gradient(ellipse at 15% 0%,rgba(79,184,168,.1) 0%,transparent 55%),radial-gradient(ellipse at 85% 100%,rgba(217,127,74,.1) 0%,transparent 50%);border-bottom:1px solid var(--line);padding:52px 0 56px;position:relative;overflow:hidden}
.hero::before{content:"";position:absolute;inset:0;background-image:repeating-linear-gradient(0deg,transparent 0,transparent 39px,rgba(255,255,255,.012) 39px,rgba(255,255,255,.012) 40px);pointer-events:none}
.hero-inner{display:grid;grid-template-columns:1.4fr 1fr;gap:48px;align-items:end;position:relative}
.hero h1{font-family:'Fraunces',serif;font-weight:700;font-size:clamp(40px,5vw,72px);line-height:.93;letter-spacing:-.038em}
.hero h1 em{font-style:italic;color:var(--teal);font-weight:400}
.hero-sub{font-size:14px;color:var(--ink-dim);max-width:560px;margin-top:22px;line-height:1.7}
.hero-sub strong{color:var(--ink);font-weight:500}
.version-tag{display:inline-block;padding:3px 10px;background:var(--teal);color:var(--bg-deep);font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.2em;font-weight:700;margin-bottom:14px}
.hero-stamps{display:flex;flex-direction:column;gap:10px;align-items:flex-end}
.stamp{display:inline-flex;align-items:center;gap:10px;padding:8px 14px;border:1px solid var(--line);background:var(--bg-panel);border-radius:2px;font-family:'JetBrains Mono',monospace;font-size:11px;letter-spacing:.1em;color:var(--ink-dim)}
.stamp .dot{width:6px;height:6px;border-radius:50%}
.ticker{display:grid;grid-template-columns:repeat(5,1fr);border-top:1px solid var(--line);border-bottom:1px solid var(--line);background:var(--bg-panel)}
.ticker .st{padding:20px 22px;border-right:1px solid var(--line)}
.ticker .st:last-child{border-right:none}
.ticker .num{font-family:'Fraunces',serif;font-size:34px;font-weight:600;line-height:1;letter-spacing:-.02em}
.ticker .lbl{margin-top:7px;font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.13em;text-transform:uppercase;color:var(--ink-mute)}
.ticker .sub{font-family:'JetBrains Mono',monospace;font-size:9px;color:var(--ink-mute);margin-top:3px}
section{padding:56px 0;border-bottom:1px solid var(--line-soft)}
.section-head{display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:32px;gap:32px}
.section-head h2{font-family:'Fraunces',serif;font-size:34px;font-weight:600;letter-spacing:-.025em;line-height:1.05}
.section-head h2 em{font-style:italic;color:var(--copper);font-weight:400}
.section-head .desc{max-width:480px;font-size:14px;color:var(--ink-dim);line-height:1.65}
.otp-two-col{display:grid;grid-template-columns:1fr 1fr;gap:20px}
.otp-panel{background:var(--bg-panel);border:1px solid var(--line);padding:24px}
.otp-panel-head{font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:var(--ink-mute);margin-bottom:16px;padding-bottom:10px;border-bottom:1px solid var(--line)}
.otp-row{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid var(--line-soft)}
.otp-row:last-child{border-bottom:none}
.otp-metric{font-size:13px;color:var(--ink-dim)}
.otp-val{font-family:'Fraunces',serif;font-size:20px;font-weight:600;letter-spacing:-.01em}
.otp-row.stat-row .otp-val{font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:400}
.route-table{width:100%;border-collapse:collapse;font-size:13px}
.route-table th{font-family:'JetBrains Mono',monospace;font-size:9px;letter-spacing:.13em;text-transform:uppercase;color:var(--ink-mute);padding:8px 10px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
.route-table th.r{text-align:right}
.route-table td{padding:8px 10px;border-bottom:1px solid var(--line-soft);vertical-align:middle}
.route-table td.r{text-align:right;font-family:'JetBrains Mono',monospace;font-size:12px}
.route-table tr:hover td{background:rgba(255,255,255,.018)}
.rt-code{font-family:'Fraunces',serif;font-size:15px;font-weight:600;letter-spacing:-.01em}
.tbl-wrap{overflow-x:auto}
.pill{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:9px;letter-spacing:.09em;padding:2px 8px;font-weight:600}
.pill-ok{background:rgba(79,184,168,.1);color:var(--teal)}
.pill-cancel{background:rgba(229,78,78,.1);color:var(--crimson)}
.pill-divert{background:rgba(232,185,74,.1);color:var(--gold)}
.pill-inflight{background:rgba(96,165,250,.1);color:#60a5fa}
.chk{color:var(--teal);font-weight:700}.crs{color:var(--crimson);font-weight:700}
footer{padding:40px 0 52px;color:var(--ink-mute);font-family:'JetBrains Mono',monospace;font-size:11px;letter-spacing:.08em;text-align:center;border-top:1px solid var(--line);line-height:2}
footer em{font-family:'Fraunces',serif;font-style:italic;color:var(--copper)}
footer a{color:var(--ink-mute);text-decoration:none}
footer a:hover{color:var(--teal)}
@media(max-width:960px){
  .hero-inner,.otp-two-col{grid-template-columns:1fr}
  .hero-stamps{align-items:flex-start}
  .ticker{grid-template-columns:repeat(2,1fr)}
  .ticker .st{border-bottom:1px solid var(--line)}
  .section-head{flex-direction:column;align-items:flex-start}
}
@media(max-width:580px){.shell{padding:0 18px}.ticker{grid-template-columns:1fr}}
"""

    # ── ticker helper ──────────────────────────────────────────────────────────
    d15  = ov.get("D15_%");  a15  = ov.get("A15_%")
    ddel = ov.get("avg_dep_delay_min"); adel = ov.get("avg_arr_delay_min")
    wb   = ov.get("within_block_%")

    def ticker_st(lbl, val, sub, cls):
        return (f'<div class="st"><div class="num {cls}">{esc(val)}</div>'
                f'<div class="lbl">{esc(lbl)}</div>'
                f'<div class="sub">{esc(sub)}</div></div>')

    ticker_html = (
        '<div class="ticker">'
        + ticker_st("D+15 Dep OTP",  fp(d15),  "< 15 min late",         pclass(d15))
        + ticker_st("A+15 Arr OTP",  fp(a15),  "< 15 min late",         pclass(a15))
        + ticker_st("Avg Dep Delay", f"{fd(ddel)} min", "all completed", dclass(ddel))
        + ticker_st("Avg Arr Delay", f"{fd(adel)} min", "all completed", dclass(adel))
        + ticker_st("Within Block",  fp(wb),   "actual ≤ scheduled",    pclass(wb, good=70, warn=50))
        + '</div>'
    )

    # ── OTP panels ─────────────────────────────────────────────────────────────
    def otp_panel(title, metrics, avg_key, med_key, p90_key):
        rows_h = ""
        for lbl, key in metrics:
            v = ov.get(key)
            rows_h += (f'<div class="otp-row"><span class="otp-metric">{esc(lbl)}</span>'
                       f'<span class="otp-val {pclass(v)}">{fp(v)}</span></div>')
        for stat_lbl, stat_key in [("Avg delay", avg_key), ("Median", med_key), ("P90", p90_key)]:
            sv = ov.get(stat_key)
            rows_h += (f'<div class="otp-row stat-row">'
                       f'<span class="otp-metric" style="color:var(--ink-mute)">{esc(stat_lbl)}</span>'
                       f'<span class="otp-val {dclass(sv)}">{fd(sv)} min</span></div>')
        return f'<div class="otp-panel"><div class="otp-panel-head">{esc(title)}</div>{rows_h}</div>'

    dep_m = [("D+0 on time / early","D0_%"),("D+5 < 5 min","D5_%"),("D+15 < 15 min","D15_%"),
             ("D+30 < 30 min","D30_%"),("D+60 < 60 min","D60_%"),("D+360 < 6 hrs","D360_%")]
    arr_m = [("A+0 on time / early","A0_%"),("A+14 DOT standard","A14_%"),("A+15 < 15 min","A15_%"),
             ("A+30 < 30 min","A30_%"),("A+60 < 60 min","A60_%"),("A+360 < 6 hrs","A360_%")]
    otp_panels_html = (
        '<div class="otp-two-col">'
        + otp_panel("Departure Performance", dep_m, "avg_dep_delay_min", "med_dep_delay_min", "p90_dep_delay_min")
        + otp_panel("Arrival Performance",   arr_m, "avg_arr_delay_min", "med_arr_delay_min", "p90_arr_delay_min")
        + '</div>'
    )

    # ── route table ────────────────────────────────────────────────────────────
    def route_table_html(rows_df: pd.DataFrame) -> str:
        if rows_df.empty:
            return '<p style="color:var(--ink-mute);font-family:\'JetBrains Mono\',monospace;font-size:11px;padding:12px 0">No route data.</p>'
        thead = ("<tr><th>Route</th><th class='r'>Flights</th><th class='r'>D+15%</th>"
                 "<th class='r'>A+15%</th><th class='r'>Avg Dep</th><th class='r'>Avg Arr</th>"
                 "<th class='r'>P90 Arr</th><th class='r'>Within Blk</th></tr>")
        tbody = ""
        for _, row in rows_df.iterrows():
            v_d15 = row.get("D15_%"); v_a15 = row.get("A15_%")
            v_dd = row.get("avg_dep_delay_min"); v_ad = row.get("avg_arr_delay_min")
            v_p90 = row.get("p90_arr_delay_min"); v_wb = row.get("within_block_%")
            tbody += (f'<tr><td><span class="rt-code">{esc(str(row.get("scope_value","—")))}</span></td>'
                      f'<td class="r">{esc(fv(row.get("flights"),0))}</td>'
                      f'<td class="r {pclass(v_d15)}">{fp(v_d15)}</td>'
                      f'<td class="r {pclass(v_a15)}">{fp(v_a15)}</td>'
                      f'<td class="r {dclass(v_dd)}">{fd(v_dd)} min</td>'
                      f'<td class="r {dclass(v_ad)}">{fd(v_ad)} min</td>'
                      f'<td class="r {dclass(v_p90)}">{fd(v_p90)} min</td>'
                      f'<td class="r {pclass(v_wb,good=70,warn=50)}">{fp(v_wb)}</td></tr>')
        return f'<div class="tbl-wrap"><table class="route-table"><thead>{thead}</thead><tbody>{tbody}</tbody></table></div>'

    # ── flight number table ────────────────────────────────────────────────────
    def flight_num_table_html(rows_df: pd.DataFrame) -> str:
        if rows_df.empty:
            return '<p style="color:var(--ink-mute);font-family:\'JetBrains Mono\',monospace;font-size:11px;padding:12px 0">No data.</p>'
        thead = ("<tr><th>Flight</th><th class='r'>Legs</th><th class='r'>D+15%</th>"
                 "<th class='r'>A+15%</th><th class='r'>Avg Dep</th>"
                 "<th class='r'>Avg Arr</th><th class='r'>Cancelled</th></tr>")
        tbody = ""
        for _, row in rows_df.iterrows():
            v_d15 = row.get("D15_%"); v_a15 = row.get("A15_%")
            v_dd = row.get("avg_dep_delay_min"); v_ad = row.get("avg_arr_delay_min")
            tbody += (f'<tr><td><span class="rt-code">{esc(str(row.get("scope_value","—")))}</span></td>'
                      f'<td class="r">{esc(fv(row.get("flights"),0))}</td>'
                      f'<td class="r {pclass(v_d15)}">{fp(v_d15)}</td>'
                      f'<td class="r {pclass(v_a15)}">{fp(v_a15)}</td>'
                      f'<td class="r {dclass(v_dd)}">{fd(v_dd)} min</td>'
                      f'<td class="r {dclass(v_ad)}">{fd(v_ad)} min</td>'
                      f'<td class="r">{fp(row.get("cancelled_%",0))}</td></tr>')
        return f'<div class="tbl-wrap"><table class="route-table"><thead>{thead}</thead><tbody>{tbody}</tbody></table></div>'

    # ── flight detail table ────────────────────────────────────────────────────
    def detail_table_html(rows_df: pd.DataFrame) -> str:
        if rows_df.empty:
            return '<p style="color:var(--ink-mute);font-family:\'JetBrains Mono\',monospace;font-size:11px;padding:12px 0">No flights matched.</p>'
        detail_cols = [
            ("Flight","flight_fa",False),("Date","Date",False),("Route","route",False),
            ("Status","flight_status",False),("Dep Delay","dep_delay_min",True),
            ("Arr Delay","arr_delay_min",True),("D+15","D15",True),("A+15","A15",True),
            ("Act Block","actual_block_min",True),("Sch Block","sched_block_min",True),
            ("Aircraft","aircraft_type_fa",True),("Reg","tail_number",False),
        ]
        present = [(lbl, key, r) for lbl, key, r in detail_cols if key in rows_df.columns]
        th_parts = []
        for lbl, key, r in present:
            cls = ' class="r"' if r else ''
            th_parts.append(f'<th{cls}>{lbl}</th>')
        thead = "<tr>" + "".join(th_parts) + "</tr>"
        sort_keys = [c for c in ["date", "flight_fa"] if c in rows_df.columns]
        display_df = rows_df.sort_values(sort_keys) if sort_keys else rows_df
        tbody = ""
        for _, row in display_df.iterrows():
            cells = []
            for lbl, key, r in present:
                v = row.get(key)
                rcls = ' class="r"' if r else ''
                if key == "flight_status":
                    s = str(v).upper() if v else ""
                    if "CANCEL" in s:
                        cells.append('<td><span class="pill pill-cancel">CANCELLED</span></td>')
                    elif "DIVERT" in s:
                        cells.append('<td><span class="pill pill-divert">DIVERTED</span></td>')
                    elif "IN_FLIGHT" in s or "IN FLIGHT" in s:
                        cells.append('<td><span class="pill pill-inflight">IN FLIGHT</span></td>')
                    else:
                        cells.append('<td><span class="pill pill-ok">LANDED</span></td>')
                elif key in ("dep_delay_min", "arr_delay_min"):
                    cells.append(f'<td class="r {dclass(v)}">{fd(v)} min</td>')
                elif key in ("D15", "A15"):
                    try:
                        vi = int(float(v))
                        sym_cls = "chk" if vi == 1 else "crs"
                        sym = "&#10003;" if vi == 1 else "&#10007;"
                        cells.append(f'<td class="r"><span class="{sym_cls}">{sym}</span></td>')
                    except (TypeError, ValueError):
                        cells.append('<td class="r" style="color:var(--ink-mute)">&#8212;</td>')
                elif key == "flight_fa":
                    cells.append(f'<td><span class="rt-code">{esc(v)}</span></td>')
                elif key in ("actual_block_min", "sched_block_min"):
                    cells.append(f'<td class="r" style="color:var(--ink-dim)">{fv(v,0)} min</td>')
                else:
                    sv = str(v) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "&#8212;"
                    cells.append(f'<td{rcls}>{esc(sv)}</td>')
            tbody += "<tr>" + "".join(cells) + "</tr>"
        return f'<div class="tbl-wrap"><table class="route-table"><thead>{thead}</thead><tbody>{tbody}</tbody></table></div>'

    # ── coverage / failed section ──────────────────────────────────────────────
    def coverage_html(rows_df: pd.DataFrame) -> str:
        if rows_df.empty:
            return (f'<div style="display:inline-flex;align-items:center;gap:10px;'
                    f'background:rgba(79,184,168,.07);border:1px solid rgba(79,184,168,.22);'
                    f'padding:12px 20px;font-family:\'JetBrains Mono\',monospace;font-size:11px;'
                    f'letter-spacing:.06em;color:var(--teal)">'
                    f'&#10003; All {n_ok} flights matched — 0 unmatched</div>')
        cols = [c for c in ["Date", "flight_fa", "route", "scrape_status"] if c in rows_df.columns]
        th_parts = [f'<th{"" if i > 0 else ""}>{esc(c)}</th>' for i, c in enumerate(cols)]
        thead = "<tr>" + "".join(th_parts) + "</tr>"
        tbody = ""
        for _, row in rows_df.iterrows():
            cells = [f'<td>{esc(str(row.get(c,"—")))}</td>' for c in cols]
            tbody += "<tr>" + "".join(cells) + "</tr>"
        return f'<div class="tbl-wrap"><table class="route-table"><thead>{thead}</thead><tbody>{tbody}</tbody></table></div>'

    # ── hero sub text ──────────────────────────────────────────────────────────
    canc_pct = ov.get("cancelled_%", 0) or 0
    hero_sub = (
        f'<strong>{n_ok} of {n_total} flights</strong> matched and analysed for {esc(dates_str)}. '
        f'D+15 OTP: <strong class="{pclass(d15)}">{fp(d15)}</strong> departure, '
        f'<strong class="{pclass(a15)}">{fp(a15)}</strong> arrival. '
        f'Cancellation rate {fp(canc_pct)}. '
        f'Source: {esc(run_stats.get("engine","—"))}.'
    )
    fail_color = "var(--crimson)" if n_fail > 0 else "var(--teal)"
    failed_label = f"Failed / Not Matched ({n_fail})" if n_fail else "Coverage"

    # ── assemble HTML ──────────────────────────────────────────────────────────
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OTP Report — {esc(airline)} {esc(dates_str)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,300;9..144,400;9..144,600;9..144,700;9..144,900&family=JetBrains+Mono:wght@300;400;500;600;700&family=Inter+Tight:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>{css}</style>
</head>
<body>

<div class="nav-back">
  <div class="shell"><a href="../ey-analysis.html">&#8592; Network Analysis</a></div>
</div>

<header class="hero">
  <div class="shell">
    <div class="hero-inner">
      <div>
        <span class="version-tag">OPERATIONS INTELLIGENCE</span>
        <div class="eyebrow" style="margin-bottom:16px">{esc(airline)} · {esc(dates_str)} · {esc(run_stats.get("engine","—"))}</div>
        <h1>On-Time<br><em>Performance</em></h1>
        <p class="hero-sub">{hero_sub}</p>
      </div>
      <div class="hero-stamps">
        <div class="stamp"><span class="dot" style="background:var(--teal);box-shadow:0 0 8px var(--teal)"></span>{n_ok} FLIGHTS MATCHED</div>
        <div class="stamp"><span class="dot" style="background:{fail_color}"></span>{n_fail} NOT FOUND</div>
        <div class="stamp"><span class="dot" style="background:var(--copper)"></span>RUN TIME {esc(run_stats.get("elapsed_human","—"))}</div>
        <div style="margin-top:6px;font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.1em;color:var(--ink-mute);text-align:right">
          Generated {esc(generated_at)}<br>Cirium schedule matched
        </div>
      </div>
    </div>
  </div>
</header>

{ticker_html}

<section>
  <div class="shell">
    <div class="section-head">
      <div>
        <div class="eyebrow">OTP Thresholds</div>
        <h2>Departure &amp; Arrival <em>Performance</em></h2>
      </div>
      <p class="desc">Computed against scheduled gate times from the Cirium schedule. Actual times from FR24 wheel-off / wheel-on converted to UTC epoch.</p>
    </div>
    {otp_panels_html}
  </div>
</section>

<section>
  <div class="shell">
    <div class="section-head">
      <div>
        <div class="eyebrow">Route Analysis</div>
        <h2>Routes by Arrival <em>Delay</em></h2>
      </div>
      <p class="desc">Worst routes first. Top 25 shown.</p>
    </div>
    {route_table_html(route_rows)}
  </div>
</section>

<section>
  <div class="shell">
    <div class="section-head">
      <div>
        <div class="eyebrow">Flight Numbers</div>
        <h2>Per-Flight <em>OTP</em></h2>
      </div>
      <p class="desc">Individual flight number performance sorted by average arrival delay.</p>
    </div>
    {flight_num_table_html(flight_rows)}
  </div>
</section>

<section>
  <div class="shell">
    <div class="section-head">
      <div>
        <div class="eyebrow">Raw Data</div>
        <h2>Flight <em>Detail</em></h2>
      </div>
      <p class="desc">{n_ok} flights with actual departure and arrival times.</p>
    </div>
    {detail_table_html(ok_df)}
  </div>
</section>

<section>
  <div class="shell">
    <div class="section-head">
      <div>
        <div class="eyebrow">{esc(failed_label)}</div>
        <h2>Schedule <em>Coverage</em></h2>
      </div>
      <p class="desc">Flights in the Cirium schedule not matched in the FR24 data window.</p>
    </div>
    {coverage_html(fail_df)}
  </div>
</section>

<footer>
  <div class="shell">
    <em>Master Webscrapping CQ</em> · OTP Intelligence · {esc(generated_at)}<br>
    {esc(run_stats.get("engine","—"))} · Cirium Summer Schedule · IATA OTP standard<br>
    <a href="../ey-analysis.html">&#8592; Back to Network Analysis</a>
  </div>
</footer>

</body>
</html>"""

    for path in (out_path, latest_path):
        with open(path, "w", encoding="utf-8") as _f:
            _f.write(html_content)

    return out_path


# ──────────────────────────────────────────────────────────────────────────────
# MAIN — Orchestrate everything
# ──────────────────────────────────────────────────────────────────────────────

def main():
    # ══════════════════════════════════════════════════════════════════════
    # CONFIGURATION — Edit these to match your local environment
    # ══════════════════════════════════════════════════════════════════════

    SCHEDULE_FILE = "SummerS_EY_QR_EK.xlsx"  # Cirium schedule file (same folder as this .py)
    AIRPORTS_FILE = "airports.csv"           # Auto-downloaded from ourairports.com if missing
    ROUTES_FILTER = ""                       # "AUH-LHR,DXB-JFK" or "" for all routes
    OUTPUT_DIR = "output"                    # Output folder
    HEADLESS = True                          # True = invisible browser, False = see the browser
    MIN_DELAY = 2.0                          # Min seconds between requests per worker
    MAX_DELAY = 5.0                          # Max seconds between requests per worker

    # ══════════════════════════════════════════════════════════════════════
    # STEP 0: Interactive input — date, days back, and OTP Agents count
    # ══════════════════════════════════════════════════════════════════════
    # Three questions, Enter = default. All airlines in the schedule are
    # always scraped (the output CSV has an op_al column to filter later).

    user_input = prompt_user_inputs()
    dates = user_input["dates"]
    WORKERS = user_input["workers"]
    FLIGHT_LIMIT = user_input["flight_limit"]   # None = scrape all

    # ── Data source selection ──
    FR24_API_KEY = os.environ.get("FR24_API_KEY", "")
    print("  Data source:")
    print("    [1] FlightAware  (browser scraping — Cloudflare bypass)")
    print("    [2] FR24 API     (direct API — fast, no browser needed)")
    while True:
        src_input = input("  Choose data source [default: 1]: ").strip()
        if src_input == "" or src_input == "1":
            USE_FR24 = False
            print("  -> Data source: FlightAware (browser scraping)")
            break
        elif src_input == "2":
            USE_FR24 = True
            if not FR24_API_KEY:
                print("  WARNING: FR24_API_KEY environment variable not set.")
                print("  Set it before running:")
                print('    $env:FR24_API_KEY="your-key-here"   (PowerShell)')
                print('    export FR24_API_KEY="your-key-here" (bash)')
            print("  -> Data source: FR24 API")
            break
        else:
            print("  Enter 1 or 2.")

    config = Config(
        max_workers=WORKERS,
        headless=HEADLESS,
        output_dir=OUTPUT_DIR,
        airline_filter="EY",  # Scope: Etihad Airways full schedule
        routes_filter=[r.strip().upper() for r in ROUTES_FILTER.split(",") if r.strip()],
        min_delay_sec=MIN_DELAY,
        max_delay_sec=MAX_DELAY,
    )

    os.makedirs(config.output_dir, exist_ok=True)

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 1: Load airport codes (auto-download if missing)
    # ═══════════════════════════════════════════════════════════════════════
    # We need airports.csv to convert IATA codes (AUH, LHR) from the Cirium
    # schedule into ICAO codes (OMAA, EGLL) for the FlightAware URLs.
    # If the file doesn't exist locally, we download it automatically.

    if not os.path.exists(AIRPORTS_FILE):
        log.info(f"airports.csv not found locally — downloading from ourairports.com...")
        try:
            import urllib.request
            urllib.request.urlretrieve(
                "https://davidmegginson.github.io/ourairports-data/airports.csv",
                AIRPORTS_FILE,
            )
            log.info(f"Downloaded airports.csv ({os.path.getsize(AIRPORTS_FILE) / 1024 / 1024:.1f} MB)")
        except Exception as e:
            log.error(f"Could not download airports.csv: {e}")
            log.info("Download manually from: https://ourairports.com/data/airports.csv")
            log.info(f"Place it in: {os.path.abspath('.')}")
            sys.exit(1)

    airport_lookup = load_airport_lookup(AIRPORTS_FILE)
    coord_offsets  = load_airport_coord_offsets(AIRPORTS_FILE)

    # Pre-warm the offset cache for Gulf hubs — AUH/DXB/DOH are UTC+4/+3
    # with no DST, so the offset is constant year-round. This eliminates
    # the double-page-load penalty for every hub departure (the majority of
    # EY/EK/QR legs originate at their home hub).
    with OFFSET_CACHE_LOCK:
        OFFSET_CACHE.update(HUB_UTC_OFFSETS)
    log.info(f"Hub offsets pre-warmed: {HUB_UTC_OFFSETS}")

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 2: Load Cirium schedule and build task list
    # ═══════════════════════════════════════════════════════════════════════
    # The schedule XLSX must be in the same folder as this .py file.

    if not os.path.exists(SCHEDULE_FILE):
        log.error(f"Schedule file not found: {os.path.abspath(SCHEDULE_FILE)}")
        sys.exit(1)

    schedule_df = load_cirium_schedule(SCHEDULE_FILE, dates, config)

    if schedule_df.empty:
        sys.exit(0)

    # Build task list — one task per flight.
    # We do NOT pre-compute Zulu URLs anymore. Instead, the scraper workers
    # share an in-memory offset cache (keyed by origin ICAO). The FIRST flight
    # from each airport pays the cost of one extra round-trip: hit the generic
    # URL, read FlightAware's displayed offset (e.g. "-05"), then redirect to
    # the correct Zulu URL. All subsequent flights from that same airport use
    # the cached offset to build the Zulu URL directly on the first hit.
    # This is simpler and more accurate than any pre-computation.
    tasks = []
    for _, row in schedule_df.iterrows():
        orig_icao = to_icao(row["Orig"], airport_lookup)
        dest_icao = to_icao(row["Dest"], airport_lookup)

        tasks.append({
            "flight_fa": row["flight_fa"],
            "mkt_al": row["Mkt_Al"],      # Marketing airline — used in URL
            "op_al": row["Op_Al"],        # Operating airline — who actually flew it
            "orig": row["Orig"],
            "dest": row["Dest"],
            "orig_icao": orig_icao,
            "dest_icao": dest_icao,
            "route": f"{row['Orig']}-{row['Dest']}",
            "date_str": row["date_str"],
            "url": build_url(row["flight_fa"], row["date_str"], orig_icao, dest_icao),
            # Carry schedule data for later merge
            "sched_dep_local": row.get("Dep_Time"),
            "sched_arr_local": row.get("Arr_Time"),
            "sched_block_mins": row.get("Block_Mins"),
            "equip": row.get("Equip"),
            "seats": row.get("Seats"),
            "distance_km": row.get("Kilometers"),
        })

    log.info(f"Task list ready: {len(tasks)} flights to scrape")

    # ── Apply flight limit if the user specified one ──
    # The limit means "first N per date" — NOT "first N total". With multiple
    # dates, you want a sample from each day (e.g. 50 flights from day 1 AND
    # 50 flights from day 2), not just the first 50 flights from day 1 with
    # nothing from day 2. This makes test runs meaningful across the full
    # date range.
    if FLIGHT_LIMIT is not None:
        limited_tasks = []
        for date_str in sorted(set(t["date_str"] for t in tasks)):
            day_tasks = [t for t in tasks if t["date_str"] == date_str]
            kept = day_tasks[:FLIGHT_LIMIT]
            limited_tasks.extend(kept)
            if len(day_tasks) > FLIGHT_LIMIT:
                log.info(f"  {date_str}: keeping {FLIGHT_LIMIT} of {len(day_tasks)} flights")
            else:
                log.info(f"  {date_str}: keeping all {len(day_tasks)} flights (less than limit)")
        tasks = limited_tasks

    log.info(f"Will scrape: {len(tasks)} flights total")

    # ── Estimate total runtime ──
    # Based on actual observed runs (not theoretical):
    #   - A warm-cache flight takes ~8-10 seconds total:
    #       * ~4s of page load + JS render (FlightAware is slow)
    #       * ~2s random throttle between requests
    #       * ~2-4s of Chrome overhead + processing
    #   - A cold-cache flight (first for each airport) takes ~15-20 seconds
    #     because it does TWO round-trips to discover the UTC offset.
    #
    # We assume the cache warms up quickly (most airports seen early),
    # so the effective average is dominated by warm-cache flights.
    unique_origins = len(set(t["orig_icao"] for t in tasks))
    warm_time = 9.0           # Average seconds per warm-cache flight
    cold_extra = 8.0          # Extra seconds per cold-cache flight (first of each airport)
    total_worker_sec = len(tasks) * warm_time + unique_origins * cold_extra
    est_sec = total_worker_sec / config.max_workers
    est_min = est_sec / 60
    log.info(f"Estimated time: ~{est_min:.0f} minutes "
             f"({len(tasks)} flights, {unique_origins} unique origins, "
             f"{config.max_workers} agents)")
    log.info(f"  (Actual speed varies: cold start is slow, then accelerates)")

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 3: Collect actual flight data
    # ═══════════════════════════════════════════════════════════════════════
    if USE_FR24:
        results, run_stats = run_fr24_otp(tasks, config, FR24_API_KEY, coord_offsets)
    else:
        results, run_stats = run_parallel_scrape(tasks, config)

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 4: Compute OTP metrics for each flight
    # ═══════════════════════════════════════════════════════════════════════
    for r in results:
        if r.get("scrape_status") == "OK":
            r.update(compute_metrics(r))

    df = pd.DataFrame(results)

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 5: Save outputs
    # ═══════════════════════════════════════════════════════════════════════
    if df.empty:
        log.warning("No flights were scraped. Check browser engine and network.")
        return

    ts = datetime.now().strftime("%Y%m%d_%H%M")

    # Add a nicely formatted Date column (YYYY-MM-DD) for analysis tools.
    # The existing 'date' field is kept as YYYYMMDD (URL-compatible format).
    if "date" in df.columns:
        df["Date"] = df["date"].astype(str).apply(
            lambda s: f"{s[:4]}-{s[4:6]}-{s[6:]}" if len(s) == 8 and s.isdigit() else s
        )

    # ═══════════════════════════════════════════════════════════════════════
    # Save everything to ONE Excel workbook with multiple sheets:
    #
    #   - Summary         : aggregated OTP metrics (OVERALL, DATE, AIRLINE, ROUTE, FLIGHT)
    #   - All_Flights     : every flight scraped, with full metrics, keyed by Date
    #   - Per-day sheets  : one sheet per scraped date (e.g. "2026-04-16")
    #   - Failed_Scrapes  : only the flights that didn't return data (for debugging)
    #
    # Using XLSX (not CSV) so everything lives in a single file you can
    # share with stakeholders without losing the structure.
    # ═══════════════════════════════════════════════════════════════════════

    ok_df = df[df["scrape_status"] == "OK"].copy()
    fail_df = df[df["scrape_status"] != "OK"].copy()

    output_file = os.path.join(config.output_dir, f"otp_report_{ts}.xlsx")

    # Column ordering: put identifiers first so they're visible in Excel.
    # The Date column goes right after flight_fa for fast filtering.
    preferred_order = [
        "Date", "date", "flight_fa", "mkt_al", "op_al", "orig", "dest", "route",
        "orig_icao", "dest_icao", "flight_status", "status_raw",
        "dep_delay_min", "arr_delay_min",
        "D0", "D5", "D15", "D30", "D60", "D360",
        "A0", "A14", "A15", "A30", "A60", "A360",
        "actual_block_min", "sched_block_min", "block_diff_min", "block_pct", "within_block",
        "completed", "cancelled", "diverted",
        "taxi_out_min", "taxi_in_min", "departure_gate",
        "aircraft_type_fa", "tail_number", "distance_sm",
        "equip_sched", "seats_sched", "distance_km",
        "sched_dep_local", "sched_arr_local", "sched_block_mins",
        "scrape_status", "url",
    ]

    def order_columns(frame):
        """Arrange columns with key identifiers first, keep any extras at the end."""
        cols_present = [c for c in preferred_order if c in frame.columns]
        extras = [c for c in frame.columns if c not in cols_present]
        return frame[cols_present + extras]

    try:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            # ── Sheet 1: Summary metrics (scope_type × scope_value × Date) ──
            if not ok_df.empty:
                summary = generate_summary(ok_df)
                summary.to_excel(writer, sheet_name="Summary", index=False)
                log.info(f"Sheet 'Summary': {len(summary)} rows")
            else:
                summary = pd.DataFrame()

            # ── Sheet 2: All scraped flights with Date column ──
            # Single tab with every flight across all dates. The Date column
            # lets you filter or pivot by day directly in Excel/Power BI,
            # so there's no need for separate per-day sheets.
            all_flights_df = order_columns(df.copy())
            all_flights_df.to_excel(writer, sheet_name="All_Flights", index=False)
            log.info(f"Sheet 'All_Flights': {len(all_flights_df)} rows")

            # ── Sheet 3: Run info — timing and agent stats for this run ──
            # Lets the reader audit performance over time (did we get faster
            # when we bumped agents from 5 to 8? Are we slower today?)
            run_info_rows = [
                ("Run timestamp",       datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                ("Dates scraped",       ", ".join(sorted(df["Date"].dropna().unique()))),
                ("OTP Agents",          run_stats["otp_agents"]),
                ("Browser engine",      run_stats["engine"]),
                ("Total tasks",         run_stats["total_tasks"]),
                ("Successful scrapes",  run_stats["ok_count"]),
                ("Failed scrapes",      run_stats["fail_count"]),
                ("Success rate %",      round(run_stats["ok_count"] / max(run_stats["total_tasks"], 1) * 100, 1)),
                ("Total time",          run_stats["elapsed_human"]),
                ("Total time (sec)",    run_stats["elapsed_sec"]),
                ("Avg sec / flight",    run_stats["avg_sec_per_flight"]),
                ("Throughput (flights/min)", round(60 / max(run_stats["avg_sec_per_flight"], 0.01), 1)),
            ]
            run_info_df = pd.DataFrame(run_info_rows, columns=["Metric", "Value"])
            run_info_df.to_excel(writer, sheet_name="Run_Info", index=False)
            log.info(f"Sheet 'Run_Info': {len(run_info_df)} rows")

            # ── Sheet 4: Failed scrapes (only if there were any) ──
            if not fail_df.empty:
                fail_cols = ["Date", "flight_fa", "route", "scrape_status", "url"]
                fail_cols = [c for c in fail_cols if c in fail_df.columns]
                fail_df[fail_cols].to_excel(writer, sheet_name="Failed_Scrapes", index=False)
                log.info(f"Sheet 'Failed_Scrapes': {len(fail_df)} rows")

        log.info(f"Excel report saved: {output_file}")
    except Exception as e:
        # Fallback to CSV if openpyxl fails for any reason (e.g. corrupted template)
        log.warning(f"Excel write failed ({e}). Falling back to CSV.")
        df.to_csv(os.path.join(config.output_dir, f"otp_raw_{ts}.csv"), index=False)
        if not ok_df.empty:
            summary.to_csv(os.path.join(config.output_dir, f"otp_metrics_{ts}.csv"), index=False)

    # ── HTML report ──
    try:
        html_file = generate_html_report(df, summary, ok_df, fail_df, run_stats, config, ts)
        log.info(f"HTML report saved: {html_file}")
        print(f"HTML:   {html_file}")
    except Exception as _he:
        log.warning(f"HTML report generation failed: {_he}")

    if not ok_df.empty:
        # Console report
        ov = summary[summary["scope_type"] == "OVERALL"].iloc[0]
        print()
        print("┌─────────────────────────────────────────────┐")
        print("│          OTP PERFORMANCE SUMMARY             │")
        print("├─────────────────────────────────────────────┤")
        print(f"│  RUN INFO                                   │")
        print(f"│    OTP Agents:        {run_stats['otp_agents']:>6}                │")
        print(f"│    Total time:        {run_stats['elapsed_human']:>10}            │")
        print(f"│    Avg sec / flight:  {run_stats['avg_sec_per_flight']:>6.2f}                │")
        print(f"│    Throughput:        {60/max(run_stats['avg_sec_per_flight'],0.01):>6.1f} flights/min     │")
        print(f"│                                             │")
        print(f"│  Flights scraped:      {int(ov['flights']):>6}               │")
        print(f"│  Schedule completion:  {ov.get('completed_%', '?'):>6}%              │")
        print(f"│  Cancellation rate:    {ov.get('cancelled_%', '?'):>6}%              │")
        print(f"│                                             │")
        print(f"│  DEPARTURE                                  │")
        print(f"│    D+0   (on time):    {ov.get('D0_%', '?'):>6}%              │")
        print(f"│    D+15  (<15 min):    {ov.get('D15_%', '?'):>6}%              │")
        print(f"│    D+30  (<30 min):    {ov.get('D30_%', '?'):>6}%              │")
        print(f"│    D+60  (<60 min):    {ov.get('D60_%', '?'):>6}%              │")
        print(f"│    D+360 (<6 hours):   {ov.get('D360_%', '?'):>6}%              │")
        print(f"│    Avg delay:          {ov.get('avg_dep_delay_min', '?'):>6} min           │")
        print(f"│                                             │")
        print(f"│  ARRIVAL                                    │")
        print(f"│    A+0   (on time):    {ov.get('A0_%', '?'):>6}%              │")
        print(f"│    A+14  (DOT):        {ov.get('A14_%', '?'):>6}%              │")
        print(f"│    A+15  (<15 min):    {ov.get('A15_%', '?'):>6}%              │")
        print(f"│    A+30  (<30 min):    {ov.get('A30_%', '?'):>6}%              │")
        print(f"│    A+60  (<60 min):    {ov.get('A60_%', '?'):>6}%              │")
        print(f"│    A+360 (<6 hours):   {ov.get('A360_%', '?'):>6}%              │")
        print(f"│    Avg delay:          {ov.get('avg_arr_delay_min', '?'):>6} min           │")
        print(f"│    P90 delay:          {ov.get('p90_arr_delay_min', '?'):>6} min           │")
        print(f"│                                             │")
        print(f"│  BLOCK TIME                                 │")
        print(f"│    % Within block:     {ov.get('within_block_%', '?'):>6}%              │")
        print("└─────────────────────────────────────────────┘")
    else:
        log.warning("No flights scraped successfully.")

    ok_count = len(ok_df)
    fail_count = len(df) - ok_count
    print(f"\nDone: {ok_count} OK / {fail_count} failed / {len(df)} total")
    print(f"Output: {output_file}")

    # ── Write final status + OTP JSON for the live viewer ──
    import json as _json
    status_file = os.path.join(config.output_dir, "scraper_status.json")
    otp_file    = os.path.join(config.output_dir, "otp_latest.json")
    try:
        with open(status_file, "w") as _f:
            _json.dump({
                "status": "done",
                "total": run_stats["total_tasks"],
                "done": run_stats["total_tasks"],
                "ok": run_stats["ok_count"],
                "fail": run_stats["fail_count"],
                "pct": 100.0,
                "elapsed_sec": run_stats["elapsed_sec"],
                "elapsed_human": run_stats["elapsed_human"],
                "last_updated": datetime.now().isoformat(),
                "output_file": output_file,
            }, _f)

        otp_payload = {
            "status": "done",
            "airline": config.airline_filter or "ALL",
            "run_date": ts,
            "elapsed_human": run_stats["elapsed_human"],
            "total": run_stats["total_tasks"],
            "ok": run_stats["ok_count"],
            "fail": run_stats["fail_count"],
            "output_file": output_file,
        }
        if not ok_df.empty:
            otp_payload["summary"] = summary.to_dict(orient="records")
        with open(otp_file, "w") as _f:
            _json.dump(otp_payload, _f, default=str)
    except Exception as _e:
        log.warning(f"Could not write status/OTP JSON: {_e}")


if __name__ == "__main__":
    main()
