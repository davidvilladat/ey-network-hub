"""
FR24 OTP runner — drives Master Webscrapping CQ.py via subprocess so stdin
piping works correctly on Windows (avoids PowerShell blank-line stripping).

Usage:
  python _run_fr24_test.py              # 7-day rolling window, all flights
  python _run_fr24_test.py 1            # yesterday only
  python _run_fr24_test.py 7            # last 7 days (default)
  python _run_fr24_test.py 30           # last 30 days (FR24 history limit ~30d)
  python _run_fr24_test.py 7 50         # last 7 days, max 50 flights per day
"""
import subprocess, os, sys

days_back = int(sys.argv[1]) if len(sys.argv) > 1 else 7
limit     = sys.argv[2]     if len(sys.argv) > 2 else ""   # empty = ALL

from datetime import date, timedelta
ref_date = (date.today() - timedelta(days=1)).isoformat()   # yesterday

# Inputs answered in order by the scraper prompts:
#   1. Reference date       → yesterday
#   2. Days back            → days_back - 1  (0 = ref_date only, N-1 = N days)
#   3. Workers              → 1
#   4. Flights per day      → limit (empty = ALL)
#   5. Data source          → 2 (FR24)
days_arg = str(days_back - 1)
inputs = f"{ref_date}\n{days_arg}\n1\n{limit}\n2\n".encode()

env = os.environ.copy()
env["FR24_API_KEY"] = (
    "019d23e6-8df3-72f5-9114-8010e73a32fe"
    "|OlAU6CrTT9A8UDJcHbhs8JKYCW2bQLLqEegvEewJbf109a85"
)

subprocess.run(
    [sys.executable, "-X", "utf8", "Master Webscrapping CQ.py"],
    input=inputs,
    cwd=os.path.dirname(os.path.abspath(__file__)),
    env=env,
)
