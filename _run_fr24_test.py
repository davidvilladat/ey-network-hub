import subprocess, os, sys

inputs = b"2026-05-17\n0\n1\n\n2\n"
env = os.environ.copy()
env["FR24_API_KEY"] = "019d23e6-8df3-72f5-9114-8010e73a32fe|OlAU6CrTT9A8UDJcHbhs8JKYCW2bQLLqEegvEewJbf109a85"

subprocess.run(
    [sys.executable, "-X", "utf8", "Master Webscrapping CQ.py"],
    input=inputs,
    cwd=os.path.dirname(os.path.abspath(__file__)),
    env=env,
)
