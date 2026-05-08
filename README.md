# Etihad Crew · Network Analysis Hub

Static site hosted on Cloudflare Pages at **cristianvillada.com** (or subdomain).  
Built for Etihad crew — ZED travel intelligence and network strategy tools.

## Tools

| Page | Description |
|------|-------------|
| `/` | Landing hub — links to both tools |
| `/zed-atlas` | ZED Atlas v4 — 218 agreements, interactive world map, destination finder |
| `/ey-analysis` | EY Network Analysis — route scatter, 2025/2026 expansion map |

## Data files (`data/`)

| File | Description | Updated |
|------|-------------|---------|
| `zed_agreements.json` | ZED fare matrix — 218 carriers | 30 Apr 2026 |
| `networks.json` | FR24 route/fleet/schedule — 38 carriers | 1–3 May 2026 |
| `airports.json` | Coordinates — 345 airports | 30 Apr 2026 |
| `world_path.txt` | World map SVG geometry | static |
| `icao_map.json` | Airline ICAO→IATA mapping | 30 Apr 2026 |

## Updating data

Replace any file in `data/` and push to `main` — Cloudflare Pages deploys automatically within ~30 seconds. The HTML files never need to change.

## Local development

```bash
python3 -m http.server 8080
# open http://localhost:8080
```

## Deployment

Cloudflare Pages — connected to this GitHub repo, branch `main`, no build step, output directory `/`.

## Disclaimer

For crew use only. Always verify fares in OSS / IDTGR before purchasing.  
FR24 data is a sample (~20 flights/carrier/day) and may not reflect full schedules.
