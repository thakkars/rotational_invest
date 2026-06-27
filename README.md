# Rotation Strategy — NiftyBees / GoldBees / SilverBees

This project auto-fetches daily price data via yfinance, updates the Excel workbook, and regenerates the HTML dashboard — all via GitHub Actions every weekday at 4:00 PM IST. The dashboard is deployed to Cloudflare Workers.

## Files
| File | Purpose |
|------|--------|
| `rotation_invest.py` | Main script: fetch data, update Excel, generate HTML |
| `worker/index.js` | Cloudflare Worker: `/api/trigger` + `/api/status` for the "Run Now" button |
| `wrangler.toml` | Cloudflare Workers config (serves `output/` as static assets) |
| `requirements.txt` | Python dependencies |
| `output/rotation_invest.xlsx` | Auto-updated Excel workbook (downloadable from dashboard) |
| `output/index.html` | Live dashboard |
| `.github/workflows/daily-refresh.yml` | GitHub Actions cron job (Mon–Fri 4 PM IST) + Wrangler deploy |

## GitHub Secrets (Actions)
| Secret | Value |
|--------|-------|
| `CLOUDFLARE_API_TOKEN` | Cloudflare API token ("Edit Cloudflare Workers" template) |
| `CLOUDFLARE_ACCOUNT_ID` | Your Cloudflare account ID |

## Cloudflare Worker Secrets
Set via Cloudflare dashboard → Workers & Pages → rotational-invest → Settings → Variables/Secrets, **or** via:
```bash
wrangler secret put GH_PAT
```
| Secret | Value |
|--------|-------|
| `GH_PAT` | GitHub PAT with `repo` + `workflow` scopes (used by the "Run Now" button) |

## How it works
1. GitHub Actions triggers at 10:30 UTC (4:00 PM IST), Mon–Fri
2. Python fetches latest NIFTYBEES, GOLDBEES, SILVERBEES data
3. Excel workbook is updated in `output/`
4. `index.html` dashboard is regenerated in `output/`
5. Files are committed back to `main` branch
6. Wrangler deploys `output/` to Cloudflare Workers
7. The dashboard's "Run Now" button triggers a manual run via the Worker

## Live site
**https://rotational-invest.thakkars912.workers.dev/**
