# Rotation Strategy — NiftyBees / GoldBees / SilverBees

This project auto-fetches daily price data from TradingView, updates the Excel workbook, and regenerates the HTML dashboard — all via GitHub Actions every weekday at 4:00 PM IST.

## Files
| File | Purpose |
|------|--------|
| `rotation_invest.py` | Main script: fetch data, update Excel, generate HTML |
| `requirements.txt` | Python dependencies |
| `output/rotation_invest.xlsx` | Auto-updated Excel workbook |
| `output/index.html` | Live dashboard (served via Cloudflare Pages) |
| `.github/workflows/daily-refresh.yml` | GitHub Actions cron job (Mon–Fri 4 PM IST) |

## GitHub Secrets Required
| Secret | Value |
|--------|-------|
| `TV_USERNAME` | TradingView username |
| `TV_PASSWORD` | TradingView password |

## How it works
1. GitHub Actions triggers at 10:30 UTC (4:00 PM IST), Mon–Fri
2. Python fetches latest NIFTYBEES, GOLDBEES, SILVERBEES data
3. Excel workbook is updated in `output/`
4. `index.html` dashboard is regenerated in `output/`
5. Files are committed back to `main` branch
6. Cloudflare Pages serves the latest `output/index.html` automatically
