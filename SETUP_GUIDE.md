# Rotation Signal Bot — Setup Guide
**Zero cost. Runs daily at 4 PM IST. Sends Telegram alerts.**

---

## Files in This Repo

```
rotational_invest/
├── signal_bot.py
├── requirements.txt
├── SETUP_GUIDE.md
└── .github/
    └── workflows/
        └── signal_bot.yml
```

---

## Step 1 — Create Telegram Bot (5 min)

1. Open Telegram → search `@BotFather` → `/newbot`
2. Give it a name (e.g. `Rotation Signal Bot`) and username ending in `bot`
3. Copy the **Bot Token** (looks like `7123456789:AAF...`)
4. Start a chat with your bot → send `/start`
5. Get your **Chat ID** by opening:
   `https://api.telegram.org/bot<YOUR_TOKEN>/getUpdates`
   Look for `"chat":{"id":XXXXXXX}` — that number is your Chat ID

---

## Step 2 — Create Cloudflare R2 Bucket (5 min)

1. Log in to [dash.cloudflare.com](https://dash.cloudflare.com)
2. Left sidebar → **R2 Object Storage** → **Create bucket**
3. Name: `rotation-signals` → Create
4. Go to **R2 Overview** → **Manage R2 API Tokens**
5. Click **Create API Token** → select **Object Read & Write**
6. Copy: Access Key ID, Secret Access Key, Endpoint URL
   (format: `https://<ACCOUNT_ID>.r2.cloudflarestorage.com`)

---

## Step 3 — Add GitHub Secrets (5 min)

In this repo → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

| Secret Name        | Value                                           |
|--------------------|-------------------------------------------------|
| `TELEGRAM_TOKEN`   | Your Telegram bot token                         |
| `TELEGRAM_CHAT_ID` | Your Telegram chat ID (number)                  |
| `R2_ACCESS_KEY`    | Cloudflare R2 Access Key ID                     |
| `R2_SECRET_KEY`    | Cloudflare R2 Secret Access Key                 |
| `R2_ENDPOINT`      | `https://<ACCOUNT_ID>.r2.cloudflarestorage.com` |
| `R2_BUCKET`        | `rotation-signals`                              |

---

## Step 4 — Test It Now

1. Go to **Actions** tab in this repo
2. Click **Rotation Signal Bot** → **Run workflow** → **Run workflow**
3. Watch the logs — completes in ~30 seconds
4. Check Telegram — you should receive the first message!

---

## Daily Telegram Message Examples

**Normal day (no signal):**
```
📊 Daily Rotation Update

Prices (02-Jun-2026 04:00 PM IST)
  NiftyBees : ₹265.40
  GoldBees  : ₹124.61
  SilverBees: ₹68.30

Signal Status
  ⏸ S1: HOLD | Ratio: 2.1270
     Position → GOLDBEES
  ⏸ S2: HOLD | Ratio: 3.8870
     Position → NIFTYBEES
  ⏸ S3: HOLD | Ratio: 1.8240
     Position → GOLDBEES
```

**When a signal fires:**
```
🚨 NEW SIGNAL FIRED!

Prices (23-Mar-2026 04:00 PM IST)
  NiftyBees : ₹255.00
  GoldBees  : ₹110.72

Signal Status
  🟢 S1: BUY NIFTYBEES | Ratio: 2.3045
     Position → NIFTYBEES

⚡ Execute switch at market open tomorrow (NSE 9:15 AM)
```

---

## Free Tier Limits

| Service         | Free Limit       | Your Usage     |
|-----------------|------------------|----------------|
| GitHub Actions  | 2,000 min/month  | ~50 min/month  |
| Cloudflare R2   | 10 GB storage    | < 1 MB/month   |
| Telegram Bot    | Unlimited        | 1 msg/day      |
| Yahoo Finance   | Unlimited        | 3 tickers/day  |

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Bot token invalid | Re-copy from BotFather, no spaces |
| Chat ID not found | Make sure you started a chat with the bot first |
| R2 endpoint error | Use full URL with `https://` prefix |
| No data for ticker | NSE may be closed (holiday) — bot shows N/A and skips |
| Workflow not triggering | Use 'Run workflow' button in Actions tab to test |
