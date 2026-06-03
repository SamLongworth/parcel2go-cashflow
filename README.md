# Parcel2Go Cash Flow Intelligence

A live AI-powered cash flow forecasting dashboard for Parcel2Go.

## Features

- **Live data connectors** — Stripe, Xero, TrueLayer (UK Open Banking)
- **Daily forecast** — projects cash balance to month-end, day by day
- **Scenario Modeller** — drag sliders to stress-test receipt delays and bad debt
- **AI CFO Briefing** — daily management summary powered by Claude AI
- **Cash Flow Assistant** — conversational Q&A over your live data
- **Export** — one-click Excel (multi-sheet) and CSV download

---

## Deploy to Vercel (share a live link in 5 minutes)

### Step 1 — Push to GitHub

```bash
# In this folder:
git init
git add .
git commit -m "Parcel2Go Cash Flow Dashboard"
git branch -M main

# Create a new repo at github.com, then:
git remote add origin https://github.com/YOUR_USERNAME/parcel2go-cashflow.git
git push -u origin main
```

### Step 2 — Deploy on Vercel

1. Go to **vercel.com** and sign in with GitHub (free)
2. Click **"Add New Project"**
3. Import your `parcel2go-cashflow` repo
4. Vercel auto-detects Vite — click **Deploy**
5. In ~60 seconds you get a live URL: `parcel2go-cashflow.vercel.app`

That's it. Share the link. It works on any device, any browser.

---

## Run locally

```bash
npm install
npm run dev
# Opens at http://localhost:5173
```

---

## Connecting live data (optional)

All keys are entered in the dashboard UI under **"Connect Live Data"**.  
They live in browser memory only — never sent anywhere except the respective API.

### Stripe
1. Go to `dashboard.stripe.com` → Developers → API keys
2. Create a **Restricted Key** with read access to: Payment Intents, Payouts, Balance
3. Paste the `sk_live_...` key in the dashboard

### Xero
1. Register at `developer.xero.com` → New App → Web App
2. Complete the OAuth 2.0 flow to get an **Access Token**
3. Find your **Tenant ID** in the Xero Connections page
4. Paste both into the dashboard
> Note: Xero tokens expire after 30 minutes. For production use, add a backend token-refresh proxy.

### Bank Feed (TrueLayer — UK Open Banking)
1. Register free at `console.truelayer.com`
2. Create an app and connect your bank (Barclays, HSBC, Lloyds, NatWest, Monzo, Starling etc.)
3. Complete the bank OAuth flow to get an **Access Token**
4. Paste it into the dashboard
> Note: TrueLayer requires a backend proxy for production CORS compliance.

---

## Project structure

```
parcel2go/
├── index.html          # Entry point
├── vite.config.js      # Vite config
├── package.json        # Dependencies
├── src/
│   ├── main.jsx        # React root
│   └── App.jsx         # Full dashboard (connectors, forecast, AI, export)
└── README.md
```

---

## Tech stack

| Layer | Technology |
|---|---|
| UI framework | React 18 |
| Charts | Recharts |
| Export | SheetJS (xlsx) |
| Build | Vite |
| Hosting | Vercel (free) |
| AI commentary | Claude API (falls back to rule-based logic) |
| Stripe connector | Stripe REST API |
| Xero connector | Xero Accounting API (OAuth 2.0) |
| Bank connector | TrueLayer Open Banking API |
