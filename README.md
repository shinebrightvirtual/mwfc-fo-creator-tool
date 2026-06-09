# MWFC Fragrance Oil Listing Tool

Internal tool for Midwest Fragrance Co. — manages fragrance oil data and generates Shopify product listing HTML.

## Setup

1. Deploy to Vercel by connecting this GitHub repo
2. Open the deployed URL
3. Paste your Apps Script URL into the connection bar and click Connect
4. The tool will load all existing oils from your Google Sheet

## Apps Script

The `Code.gs` file must be deployed as a Google Apps Script Web App:
- Execute as: Me
- Who has access: Anyone

The Web App URL goes into the tool's connection bar.

## Files

- `index.html` — the full tool (styles + UI + logic, single file)
- `Code.gs` — Google Apps Script (data layer, sheet + Drive)
- `vercel.json` — Vercel routing config
