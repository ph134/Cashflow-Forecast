# Cashflow Forecast Web App

This is a dependency-free browser rebuild of the Excel cashflow tool from `Cashflow Template_2026_Feb 20.xlsm`.

## What it includes

- Contract and timeline inputs
- Revenue milestones with percent and payment month
- Cost elements with cumulative monthly completion percentages
- Monthly revenue, cumulative revenue, monthly cost, cumulative cost, net cash flow, and cumulative net
- Net cash flow chart rendered in SVG

## How to run

Open `index.html` in a browser.

If you prefer serving it from a local web server, PowerShell includes an easy option when Python is installed:

```powershell
Set-Location "C:\Users\405795\cashflow-web-app"
python -m http.server 8000
```

Then open `http://localhost:8000`.

## Share Link + Keep Main Editable

Use a separate publish branch so users get a stable URL while you keep developing in your main branch.

1. Keep daily edits in `main`.
2. Publish only approved updates to `live-share`.
3. Configure GitHub Pages to serve from branch `live-share` and folder `/ (root)`.
4. Share the Pages URL with users.

### One-command publish

Run this from the repo root in PowerShell:

```powershell
.\scripts\publish-live.ps1
```

Optional parameters:

```powershell
.\scripts\publish-live.ps1 -FromBranch main -LiveBranch live-share -Remote origin
```

What the script does:

- Verifies clean working tree
- Pulls latest `main`
- Creates/updates `live-share`
- Fast-forward merges from `main`
- Pushes `live-share`

After first push, enable GitHub Pages:

1. GitHub repo -> Settings -> Pages
2. Source: Deploy from a branch
3. Branch: `live-share`, Folder: `/ (root)`
4. Save and share the generated URL

## Workbook logic mapped

- Timeline months = quoted lead time + `ROUND(net days / 30, 0)`
- Revenue per milestone = contract value x milestone percent x exchange rate
- Revenue is placed into the month matching the milestone payment month
- Cost schedule uses cumulative completion percentages by month for each cost element
- Monthly cost is the incremental change in cumulative completion value
- Net cash flow = monthly revenue - monthly cost
