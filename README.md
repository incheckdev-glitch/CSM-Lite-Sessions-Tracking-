# InCheck Completion Dashboard — Static Standalone Version

A standalone, responsive completion dashboard based on the supplied InCheck Completion Report. It does **not** connect to the ERP or Supabase. All data is stored locally in `data.js`.

## Run locally

Open `index.html` directly in a browser, or use a simple local server:

```bash
python -m http.server 8080
```

Then open `http://localhost:8080`.

## Publish with GitHub Pages

1. Create a new GitHub repository.
2. Upload all files from this folder to the repository root.
3. Open **Settings → Pages**.
4. Select **Deploy from a branch**.
5. Choose the `main` branch and `/root` folder.
6. Save. GitHub will generate a public dashboard URL.

## Included features

- Responsive professional dashboard
- Weekly and monthly static periods
- Client, brand, status, period, and location filters
- KPI cards and completion charts
- Brand performance drill-down
- Historical trend chart
- Search, sorting, and pagination
- CSV export
- Print / Save as PDF with A4 landscape styling
- Dark mode
- No external libraries or build step

## Edit data

Open `data.js` and update:

- `periods`
- `brands`
- `locations`

The dashboard recalculates all averages automatically.
