# Website Demo

This folder contains a standalone website demo for `KPI Card`.

It is intentionally separate from the Tableau extension runtime and mirrors the split-card view:

- `Sales`
- `Month of Order Date`
- split by `Ship Mode`
- filters kept for `Region`, `Segment`, and `Category`

## Files

- `index.html`: Demo page
- `styles.css`: Demo styling
- `app.js`: Browser logic
- `sample_-_superstore.xls`: Sample data source

## Publish

Upload the contents of this folder to a static host such as GitHub Pages.

Important:

- Keep `sample_-_superstore.xls` in the same folder as `index.html`
- The page reads the workbook directly in the browser
- Internet access is required for the Google Fonts and SheetJS CDN scripts

## Demo Behavior

- Filters: `Region`, `Segment`, `Category`
- One card per `Ship Mode`
- Each card shows KPI Summary + Bar Chart
- Hover a bar inside a card to preview that month
