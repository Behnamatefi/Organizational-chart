# Organizational Chart Explorer

A lightweight, client-side web app for visualizing an organization's structure from an Excel workbook or Google Sheets link. Switch between tribe-based and manager-based views, search across employee attributes, and navigate large hierarchies with a built-in mini map and zoom controls.

## Getting started

1. Open `index.html` in your browser. No build tools are required.
2. Upload an `.xlsx` file or paste a Google Sheets share link with columns for:
   - Employee name
   - Tribe
   - Squad
   - Expertise
   - Job level
   - Manager name
3. Use the view toggle to switch between tribe and manager views.
4. Use the search bar to filter by any employee attribute.
5. Explore the chart with the floating controls (expand/collapse, zoom, and zoom-to-fit) or the minimap.

A curated sample dataset is loaded automatically so you can try the interface immediately. Click **Use sample data** to restore the sample if needed.

## Development notes

- All processing happens in the browser using [SheetJS](https://sheetjs.com/) via CDN to parse Excel files.
- Google Sheets links are converted to CSV via the public export endpoint. Ensure the sheet is shared accordingly.
- The layout is computed on the fly with a simple tree layout algorithm; no external frameworks are required.

Feel free to customize the styling in `styles.css` or extend the rendering logic in `app.js` to suit your organization.
