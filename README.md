# Excel to JSON Converter

A lightweight, client‑side web app to convert Excel (.xlsx, .xls) and CSV files to JSON. Drag and drop a file (or use the file picker), preview the JSON, choose which columns to include, and download or copy the result — all in your browser. No uploads, no backend.

> Data processing happens entirely in your browser. Your files never leave your device.

## Live demo

https://excel-json-converter.netlify.app/

## Features

- Drag & drop or file picker input
- Supports Excel (.xlsx, .xls) and CSV
- Live JSON preview with optional pretty print
- Column selector (choose which fields to keep)
- Toggle to include/exclude null/empty values
- Download JSON (UTF‑8 with BOM) or copy to clipboard
- Clean, responsive UI with Bootstrap 5

## Quick start

- Option 1: Just open `index.html` in your browser (double‑click on Windows/macOS)
- Option 2: Serve the folder with any static server (useful if your browser blocks local file access)

Once open:
1. Drag a spreadsheet onto the dashed area, or click to choose a file.
2. Use the Select Columns panel to include/exclude fields.
3. Toggle Pretty print and Show null/empty values as needed.
4. Download the JSON or Copy to Clipboard.

You can try it with the included sample: `intel-cpus.csv`.

## How it works

- Uses [SheetJS (xlsx)](https://sheetjs.com/) via CDN to parse Excel/CSV in the browser.
- Only the first worksheet is processed (for Excel files).
- The app expects headers in the first row. Keys are normalized to camelCase by default and non‑word symbols are removed.
- Values are cleaned:
  - `NaN`, `undefined`, and empty strings become `null`.
  - Strings are normalized (NFKC) and non‑ASCII characters are stripped except common symbols like ®, ™, ©, ±, µ.
- Privacy by design: no network requests are made for your data; everything happens locally.

## UI overview

- Drag & Drop zone and file input
- Options:
  - Pretty print JSON
  - Show null/empty values
- Column selection (at least one column must remain selected)
- Live preview panel
- Actions: Download JSON, Copy to Clipboard, Clear

## Files in this repo

- `index.html` — Main UI (Bootstrap 5 + CDN for SheetJS)
- `app.js` — Core logic: file handling, parsing, cleaning, preview, column filtering, download/copy
- `excelProcessor.js` — Web Worker version of the parsing/cleaning logic (not currently wired in the UI). You can integrate it to keep parsing off the main thread for very large files.
- `intel-cpus.csv` — Sample dataset for testing
- `icon.png` — Favicon

## Development notes

- No build step required; it’s a static site. Open `index.html` directly or serve the folder.
- If you enable the worker (`excelProcessor.js`), make sure to wire it from `app.js` and host the files via an HTTP server (some browsers restrict workers on the `file://` scheme).

## Known limitations

- Only the first worksheet is processed.
- Expects headers in the first row.
- Very large files may be slow in the main thread; use the worker approach for better responsiveness.
- Column names in the JSON reflect the cleaned/camelCased headers, not the original display headers.

## Deploying to GitHub Pages

1. Push this repository to GitHub.
2. In your repository settings, enable GitHub Pages (deploy from the `main` branch, `/` root).
3. Your app will be served at `https://<your-username>.github.io/<repo-name>/`.

## Acknowledgments

- [SheetJS/xlsx](https://github.com/SheetJS/sheetjs) for in‑browser Excel/CSV parsing
- [Bootstrap 5](https://getbootstrap.com/) for styling

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.
