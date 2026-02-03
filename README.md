# KCI Upload Creation Tool

The KCI Upload Creation Tool is a browser-only utility for preparing KCI repair and closed case outputs from Excel and CSV inputs. It runs entirely in the browser and is designed to work on GitHub Pages with no backend or build steps.

Use the app to load your KCI Excel workbook and supporting CSV files, process repair and closed cases, and export the resulting datasets as CSV downloads. All processing happens locally in your browser, so your data never leaves your machine.

## Live Site

https://<username>.github.io/kci-upload-creation-tool/

Enable GitHub Pages with the `/docs` folder as the site source.

## How to Use

1. Upload the KCI Excel workbook.
2. Upload the CSO CSV file.
3. Upload the Tracking CSV file (optional, used for tracking URL exports).
4. Click **Process Repair Cases**.
5. Click **Process Closed Cases**.
6. Use the **Export** buttons to download outputs.

## Supported Browsers

- Latest Chrome
- Latest Edge
- Latest Firefox
- Latest Safari

## Limitations

- Browser-only processing means very large files may be slow or memory intensive.
- The tool does not validate business rules beyond the existing processing logic.

## Privacy & Data Handling

This app is fully client-side. No backend services are used, and your data never leaves your browser.
