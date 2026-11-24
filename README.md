# Class X Student Data Extractor

A browser-only app for extracting Class X records from CSV, Excel, ZIP, and RAR files. Files are read locally; student data is not uploaded to an app server.

## Use

1. Select one or more CSV, Excel, ZIP, or RAR files.
2. Choose one combined Excel workbook or separate Excel files bundled in a ZIP.
3. Select **Extract Class X data**, review the preview, and download the result.

ZIP files can contain CSV, Excel, and RAR files. RAR files can contain CSV, Excel, and ZIP files. Encrypted or damaged archives may fail to extract. The first non-empty worksheet in each Excel file is used.

Rows match when `classId` is `10` or `X`, or `classDesc` is `X`, `10`, `10th`, `Class X`, `Class 10`, `Class 10th`, `Grade X`, `Grade 10`, or `Grade 10th` (case-insensitive). The combined workbook adds a `Source File` column.

## Run locally

```sh
npm install
npm run dev
```

## Deploy to Vercel

Import this project directory into Vercel. Vercel uses `npm run build` and serves the generated `dist` directory as configured in `vercel.json`. No environment variables or backend services are required. The build copies the RAR WebAssembly runtime into the static output. Node.js 20.19 or newer is required.

The RAR runtime is provided by [`@unrar-browser/core`](https://github.com/DavidKk/unrar-browser-wasm), which uses the UnRAR library and has its own license terms; review those terms if redistributing the app.
