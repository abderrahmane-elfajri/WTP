# Word to PDF Batch Web Tool

This project is a website tool that converts multiple Word files to PDF.

## Features

- Upload `.doc` and `.docx` files
- Convert files to PDF on the server
- Download all converted PDFs as one ZIP file
- Maximum selection limit: **100 files**
- File size limit: **20MB per file**

## Requirements

- Node.js 18+
- LibreOffice installed on your machine
  - Windows: install LibreOffice from the official installer
  - Make sure `soffice` is available in system PATH

## Run locally

```bash
npm install
npm start
```

Then open:

`http://localhost:3000`

## Notes

- Conversion quality depends on LibreOffice.
- If conversion fails with a `soffice` error, check LibreOffice installation and PATH.

## Windows Troubleshooting (`Could not find soffice binary`)

1. Install LibreOffice:

```powershell
winget install --id TheDocumentFoundation.LibreOffice -e
```

2. Verify the binary exists:

```powershell
Test-Path "C:\Program Files\LibreOffice\program\soffice.exe"
```

3. Restart the app:

```bash
npm start
```

You can also set this env var if LibreOffice is installed in a custom location:

```powershell
$env:LIBRE_OFFICE_EXE="C:\path\to\LibreOffice\program\soffice.exe"
npm start
```
