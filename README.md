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

## Vercel Deployment Note

- This project cannot perform Word-to-PDF conversion inside Vercel Serverless Functions because LibreOffice is not available in that runtime.
- The app now returns a clear `503` error on Vercel instead of crashing.
- For real conversion hosting, deploy to an environment where LibreOffice can be installed (for example: Render, Railway, Fly.io, or your own VPS/Windows/Linux server).

## Fly.io Deployment

This repo now includes a `Dockerfile` that installs LibreOffice, so Fly.io can run the converter correctly.

1. Install Fly CLI:

```powershell
winget install Fly-io.flyctl
```

2. Login:

```bash
fly auth login
```

3. Launch the app from this project folder:

```bash
fly launch
```

When Fly asks questions:

- App name: choose any unique name
- Region: choose the nearest region
- PostgreSQL: `No`
- Redis: `No`
- Deploy now: `Yes`

4. If Fly asks for the internal port, use:

```text
3000
```

5. Open the deployed app:

```bash
fly open
```

6. Check logs if something fails:

```bash
fly logs
```

Useful commands:

```bash
fly status
fly deploy
fly logs
```

Health check endpoint:

```text
/healthz
```

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
