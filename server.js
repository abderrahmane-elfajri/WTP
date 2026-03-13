const express = require("express");
const path = require("path");
const fs = require("fs");
const os = require("os");
const fsp = require("fs/promises");
const multer = require("multer");
const archiver = require("archiver");
const libre = require("libreoffice-convert");
const { promisify } = require("util");

const libreConvertAsync = promisify(libre.convert);

const app = express();
const PORT = process.env.PORT || 3000;
const MAX_FILES = 100;
const IS_VERCEL = Boolean(process.env.VERCEL);

const writableRoot = IS_VERCEL ? os.tmpdir() : __dirname;
const uploadDir = path.join(writableRoot, "uploads");
const outputDir = path.join(writableRoot, "output");
const detectedSofficePath = detectSofficeBinary();

if (detectedSofficePath) {
  // libreoffice-convert checks this env var on Windows when PATH does not include soffice.
  process.env.LIBRE_OFFICE_EXE = detectedSofficePath;
  console.log(`Using LibreOffice binary: ${detectedSofficePath}`);
} else {
  console.warn("LibreOffice binary not found. Install LibreOffice to enable Word to PDF conversion.");
}

for (const dir of [uploadDir, outputDir]) {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => {
    const safeName = file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_");
    cb(null, `${Date.now()}-${Math.round(Math.random() * 1e9)}-${safeName}`);
  }
});

const upload = multer({
  storage,
  limits: {
    files: MAX_FILES,
    fileSize: 20 * 1024 * 1024
  },
  fileFilter: (req, file, cb) => {
    const allowed = [".doc", ".docx"];
    const ext = path.extname(file.originalname).toLowerCase();
    if (!allowed.includes(ext)) {
      return cb(new Error("Only .doc and .docx files are allowed"));
    }
    cb(null, true);
  }
});

app.use(express.static(path.join(__dirname, "public")));

app.get("/healthz", (req, res) => {
  res.status(200).json({ ok: true });
});

app.post("/api/convert", upload.array("wordFiles", MAX_FILES), async (req, res) => {
  if (IS_VERCEL) {
    return res.status(503).json({
      error: "This converter needs LibreOffice, which is not available in Vercel Serverless Functions. Deploy this app on a VPS/VM (Render, Railway, Fly.io, or your own server)."
    });
  }

  const files = req.files || [];

  if (!files.length) {
    return res.status(400).json({ error: "No files uploaded" });
  }

  if (files.length > MAX_FILES) {
    await cleanupFiles(files.map((f) => f.path));
    return res.status(400).json({ error: `Maximum ${MAX_FILES} files allowed` });
  }

  const createdPdfPaths = [];

  try {
    for (const file of files) {
      const inputBuffer = await fsp.readFile(file.path);
      const pdfBuffer = await libreConvertAsync(inputBuffer, ".pdf", undefined);

      const baseName = path.parse(file.originalname).name;
      const safeBaseName = baseName.replace(/[^a-zA-Z0-9._-]/g, "_");
      const pdfName = `${safeBaseName}.pdf`;
      const pdfPath = path.join(outputDir, `${Date.now()}-${Math.round(Math.random() * 1e9)}-${pdfName}`);

      await fsp.writeFile(pdfPath, pdfBuffer);
      createdPdfPaths.push({ name: pdfName, path: pdfPath });
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", "attachment; filename=converted-pdfs.zip");

    const archive = archiver("zip", { zlib: { level: 9 } });

    archive.on("error", async (err) => {
      await cleanupFiles(files.map((f) => f.path));
      await cleanupFiles(createdPdfPaths.map((f) => f.path));
      if (!res.headersSent) {
        res.status(500).json({ error: err.message });
      } else {
        res.end();
      }
    });

    archive.on("end", async () => {
      await cleanupFiles(files.map((f) => f.path));
      await cleanupFiles(createdPdfPaths.map((f) => f.path));
    });

    archive.pipe(res);

    for (const file of createdPdfPaths) {
      archive.file(file.path, { name: file.name });
    }

    await archive.finalize();
  } catch (err) {
    await cleanupFiles(files.map((f) => f.path));
    await cleanupFiles(createdPdfPaths.map((f) => f.path));

    const libreOfficeHint = err.message && err.message.toLowerCase().includes("soffice")
      ? " Install LibreOffice, then restart the server. On Windows, the expected binary is usually C:/Program Files/LibreOffice/program/soffice.exe."
      : "";

    return res.status(500).json({ error: `Conversion failed: ${err.message}.${libreOfficeHint}` });
  }
});

app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === "LIMIT_FILE_COUNT") {
      return res.status(400).json({ error: `Maximum ${MAX_FILES} files allowed` });
    }
    if (err.code === "LIMIT_FILE_SIZE") {
      return res.status(400).json({ error: "Each file must be 20MB or smaller" });
    }
    return res.status(400).json({ error: err.message });
  }

  if (err) {
    return res.status(400).json({ error: err.message });
  }

  next();
});

if (!IS_VERCEL) {
  app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

module.exports = app;

async function cleanupFiles(paths) {
  await Promise.all(
    paths.map(async (filePath) => {
      try {
        await fsp.unlink(filePath);
      } catch {
        // Ignore cleanup failures for files that already do not exist.
      }
    })
  );
}

function detectSofficeBinary() {
  const candidates = [
    process.env.LIBRE_OFFICE_EXE,
    path.join(process.env["PROGRAMFILES"] || "", "LibreOffice", "program", "soffice.exe"),
    path.join(process.env["PROGRAMFILES(X86)"] || "", "LibreOffice", "program", "soffice.exe"),
    "C:/Program Files/LibreOffice/program/soffice.exe",
    "C:/Program Files (x86)/LibreOffice/program/soffice.exe"
  ].filter(Boolean);

  for (const binaryPath of candidates) {
    if (fs.existsSync(binaryPath)) {
      return binaryPath;
    }
  }

  return "";
}
