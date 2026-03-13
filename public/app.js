const form = document.getElementById("convertForm");
const fileInput = document.getElementById("wordFiles");
const fileCount = document.getElementById("fileCount");
const statusEl = document.getElementById("status");
const submitBtn = document.getElementById("submitBtn");

const MAX_FILES = 100;

fileInput.addEventListener("change", () => {
  const selectedCount = fileInput.files.length;

  if (!selectedCount) {
    fileCount.textContent = "No files selected";
    statusEl.textContent = "";
    statusEl.className = "status";
    return;
  }

  fileCount.textContent = `${selectedCount} file(s) selected`;

  if (selectedCount > MAX_FILES) {
    statusEl.textContent = `You can select maximum ${MAX_FILES} files.`;
    statusEl.className = "status error";
  } else {
    statusEl.textContent = "";
    statusEl.className = "status";
  }
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!fileInput.files.length) {
    setStatus("Please choose at least one Word file.", "error");
    return;
  }

  if (fileInput.files.length > MAX_FILES) {
    setStatus(`You can select maximum ${MAX_FILES} files.`, "error");
    return;
  }

  submitBtn.disabled = true;
  setStatus("Converting files... Please wait.", "");

  try {
    const formData = new FormData();

    for (const file of fileInput.files) {
      formData.append("wordFiles", file);
    }

    const response = await fetch("/api/convert", {
      method: "POST",
      body: formData
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error || "Conversion failed");
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "converted-pdfs.zip";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    setStatus("Done. Your ZIP file has been downloaded.", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    submitBtn.disabled = false;
  }
});

function setStatus(message, type) {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`.trim();
}
