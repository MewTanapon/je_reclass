"use strict";

const API_URL = "/api/run";

const inputFolder  = document.getElementById("inputFolder");
const outputFolder = document.getElementById("outputFolder");
const runBtn       = document.getElementById("runBtn");
const statusBox    = document.getElementById("statusBox");

// ── Status helpers ────────────────────────────────────────────────────────────

function showLoading() {
  statusBox.className = "status status--loading";
  statusBox.innerHTML = '<span class="spinner"></span>Processing… please wait.';
  statusBox.style.display = "";
}

function showSuccess(filename, outputPath) {
  statusBox.className = "status status--success";
  statusBox.innerHTML = `
    <strong>&#10003; Done!</strong><br/>
    File saved: <code>${escapeHtml(filename)}</code><br/>
    Path: <code>${escapeHtml(outputPath)}</code>
  `;
}

function showError(message) {
  statusBox.className = "status status--error";
  statusBox.innerHTML = `<strong>&#10007; Error:</strong> ${escapeHtml(message)}`;
}

function hideStatus() {
  statusBox.className = "status status--hidden";
}

// ── Validation ────────────────────────────────────────────────────────────────

function validate() {
  const inputVal  = inputFolder.value.trim();
  const outputVal = outputFolder.value.trim();

  if (!inputVal) {
    showError("Please enter the input folder path.");
    inputFolder.focus();
    return false;
  }
  if (!outputVal) {
    showError("Please enter the output folder path.");
    outputFolder.focus();
    return false;
  }
  return true;
}

// ── Main action ───────────────────────────────────────────────────────────────

async function runReclass() {
  hideStatus();

  if (!validate()) return;

  runBtn.disabled = true;
  showLoading();

  try {
    const response = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        input_folder:  inputFolder.value.trim(),
        output_folder: outputFolder.value.trim(),
      }),
    });

    const data = await response.json();

    if (response.ok && data.success) {
      showSuccess(data.output_filename, data.output_path);
    } else {
      showError(data.error || `Unexpected server response (HTTP ${response.status}).`);
    }
  } catch (err) {
    showError("Could not reach the server. Make sure the backend is running on port 5000.");
  } finally {
    runBtn.disabled = false;
  }
}

// ── Event bindings ────────────────────────────────────────────────────────────

runBtn.addEventListener("click", runReclass);

// Allow Enter key to trigger run from either input field
[inputFolder, outputFolder].forEach((el) => {
  el.addEventListener("keydown", (e) => {
    if (e.key === "Enter") runReclass();
  });
});

// ── Utilities ─────────────────────────────────────────────────────────────────

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
