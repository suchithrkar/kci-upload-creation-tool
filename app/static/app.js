const outputEl = document.getElementById("output");
const statusEl = document.getElementById("status");

const kciInput = document.getElementById("kciExcel");
const csoInput = document.getElementById("csoCsv");
const trackingInput = document.getElementById("trackingCsv");

const uploadKciBtn = document.getElementById("uploadKci");
const uploadCsoBtn = document.getElementById("uploadCso");
const uploadTrackingBtn = document.getElementById("uploadTracking");
const processRepairBtn = document.getElementById("processRepair");
const processClosedBtn = document.getElementById("processClosed");
const copySoBtn = document.getElementById("copySo");
const copyTrackingBtn = document.getElementById("copyTracking");

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "crimson" : "inherit";
}

function setOutput(content) {
  if (typeof content === "string") {
    outputEl.textContent = content;
    return;
  }
  outputEl.textContent = JSON.stringify(content, null, 2);
}

async function uploadFile(url, file) {
  if (!file) {
    setStatus("Please choose a file before uploading.", true);
    return;
  }
  const formData = new FormData();
  formData.append("file", file);

  setStatus("Uploading...");
  try {
    const response = await fetch(url, { method: "POST", body: formData });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.detail || "Upload failed.");
    }
    setOutput(data);
    setStatus("Upload complete.");
  } catch (error) {
    setStatus(error.message, true);
  }
}

async function callJsonEndpoint(url) {
  setStatus("Processing...");
  try {
    const response = await fetch(url, { method: "POST" });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.detail || "Request failed.");
    }
    setOutput(data);
    setStatus("Done.");
  } catch (error) {
    setStatus(error.message, true);
  }
}

async function callTextEndpoint(url) {
  setStatus("Processing...");
  try {
    const response = await fetch(url);
    const data = await response.text();
    if (!response.ok) {
      throw new Error(data || "Request failed.");
    }
    setOutput(data);
    setStatus("Done.");
  } catch (error) {
    setStatus(error.message, true);
  }
}

uploadKciBtn.addEventListener("click", () => {
  uploadFile("/upload/kci-excel", kciInput.files[0]);
});

uploadCsoBtn.addEventListener("click", () => {
  uploadFile("/upload/cso-csv", csoInput.files[0]);
});

uploadTrackingBtn.addEventListener("click", () => {
  uploadFile("/upload/tracking-csv", trackingInput.files[0]);
});

processRepairBtn.addEventListener("click", () => {
  callJsonEndpoint("/process/repair-cases");
});

processClosedBtn.addEventListener("click", () => {
  callJsonEndpoint("/process/closed-cases");
});

copySoBtn.addEventListener("click", () => {
  callTextEndpoint("/process/copy-so");
});

copyTrackingBtn.addEventListener("click", () => {
  callTextEndpoint("/process/copy-tracking");
});
