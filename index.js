const TABLE_SCHEMAS = {
  "Dump": [
    "Case ID",
    "Full Name (Primary Contact) (Contact)",
    "Created On",
    "Created By",
    "Modified On",
    "Business Segment",
    "Country",
    "Incoming Channel",
    "Case Resolution Code",
    "Full Name (Owning User) (User)",
    "Email Status",
    "Case Ready for Closure",
    "Requires Auto-Close",
    "Is Order Created",
    "Queue",
    "OTC Code",
    "Product Serial Number",
    "ProductName"
  ],
  "WO": [
    "Case ID",
    "Work Order Number",
    "Business Segment",
    "Service Account",
    "Case Status (Case ID) (Case)",
    "System Status",
    "Created On",
    "Workgroup (Created By) (User)",
    "Country",
    "Resolution Notes/ Diagnostics"
  ],
  "MO": [
    "Order Number",
    "Case ID",
    "Created On",
    "Order Status",
    "Order Type",
    "Ready For Closure Date"
  ],
  "MO Items": [
    "Material Order",
    "MO Line Items Name",
    "Part Number",
    "Description",
    "Tracking Number",
    "Created On",
    "Tracking Url"
  ],
  "SO": [
    "Case ID",
    "Service Status",
    "Date and Time Submitted",
    "Name",
    "Order Reference ID"
  ],
  "Closed Cases": [
    "Case ID",
    "Created On",
    "Modified By",
    "Modified On",
    "Case Closed Date",
    "Case Resolution Code",
    "Created By",
    "Incoming Channel",
    "Owner",
    "Country",
    "OTC Code"
  ],
  "CSO Status": [
    "Case ID",
    "CSO",
    "Status",
    "Tracking Number",
    "Repair Status"
  ],
  "Delivery Details": [
    "CaseID",
    "CurrentStatus"
  ],
  "Repair Cases": [
    "Case ID",
    "Customer Name",
    "Created On",
    "Created By",
    "Country",
    "Case Resolution Code",
    "Case Owner",
    "OTC Code",
    "CA Group",
    "TL",
    "SBD",
    "Onsite RFC",
    "CSR RFC",
    "Bench RFC",
    "Market",
    "WO Closure Notes",
    "Tracking Status",
    "Part Number",
    "Part Name",
    "Serial Number",
    "Product Name",
    "Email Status",
    "DNAP"
  ],
  "Closed Cases Report": [
    "Case ID",
    "Customer Name",
    "Created On",
    "Created By",
    "Modified By",
    "Modified On",
    "Case Closed Date",
    "Closed By",
    "Country",
    "Case Resolution Code",
    "Case Owner",
    "OTC Code",
    "TL",
    "SBD",
    "Market"
  ]
};

const DUMP_HEADER_DISPLAY_MAP = {
  "Full Name (Primary Contact) (Contact)": "Customer Name",
  "Full Name (Owning User) (User)": "Case Owner",
  "ProductName": "Product Name",
  "Product Serial Number": "Serial Number"
};

const DB_NAME = "KCI_CASE_TRACKER_DB";
const DB_VERSION = 2;
const STORE_NAME = "sheets";

const dataTablesMap = {};
const tablesMap = {};

let db = null;
let selectedFiles = {
  kci: null,
  cso: null,
  tracking: null
};

const selectors = {
  kciInput: document.getElementById("kciInput"),
  csoInput: document.getElementById("csoInput"),
  trackingInput: document.getElementById("trackingInput"),
  processBtn: document.getElementById("processBtn"),
  tablesContainer: document.getElementById("tablesContainer"),
  soModal: document.getElementById("soModal"),
  soOutput: document.getElementById("soOutput"),
  soCount: document.getElementById("soCount"),
  soModalTitle: document.getElementById("soModalTitle"),
  overlay: document.getElementById("progressOverlay"),
  overlayStatus: document.getElementById("overlayStatusText"),
  overlayProgress: document.getElementById("overlayProgressBar"),
  overlayProgressText: document.getElementById("overlayProgressText"),
  overlayConfirm: document.getElementById("overlayConfirmBtn")
};

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = event => {
      const database = event.target.result;
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        database.createObjectStore(STORE_NAME, { keyPath: "sheetName" });
      }
    };

    request.onsuccess = event => {
      db = event.target.result;
      resolve(db);
    };

    request.onerror = () => reject("Failed to open IndexedDB");
  });
}

function getStore(mode = "readonly") {
  const tx = db.transaction(STORE_NAME, mode);
  return tx.objectStore(STORE_NAME);
}

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function cleanCell(value) {
  if (value === null || value === undefined) return "";
  return String(value)
    .replace(/\u00A0/g, " ")
    .replace(/[\u0000-\u001F\u007F]/g, "")
    .trim();
}

function excelDateToJSDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);

  const fractionalDay = serial - Math.floor(serial) + 0.0000001;
  let totalSeconds = Math.floor(86400 * fractionalDay);

  const seconds = totalSeconds % 60;
  totalSeconds -= seconds;
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor(totalSeconds / 60) % 60;

  dateInfo.setHours(hours);
  dateInfo.setMinutes(minutes);
  dateInfo.setSeconds(seconds);

  return dateInfo;
}

function formatDate(date) {
  const pad = number => String(number).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function normalizeRowToSchema(row, sheetName) {
  const headers = TABLE_SCHEMAS[sheetName];
  const normalized = new Array(headers.length).fill("");

  for (let i = 0; i < headers.length; i += 1) {
    if (row && row[i] !== undefined) {
      normalized[i] = row[i];
    }
  }

  return normalized;
}

function getColumnIndex(sheetName, columnName) {
  const headers = TABLE_SCHEMAS[sheetName];
  return headers ? headers.indexOf(columnName) + 1 : -1;
}

function attachSerialNumber(dt) {
  dt.on("order.dt search.dt draw.dt", () => {
    dt.column(0, { search: "applied", order: "applied" })
      .nodes()
      .each((cell, i) => {
        cell.innerHTML = i + 1;
      });
  });
}

function initEmptyTables() {
  selectors.tablesContainer.innerHTML = "";

  const tabs = document.createElement("div");
  tabs.className = "sheet-tabs";

  const left = document.createElement("div");
  left.className = "sheet-tabs-left";

  const right = document.createElement("div");
  right.className = "sheet-tabs-right";

  tabs.appendChild(left);
  tabs.appendChild(right);
  selectors.tablesContainer.appendChild(tabs);

  let first = true;

  Object.keys(TABLE_SCHEMAS).forEach(sheetName => {
    const headers = TABLE_SCHEMAS[sheetName];

    const tableWrapper = document.createElement("div");
    tableWrapper.className = "table-scroll-wrapper";
    tableWrapper.style.display = first ? "block" : "none";
    tableWrapper.dataset.sheet = sheetName;

    const table = document.createElement("table");
    table.className = "display";

    const thead = document.createElement("thead");
    const tr = document.createElement("tr");

    const snTh = document.createElement("th");
    snTh.textContent = "S.No";
    tr.appendChild(snTh);

    headers.forEach(header => {
      const th = document.createElement("th");
      if (sheetName === "Dump" && DUMP_HEADER_DISPLAY_MAP[header]) {
        th.textContent = DUMP_HEADER_DISPLAY_MAP[header];
      } else {
        th.textContent = header;
      }
      tr.appendChild(th);
    });

    thead.appendChild(tr);
    table.appendChild(thead);
    table.appendChild(document.createElement("tbody"));

    tableWrapper.appendChild(table);
    selectors.tablesContainer.appendChild(tableWrapper);

    const columnDefs = [
      {
        targets: 0,
        searchable: false,
        orderable: false
      }
    ];

    if (sheetName === "WO") {
      const idx = getColumnIndex(sheetName, "Resolution Notes/ Diagnostics");
      if (idx !== -1) {
        columnDefs.push({
          targets: idx,
          width: "300px",
          className: "fixed-notes-col"
        });
      }
    }

    if (sheetName === "Repair Cases") {
      const idx = getColumnIndex(sheetName, "WO Closure Notes");
      if (idx !== -1) {
        columnDefs.push({
          targets: idx,
          width: "300px",
          className: "fixed-notes-col"
        });
      }
    }

    const dt = $(table).DataTable({
      pageLength: 10,
      autoWidth: false,
      columnDefs,
      order: [[1, "asc"]],
      fixedHeader: {
        header: true,
        headerOffset: 92
      },
      dom: "<'dt-top'l f><'dt-middle't><'dt-bottom'i p>"
    });

    attachSerialNumber(dt);
    dataTablesMap[sheetName] = dt;
    tablesMap[sheetName] = tableWrapper;

    const tab = document.createElement("div");
    tab.className = `sheet-tab${first ? " active" : ""}`;
    tab.textContent = sheetName;
    tab.addEventListener("click", () => switchSheet(sheetName));

    if (sheetName === "Repair Cases" || sheetName === "Closed Cases Report") {
      right.appendChild(tab);
    } else {
      left.appendChild(tab);
    }

    first = false;
  });
}

function switchSheet(sheetName) {
  document.querySelectorAll(".sheet-tab").forEach(tab => {
    tab.classList.toggle("active", tab.textContent === sheetName);
  });

  Object.keys(tablesMap).forEach(name => {
    const wrapper = tablesMap[name];
    const isActive = name === sheetName;
    wrapper.style.display = isActive ? "block" : "none";

    if (isActive) {
      const dataTable = dataTablesMap[sheetName];
      dataTable.columns.adjust().draw(false);
    }
  });
}

function loadDataFromDB() {
  const store = getStore("readonly");
  const request = store.getAll();

  request.onsuccess = () => {
    const records = request.result;

    records.forEach(record => {
      const sheetName = record.sheetName;
      const dt = dataTablesMap[sheetName];
      if (!dt) return;

      dt.clear();
      record.rows.forEach(row => {
        const normalized = normalizeRowToSchema(row, sheetName);
        dt.row.add(["", ...normalized]);
      });
      dt.draw(false);
    });
  };
}

function showProgressOverlay() {
  selectors.overlay.style.display = "flex";
}

function hideProgressOverlay() {
  selectors.overlay.style.display = "none";
  selectors.overlay.classList.remove("progress-complete");
  selectors.overlayConfirm.style.display = "none";
  selectors.overlayProgress.style.width = "0%";
  selectors.overlayProgressText.textContent = "0%";
}

let progressContext = null;
let displayedProgress = 0;
let progressAnimFrame = null;

function startProgressContext(label) {
  progressContext = { label, value: 0 };
  displayedProgress = 0;

  showProgressOverlay();
  selectors.overlayStatus.textContent = label;
  selectors.overlayProgress.style.width = "0%";
  selectors.overlayProgressText.textContent = "0%";
}

function animateProgressTo(targetPercent, duration = 280) {
  if (progressAnimFrame) {
    cancelAnimationFrame(progressAnimFrame);
    progressAnimFrame = null;
  }

  const start = performance.now();
  const startPercent = displayedProgress;
  const delta = targetPercent - startPercent;

  const step = now => {
    const elapsed = now - start;
    const progress = Math.min(elapsed / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3);
    const current = startPercent + delta * eased;

    displayedProgress = current;
    selectors.overlayProgress.style.width = `${current.toFixed(2)}%`;
    selectors.overlayProgressText.textContent = `${Math.round(current)}%`;

    if (progress < 1) {
      progressAnimFrame = requestAnimationFrame(step);
    } else {
      displayedProgress = targetPercent;
      selectors.overlayProgress.style.width = `${targetPercent}%`;
      selectors.overlayProgressText.textContent = `${targetPercent}%`;
      progressAnimFrame = null;
    }
  };

  progressAnimFrame = requestAnimationFrame(step);
}

function updateProgressContext(current, total, text) {
  if (!progressContext || total === 0) return;

  const percent = Math.min(100, Math.round((current / total) * 100));
  animateProgressTo(percent);

  if (text) {
    selectors.overlayStatus.textContent = text;
  }
}

function endProgressContext(text = "Completed") {
  animateProgressTo(100, 200);
  selectors.overlayStatus.textContent = text;
  selectors.overlay.classList.add("progress-complete");
  selectors.overlayConfirm.style.display = "inline-block";

  selectors.overlayConfirm.onclick = () => {
    hideProgressOverlay();
    progressContext = null;
  };
}

function formatSentenceCase(value) {
  if (!value) return "";
  const str = String(value).trim().toLowerCase();
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function processExcelFile(file, allowedSheets) {
  return new Promise(resolve => {
    const reader = new FileReader();

    reader.onload = async event => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const filteredWorkbook = {
        SheetNames: workbook.SheetNames.filter(sheet => allowedSheets.includes(sheet)),
        Sheets: workbook.Sheets
      };

      await buildSheetTables(filteredWorkbook);
      resolve();
    };

    reader.readAsArrayBuffer(file);
  });
}

function buildSheetTables(workbook) {
  return new Promise(async resolve => {
    const sheetNames = workbook.SheetNames;

    let processedRows = 0;
    const totalRows = sheetNames.reduce((sum, sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      return sum + (sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1 }).length : 0);
    }, 0);

    for (let index = 0; index < sheetNames.length; index += 1) {
      const sheetName = sheetNames[index];
      await new Promise(requestAnimationFrame);

      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
      if (!json.length) continue;

      const headers = TABLE_SCHEMAS[sheetName];
      if (!headers) continue;

      const rows = json
        .slice(1)
        .map(row => {
          const excelRow = row.slice(3);
          if (!excelRow.some(cell => String(cell || "").trim() !== "")) {
            return null;
          }

          const cleanedRow = [];
          for (let i = 0; i < headers.length; i += 1) {
            let cell = excelRow[i];
            if (typeof cell === "number" && cell > 40000 && cell < 60000) {
              try {
                cleanedRow.push(formatDate(excelDateToJSDate(cell)));
              } catch (error) {
                cleanedRow.push(cleanCell(cell));
              }
            } else {
              cleanedRow.push(cleanCell(cell));
            }
          }

          return cleanedRow;
        })
        .filter(Boolean);

      const dataTable = dataTablesMap[sheetName];
      dataTable.clear();

      rows.forEach(row => {
        dataTable.row.add(["", ...row]);
        processedRows += 1;
        updateProgressContext(
          processedRows,
          totalRows,
          `Processing ${sheetName} (${processedRows}/${totalRows})`
        );
      });

      dataTable.draw(false);

      getStore("readwrite").put({
        sheetName,
        rows: rows.map(row => normalizeRowToSchema(row, sheetName)),
        lastUpdated: new Date().toISOString()
      });
    }

    resolve();
  });
}

function processCsvFile(file, targetSheetName) {
  return new Promise(resolve => {
    const reader = new FileReader();

    reader.onload = event => {
      const text = event.target.result;
      const rows = XLSX.utils.sheet_to_json(
        XLSX.read(text, { type: "string" }).Sheets.Sheet1,
        { header: 1, raw: true }
      );

      if (!rows.length) {
        resolve();
        return;
      }

      const headers = TABLE_SCHEMAS[targetSheetName];
      const dataRows = rows.slice(1).map(row => {
        const cleaned = [];
        for (let i = 0; i < headers.length; i += 1) {
          let cellValue = cleanCell(row[i]);
          if (targetSheetName === "CSO Status" && headers[i] === "Status") {
            cellValue = formatSentenceCase(cellValue);
          }
          cleaned.push(cellValue);
        }
        return cleaned;
      });

      const dt = dataTablesMap[targetSheetName];
      dt.clear();
      dataRows.forEach(row => dt.row.add(["", ...row]));
      dt.draw(false);

      const store = getStore("readwrite");
      store.put({
        sheetName: targetSheetName,
        rows: dataRows,
        lastUpdated: new Date().toISOString()
      });

      resolve();
    };

    reader.readAsText(file);
  });
}

function parseGNProCSV(text) {
  const rows = XLSX.utils.sheet_to_json(
    XLSX.read(text, { type: "string" }).Sheets.Sheet1,
    { header: 1, raw: true }
  );

  const map = new Map();
  rows.slice(1).forEach(row => {
    const caseId = cleanCell(row[0]);
    if (!caseId) return;

    map.set(caseId, {
      status: formatSentenceCase(cleanCell(row[2])),
      tracking: cleanCell(row[3]),
      repairStatus: cleanCell(row[4])
    });
  });

  return map;
}

function stripOrderSuffix(orderId) {
  return String(orderId || "").replace(/-0\d$/, "");
}

async function processGNProCSOFile(file) {
  const store = getStore("readonly");
  const allData = await new Promise(res => {
    const req = store.getAll();
    req.onsuccess = () => res(req.result);
  });

  const dump = allData.find(row => row.sheetName === "Dump")?.rows || [];
  const so = allData.find(row => row.sheetName === "SO")?.rows || [];
  const oldCso = allData.find(row => row.sheetName === "CSO Status")?.rows || [];

  const csvText = await file.text();
  const gnproMap = parseGNProCSV(csvText);

  const dumpCaseIdx = TABLE_SCHEMAS["Dump"].indexOf("Case ID");
  const dumpResIdx = TABLE_SCHEMAS["Dump"].indexOf("Case Resolution Code");

  const soCaseIdx = TABLE_SCHEMAS["SO"].indexOf("Case ID");
  const soDateIdx = TABLE_SCHEMAS["SO"].indexOf("Date and Time Submitted");
  const soOrderIdx = TABLE_SCHEMAS["SO"].indexOf("Order Reference ID");

  const oldCaseIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Case ID");
  const oldStatusIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Status");
  const oldTrackIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Tracking Number");

  const offsiteCases = [
    ...new Set(
      dump
        .filter(row => normalizeText(row[dumpResIdx]) === "offsite solution")
        .map(row => row[dumpCaseIdx])
    )
  ];

  const finalRows = [];
  let processed = 0;
  const total = offsiteCases.length;

  offsiteCases.forEach(caseId => {
    processed += 1;
    updateProgressContext(
      processed,
      total,
      `Processing CSO cases (${processed}/${total})`
    );

    const soRows = so.filter(row => row[soCaseIdx] === caseId);
    if (!soRows.length) return;

    const latestSO = soRows.sort((a, b) => new Date(b[soDateIdx]) - new Date(a[soDateIdx]))[0];
    const cso = stripOrderSuffix(latestSO[soOrderIdx]);

    let status = "Not Found";
    let tracking = "";

    const oldRow = oldCso.find(row => row[oldCaseIdx] === caseId);
    if (oldRow) {
      const oldStatus = normalizeText(oldRow[oldStatusIdx]);
      if (oldStatus === "delivered" || oldStatus === "order cancelled, not to be reopened") {
        status = formatSentenceCase(oldRow[oldStatusIdx]);
        tracking = oldRow[oldTrackIdx];
        finalRows.push([caseId, cso, status, tracking, oldRow[4] || ""]);
        return;
      }
    }

    const csvRow = gnproMap.get(caseId);
    let repairStatus = "";

    if (csvRow) {
      status = csvRow.status;
      tracking = csvRow.tracking;
      repairStatus = csvRow.repairStatus || "";
    }

    finalRows.push([caseId, cso, status, tracking, repairStatus]);
  });

  const dt = dataTablesMap["CSO Status"];
  dt.clear();
  finalRows.forEach(row => dt.row.add(["", ...row]));
  dt.draw(false);

  const writeStore = getStore("readwrite");
  writeStore.put({
    sheetName: "CSO Status",
    rows: finalRows,
    lastUpdated: new Date().toISOString()
  });
}

function toDateOnly(date) {
  const value = new Date(date);
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

function diffCalendarDays(from, to = new Date()) {
  const a = toDateOnly(from);
  const b = toDateOnly(to);
  return Math.floor((b - a) / (24 * 60 * 60 * 1000));
}

function getCAGroup(createdOn) {
  const days = diffCalendarDays(createdOn);

  if (days <= 3) return "0-3 Days";
  if (days <= 5) return "3-5 Days";
  if (days <= 10) return "5-10 Days";
  if (days <= 15) return "10-15 Days";
  if (days <= 30) return "15-30 Days";
  if (days <= 60) return "30-60 Days";
  if (days <= 90) return "60-90 Days";
  return "> 90 Days";
}

function getMOStatusPriority(status) {
  const value = normalizeText(status);

  if (value === "closed") return 1;
  if (value === "pod") return 2;
  if (value === "shipped") return 3;
  if (value === "ordered") return 4;
  if (value === "partially ordered") return 5;
  if (value === "order pending") return 6;
  if (value === "new") return 7;
  if (value === "cancelled") return 8;

  return 9;
}

function getLatestMO(caseId, mo) {
  const caseMOs = mo.filter(row => row[1] === caseId);
  if (!caseMOs.length) return null;

  caseMOs.sort((a, b) => new Date(b[2]) - new Date(a[2]));
  const latestTime = new Date(caseMOs[0][2]);

  const windowMOs = caseMOs.filter(row => Math.abs(new Date(row[2]) - latestTime) <= 5 * 60 * 1000);
  windowMOs.sort((a, b) => getMOStatusPriority(a[3]) - getMOStatusPriority(b[3]));

  return windowMOs[0];
}

function getFirstOrderDate(caseId, wo, mo, so) {
  const dates = [];

  wo.forEach(row => row[0] === caseId && row[6] && dates.push(new Date(row[6])));
  mo.forEach(row => row[1] === caseId && row[2] && dates.push(new Date(row[2])));
  so.forEach(row => row[0] === caseId && row[2] && dates.push(new Date(row[2])));

  if (!dates.length) return null;
  return new Date(Math.min(...dates));
}

function calculateSBD(caseRow, firstOrderDate, sbdConfig) {
  if (!firstOrderDate || !sbdConfig?.periods) return "NA";

  const caseCreated = new Date(caseRow.createdOn);
  const caseDateOnly = toDateOnly(caseCreated);

  const period = sbdConfig.periods.find(item =>
    item.startDate &&
    item.endDate &&
    caseDateOnly >= new Date(item.startDate) &&
    caseDateOnly <= new Date(item.endDate)
  );

  if (!period) return "NA";

  const countryRow = period.rows.find(row =>
    normalizeText(row.country) === normalizeText(caseRow.country)
  );

  if (!countryRow) return "NA";

  const cutOff = countryRow.time;
  if (!cutOff) return "NA";

  const cutOffDate = new Date(caseDateOnly);
  const [hh, mm] = cutOff.split(":");
  cutOffDate.setHours(hh, mm, 0, 0);

  if (caseCreated > cutOffDate) {
    cutOffDate.setDate(cutOffDate.getDate() + 1);
  }

  return firstOrderDate <= cutOffDate ? "Met" : "Not Met";
}

async function buildCopySOOrders() {
  const store = getStore("readonly");
  const allData = await new Promise(res => {
    const req = store.getAll();
    req.onsuccess = () => res(req.result);
  });

  const dump = allData.find(row => row.sheetName === "Dump")?.rows || [];
  const so = allData.find(row => row.sheetName === "SO")?.rows || [];
  const cso = allData.find(row => row.sheetName === "CSO Status")?.rows || [];

  const dumpIdx = TABLE_SCHEMAS["Dump"].indexOf("Case Resolution Code");
  const dumpCaseIdx = TABLE_SCHEMAS["Dump"].indexOf("Case ID");

  const soCaseIdx = TABLE_SCHEMAS["SO"].indexOf("Case ID");
  const soDateIdx = TABLE_SCHEMAS["SO"].indexOf("Date and Time Submitted");
  const soOrderIdx = TABLE_SCHEMAS["SO"].indexOf("Order Reference ID");

  const csoCaseIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Case ID");
  const csoStatusIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Status");

  const offsiteCases = [
    ...new Set(
      dump
        .filter(row => normalizeText(row[dumpIdx]) === "offsite solution")
        .map(row => row[dumpCaseIdx])
    )
  ];

  const result = [];

  offsiteCases.forEach(caseId => {
    const soRows = so.filter(row => row[soCaseIdx] === caseId);
    if (!soRows.length) return;

    const latest = soRows.sort((a, b) => new Date(b[soDateIdx]) - new Date(a[soDateIdx]))[0];
    let orderId = stripOrderSuffix(latest[soOrderIdx]);
    if (!orderId) return;

    const csoRow = cso.find(row => row[csoCaseIdx] === caseId);
    if (csoRow) {
      const status = normalizeText(csoRow[csoStatusIdx]);
      if (status === "delivered" || status === "order cancelled, not to be reopened") {
        return;
      }
    }

    result.push(`${caseId},${orderId}`);
  });

  return result.join("\n");
}

async function buildCopyTrackingURLs() {
  const store = getStore("readonly");
  const allData = await new Promise(res => {
    const req = store.getAll();
    req.onsuccess = () => res(req.result);
  });

  const dump = allData.find(row => row.sheetName === "Dump")?.rows || [];
  const mo = allData.find(row => row.sheetName === "MO")?.rows || [];
  const moItems = allData.find(row => row.sheetName === "MO Items")?.rows || [];
  const cso = allData.find(row => row.sheetName === "CSO Status")?.rows || [];
  const delivery = allData.find(row => row.sheetName === "Delivery Details")?.rows || [];

  const dumpCaseIdx = TABLE_SCHEMAS["Dump"].indexOf("Case ID");
  const dumpResIdx = TABLE_SCHEMAS["Dump"].indexOf("Case Resolution Code");

  const moOrderIdx = TABLE_SCHEMAS["MO"].indexOf("Order Number");
  const moCaseIdx = TABLE_SCHEMAS["MO"].indexOf("Case ID");
  const moCreatedIdx = TABLE_SCHEMAS["MO"].indexOf("Created On");
  const moStatusIdx = TABLE_SCHEMAS["MO"].indexOf("Order Status");

  const moItemOrderIdx = TABLE_SCHEMAS["MO Items"].indexOf("Material Order");
  const moItemNameIdx = TABLE_SCHEMAS["MO Items"].indexOf("MO Line Items Name");
  const moItemUrlIdx = TABLE_SCHEMAS["MO Items"].indexOf("Tracking Url");

  const csoCaseIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Case ID");
  const csoStatusIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Status");
  const csoTrackIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Tracking Number");

  const delCaseIdx = TABLE_SCHEMAS["Delivery Details"].indexOf("CaseID");
  const delStatusIdx = TABLE_SCHEMAS["Delivery Details"].indexOf("CurrentStatus");

  const partsShippedCases = [
    ...new Set(
      dump
        .filter(row => normalizeText(row[dumpResIdx]) === "parts shipped")
        .map(row => row[dumpCaseIdx])
    )
  ];

  const trackingMap = new Map();

  partsShippedCases.forEach(caseId => {
    const caseMOs = mo.filter(row => row[moCaseIdx] === caseId);
    if (!caseMOs.length) return;

    caseMOs.sort((a, b) => new Date(b[moCreatedIdx]) - new Date(a[moCreatedIdx]));
    const latestTime = new Date(caseMOs[0][moCreatedIdx]);

    const windowMOs = caseMOs.filter(row => Math.abs(new Date(row[moCreatedIdx]) - latestTime) <= 5 * 60 * 1000);
    windowMOs.sort((a, b) => getMOStatusPriority(a[moStatusIdx]) - getMOStatusPriority(b[moStatusIdx]));

    const selected = windowMOs[0];
    const status = normalizeText(selected[moStatusIdx]);
    if (status !== "closed" && status !== "pod") return;

    const moNumber = selected[moOrderIdx];
    const item = moItems.find(row => row[moItemOrderIdx] === moNumber && normalizeText(row[moItemNameIdx]).endsWith("- 1"));
    if (!item || !item[moItemUrlIdx]) return;

    trackingMap.set(caseId, item[moItemUrlIdx]);
  });

  cso.forEach(row => {
    const caseId = row[csoCaseIdx];
    if (trackingMap.has(caseId)) return;

    if (normalizeText(row[csoStatusIdx]) === "delivered") {
      const tn = row[csoTrackIdx];
      if (!tn) return;

      const url =
        "http://wwwapps.ups.com/WebTracking/processInputRequest" +
        "?TypeOfInquiryNumber=T&InquiryNumber1=" +
        tn;

      trackingMap.set(caseId, url);
    }
  });

  delivery.forEach(row => {
    const caseId = row[delCaseIdx];
    const status = normalizeText(row[delStatusIdx]);

    if (trackingMap.has(caseId) && status && status !== "no status found") {
      trackingMap.delete(caseId);
    }
  });

  return [...trackingMap.entries()]
    .map(([key, value]) => `${key} | ${value}`)
    .join("\n");
}

function parseTrackingResultsCSV(text) {
  const rows = XLSX.utils.sheet_to_json(
    XLSX.read(text, { type: "string" }).Sheets.Sheet1,
    { header: 1, raw: true }
  );

  const map = new Map();
  rows.slice(1).forEach(row => {
    const caseId = cleanCell(row[0]);
    const status = cleanCell(row[1]);
    if (!caseId) return;
    map.set(caseId, status);
  });

  return map;
}

async function processTrackingResultsFile(file) {
  const store = getStore("readonly");
  const allData = await new Promise(res => {
    const req = store.getAll();
    req.onsuccess = () => res(req.result);
  });

  const dump = allData.find(row => row.sheetName === "Dump")?.rows || [];
  const mo = allData.find(row => row.sheetName === "MO")?.rows || [];
  const moItems = allData.find(row => row.sheetName === "MO Items")?.rows || [];
  const cso = allData.find(row => row.sheetName === "CSO Status")?.rows || [];
  const oldDelivery = allData.find(row => row.sheetName === "Delivery Details")?.rows || [];

  const dumpCaseIdx = TABLE_SCHEMAS["Dump"].indexOf("Case ID");
  const dumpResIdx = TABLE_SCHEMAS["Dump"].indexOf("Case Resolution Code");

  const moCaseIdx = TABLE_SCHEMAS["MO"].indexOf("Case ID");
  const moOrderIdx = TABLE_SCHEMAS["MO"].indexOf("Order Number");
  const moCreatedIdx = TABLE_SCHEMAS["MO"].indexOf("Created On");
  const moStatusIdx = TABLE_SCHEMAS["MO"].indexOf("Order Status");

  const moItemOrderIdx = TABLE_SCHEMAS["MO Items"].indexOf("Material Order");
  const moItemNameIdx = TABLE_SCHEMAS["MO Items"].indexOf("MO Line Items Name");
  const moItemUrlIdx = TABLE_SCHEMAS["MO Items"].indexOf("Tracking Url");

  const csoCaseIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Case ID");
  const csoStatusIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Status");

  const oldDelCaseIdx = TABLE_SCHEMAS["Delivery Details"].indexOf("CaseID");
  const oldDelStatusIdx = TABLE_SCHEMAS["Delivery Details"].indexOf("CurrentStatus");

  const csvText = await file.text();
  const trackingCSVMap = parseTrackingResultsCSV(csvText);

  const partsShippedCases = [
    ...new Set(
      dump
        .filter(row => normalizeText(row[dumpResIdx]) === "parts shipped")
        .map(row => row[dumpCaseIdx])
    )
  ];

  const finalCaseSet = new Set();

  partsShippedCases.forEach(caseId => {
    const caseMOs = mo.filter(row => row[moCaseIdx] === caseId);
    if (!caseMOs.length) return;

    caseMOs.sort((a, b) => new Date(b[moCreatedIdx]) - new Date(a[moCreatedIdx]));
    const latestTime = new Date(caseMOs[0][moCreatedIdx]);

    const windowMOs = caseMOs.filter(row => Math.abs(new Date(row[moCreatedIdx]) - latestTime) <= 5 * 60 * 1000);
    windowMOs.sort((a, b) => getMOStatusPriority(a[moStatusIdx]) - getMOStatusPriority(b[moStatusIdx]));

    const selected = windowMOs[0];
    const status = normalizeText(selected[moStatusIdx]);

    if (status !== "closed" && status !== "pod") return;

    const moNumber = selected[moOrderIdx];
    const item = moItems.find(row =>
      row[moItemOrderIdx] === moNumber &&
      normalizeText(row[moItemNameIdx]).endsWith("- 1") &&
      row[moItemUrlIdx]
    );

    if (!item) return;
    finalCaseSet.add(caseId);
  });

  cso.forEach(row => {
    if (normalizeText(row[csoStatusIdx]) === "delivered") {
      finalCaseSet.add(row[csoCaseIdx]);
    }
  });

  const finalRows = [];
  let processed = 0;
  const total = finalCaseSet.size;

  finalCaseSet.forEach(caseId => {
    processed += 1;
    updateProgressContext(
      processed,
      total,
      `Updating tracking (${processed}/${total})`
    );

    const oldRow = oldDelivery.find(row => row[oldDelCaseIdx] === caseId);
    if (oldRow && oldRow[oldDelStatusIdx]) {
      finalRows.push([caseId, oldRow[oldDelStatusIdx]]);
      return;
    }

    if (trackingCSVMap.has(caseId)) {
      finalRows.push([caseId, trackingCSVMap.get(caseId)]);
      return;
    }

    finalRows.push([caseId, "No Status Found"]);
  });

  const dt = dataTablesMap["Delivery Details"];
  dt.clear();
  finalRows.forEach(row => dt.row.add(["", ...row]));
  dt.draw(false);

  const writeStore = getStore("readwrite");
  writeStore.put({
    sheetName: "Delivery Details",
    rows: finalRows,
    lastUpdated: new Date().toISOString()
  });
}

function createEmptySbdData() {
  return {
    sheetName: "SBD Cut Off Times",
    periods: Array.from({ length: 3 }, () => ({
      startDate: "",
      endDate: "",
      rows: []
    }))
  };
}

function datesOverlap(aStart, aEnd, bStart, bEnd) {
  return aStart && aEnd && bStart && bEnd && !(aEnd < bStart || bEnd < aStart);
}

function renderSbdModal(data) {
  const container = document.querySelector(".sbd-periods");
  container.innerHTML = "";

  data.periods.forEach((period, idx) => {
    const div = document.createElement("div");
    div.className = "sbd-period";

    div.innerHTML = `
      <h4>Date Period ${idx + 1}</h4>
      <div class="sbd-dates">
        <input type="date" class="sbd-start" value="${period.startDate}">
        <input type="date" class="sbd-end" value="${period.endDate}">
      </div>
      <table class="sbd-table">
        <thead>
          <tr><th>Country</th><th>Cut Off Time</th><th></th></tr>
        </thead>
      </table>
      <div class="sbd-tbody-scroll">
        <table class="sbd-table">
          <tbody></tbody>
        </table>
      </div>
      <div class="ghost-button add-row">+ Add Row</div>
    `;

    const tbody = div.querySelector("tbody");

    function addRow(row = { country: "", time: "" }) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><input value="${row.country}"></td>
        <td><input type="time" value="${row.time}"></td>
        <td><button class="ghost-button">✕</button></td>
      `;
      tr.querySelector("button").onclick = () => tr.remove();
      tbody.appendChild(tr);
    }

    period.rows.forEach(addRow);
    div.querySelector(".add-row").onclick = () => addRow();

    container.appendChild(div);
  });
}

async function saveSbdData() {
  const periods = [];
  const blocks = document.querySelectorAll(".sbd-period");

  blocks.forEach(block => {
    const startDate = block.querySelector(".sbd-start").value;
    const endDate = block.querySelector(".sbd-end").value;

    const rows = [];
    block.querySelectorAll("tbody tr").forEach(tr => {
      const country = tr.children[0].querySelector("input").value.trim();
      const time = tr.children[1].querySelector("input").value;
      if (country && time) rows.push({ country, time });
    });

    periods.push({ startDate, endDate, rows });
  });

  for (let i = 0; i < periods.length; i += 1) {
    for (let j = i + 1; j < periods.length; j += 1) {
      if (datesOverlap(periods[i].startDate, periods[i].endDate, periods[j].startDate, periods[j].endDate)) {
        alert("Date ranges cannot overlap between periods.");
        return;
      }
    }
  }

  const store = getStore("readwrite");
  store.put({
    sheetName: "SBD Cut Off Times",
    periods
  });

  document.getElementById("sbdModal").style.display = "none";
}

function adjustMappingModalWidth(containerId, modalId) {
  const container = document.getElementById(containerId);
  const modal = document.querySelector(`#${modalId} .modal`);
  const blocks = container.querySelectorAll(".mapping-block:not(.placeholder)");
  const count = Math.max(1, Math.min(blocks.length, 3));

  const blockWidth = 320;
  const gap = 12;

  const width = count * blockWidth + (count - 1) * gap;
  modal.style.width = `${width}px`;
}

function renderTLModal(data = []) {
  const container = document.getElementById("tlContainer");
  container.innerHTML = "";

  if (!data.length) {
    addTLBlock();
    container.firstChild.classList.add("placeholder");
  } else {
    data.forEach(tl => addTLBlock(tl.name, tl.agents));
  }

  adjustMappingModalWidth("tlContainer", "tlModal");
}

function addTLBlock(name = "", agents = []) {
  const container = document.getElementById("tlContainer");
  const block = document.createElement("div");
  block.className = "mapping-block";

  block.innerHTML = `
    <h4>
      TL: <input value="${name}">
      <button class="ghost-button" onclick="this.closest('.mapping-block').remove()">✕</button>
    </h4>
    <div class="agents"></div>
    <button class="ghost-button add-row">+ Add Agent</button>
  `;

  const agentsDiv = block.querySelector(".agents");
  const addAgentBtn = block.querySelector(".add-row");

  function addAgent(value = "") {
    const row = document.createElement("div");
    row.className = "mapping-row";
    row.innerHTML = `
      <input value="${value}">
      <button class="ghost-button" onclick="this.parentElement.remove()">✕</button>
    `;
    agentsDiv.appendChild(row);
  }

  agents.forEach(addAgent);
  addAgentBtn.onclick = () => addAgent();

  container.appendChild(block);
  adjustMappingModalWidth("tlContainer", "tlModal");
}

function renderMarketModal(data = []) {
  const container = document.getElementById("marketContainer");
  container.innerHTML = "";

  if (!data.length) {
    addMarketBlock();
    container.firstChild.classList.add("placeholder");
  } else {
    data.forEach(market => addMarketBlock(market.name, market.countries));
  }

  adjustMappingModalWidth("marketContainer", "marketModal");
}

function addMarketBlock(name = "", countries = []) {
  const container = document.getElementById("marketContainer");
  const block = document.createElement("div");
  block.className = "mapping-block";

  block.innerHTML = `
    <h4>
      Market: <input value="${name}">
      <button class="ghost-button" onclick="this.closest('.mapping-block').remove()">✕</button>
    </h4>
    <div class="countries"></div>
    <button class="ghost-button add-row">+ Add Country</button>
  `;

  const countriesDiv = block.querySelector(".countries");
  const addBtn = block.querySelector(".add-row");

  function addCountry(value = "") {
    const row = document.createElement("div");
    row.className = "mapping-row";
    row.innerHTML = `
      <input value="${value}">
      <button class="ghost-button" onclick="this.parentElement.remove()">✕</button>
    `;
    countriesDiv.appendChild(row);
  }

  countries.forEach(addCountry);
  addBtn.onclick = () => addCountry();

  container.appendChild(block);
  adjustMappingModalWidth("marketContainer", "marketModal");
}

async function buildRepairCases() {
  const store = getStore("readonly");
  const all = await new Promise(res => {
    const req = store.getAll();
    req.onsuccess = () => res(req.result);
  });

  const dump = all.find(row => row.sheetName === "Dump")?.rows || [];
  const wo = all.find(row => row.sheetName === "WO")?.rows || [];
  const mo = all.find(row => row.sheetName === "MO")?.rows || [];
  const moItems = all.find(row => row.sheetName === "MO Items")?.rows || [];
  const so = all.find(row => row.sheetName === "SO")?.rows || [];
  const cso = all.find(row => row.sheetName === "CSO Status")?.rows || [];
  const delivery = all.find(row => row.sheetName === "Delivery Details")?.rows || [];

  const tlMap = all.find(row => row.sheetName === "TL_MAP")?.data || [];
  const marketMap = all.find(row => row.sheetName === "MARKET_MAP")?.data || [];
  const sbdConfig = all.find(row => row.sheetName === "SBD Cut Off Times");

  const validCases = dump.filter(row =>
    ["parts shipped", "onsite solution", "offsite solution"].includes(normalizeText(row[8]))
  );

  const rows = [];

  validCases.forEach(dumpRow => {
    const caseId = dumpRow[0];
    const resolution = normalizeText(dumpRow[8]);

    const firstOrder = getFirstOrderDate(caseId, wo, mo, so);

    const tl =
      tlMap.find(tlRow =>
        tlRow.agents.some(agent => normalizeText(agent) === normalizeText(dumpRow[9]))
      )?.name || "";

    const market =
      marketMap.find(marketRow =>
        marketRow.countries.some(country => normalizeText(country) === normalizeText(dumpRow[6]))
      )?.name || "";

    const onsiteRFC =
      resolution === "onsite solution"
        ? (wo
            .filter(row => row[0] === caseId)
            .sort((a, b) => new Date(b[6]) - new Date(a[6]))[0]?.[5] || "")
        : "Not Found";

    const csrRFC =
      resolution === "parts shipped"
        ? (getLatestMO(caseId, mo)?.[3] || "")
        : "Not Found";

    const benchRFC =
      resolution === "offsite solution"
        ? (cso.find(row => row[0] === caseId)?.[2] || "")
        : "Not Found";

    const dnap =
      resolution === "offsite solution" &&
      normalizeText(cso.find(row => row[0] === caseId)?.[4]).includes("product returned unrepaired to customer")
        ? "True"
        : "";

    let partNumber = "";
    let partName = "";

    if (resolution === "parts shipped") {
      const latestMO = getLatestMO(caseId, mo);
      if (latestMO) {
        const moNumber = latestMO[0];
        const item = moItems.find(row => row[0] === moNumber && normalizeText(row[1]).endsWith("- 1"));

        if (item) {
          partNumber = item[2] || "";

          if (item[3]) {
            const idx = item[3].indexOf("-");
            partName = idx >= 0 ? item[3].substring(idx + 1).trim() : item[3].trim();
          }
        }
      }
    }

    rows.push([
      caseId,
      dumpRow[1],
      dumpRow[2],
      dumpRow[3],
      dumpRow[6],
      dumpRow[8],
      dumpRow[9],
      dumpRow[15],
      getCAGroup(dumpRow[2]),
      tl,
      calculateSBD({ createdOn: dumpRow[2], country: dumpRow[6] }, firstOrder, sbdConfig),
      onsiteRFC,
      csrRFC,
      benchRFC,
      market,
      onsiteRFC !== "Not Found"
        ? (wo
            .filter(row => row[0] === caseId)
            .sort((a, b) => new Date(b[6]) - new Date(a[6]))[0]?.[9] || "")
        : "",
      delivery.find(row => row[0] === caseId)?.[1] || "",
      partNumber,
      partName,
      dumpRow[16],
      dumpRow[17],
      dumpRow[10],
      dnap
    ]);
  });

  const dt = dataTablesMap["Repair Cases"];
  dt.clear();
  rows.forEach(row => dt.row.add(["", ...row]));
  dt.draw(false);

  getStore("readwrite").put({
    sheetName: "Repair Cases",
    rows,
    lastUpdated: new Date().toISOString()
  });
}

async function buildClosedCasesReport() {
  const store = getStore("readonly");
  const all = await new Promise(res => {
    const req = store.getAll();
    req.onsuccess = () => res(req.result);
  });

  const closed = all.find(row => row.sheetName === "Closed Cases")?.rows || [];
  const repair = all.find(row => row.sheetName === "Repair Cases")?.rows || [];

  const now = new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth() - 6, 1);

  const rows = [];

  closed.forEach(row => {
    const closedDate = new Date(row[4]);
    if (closedDate < cutoff) return;

    const repairRow = repair.find(repairItem => repairItem[0] === row[0]);

    let closedBy = row[2];
    if (closedBy === "# CrmWebJobUser-Prod") closedBy = "CRM Auto Closed";
    else if (
      [
        "# MSFT-ServiceSystemAdmin",
        "# CrmEEGUser-Prod",
        "# MSFT-ServiceSystemAdminDev",
        "SYSTEM"
      ].includes(closedBy)
    ) {
      closedBy = row[8];
    }

    rows.push([
      row[0],
      repairRow?.[1] || "",
      row[1],
      row[6],
      row[2],
      row[3],
      row[4],
      closedBy,
      row[9],
      row[5],
      row[8],
      row[10],
      repairRow?.[9] || "",
      repairRow?.[10] || "",
      repairRow?.[14] || ""
    ]);
  });

  const dt = dataTablesMap["Closed Cases Report"];
  dt.clear();
  rows.forEach(row => dt.row.add(["", ...row]));
  dt.draw(false);

  const write = getStore("readwrite");
  write.put({
    sheetName: "Closed Cases Report",
    rows,
    lastUpdated: new Date().toISOString()
  });

  const remaining = repair.filter(repairRow => !rows.some(closedRow => closedRow[0] === repairRow[0]));
  write.put({ sheetName: "Repair Cases", rows: remaining });
}

function setTheme(theme) {
  document.body.setAttribute("data-theme", theme);
  localStorage.setItem("kci-theme", theme);
  document.getElementById("themeToggle").textContent = theme === "dark" ? "🌙" : "☀️";
}

function openModal(id) {
  document.getElementById(id).style.display = "flex";
}

function closeModal(id) {
  document.getElementById(id).style.display = "none";
}

function enableProcessIfReady() {
  const hasFile = selectedFiles.kci || selectedFiles.cso || selectedFiles.tracking;
  selectors.processBtn.disabled = !hasFile;
}

selectors.kciInput.addEventListener("change", event => {
  const file = event.target.files[0];
  if (!file) return;

  if (!file.name.startsWith("KCI - Open Repair Case Data")) {
    alert("Invalid file. Please upload 'KCI - Open Repair Case Data' file.");
    event.target.value = "";
    return;
  }

  selectedFiles.kci = file;
  enableProcessIfReady();
});

selectors.csoInput.addEventListener("change", event => {
  const file = event.target.files[0];
  if (!file) return;

  if (!/^GNPro_Case_CSO_Status_\d{4}-\d{2}-\d{2}/.test(file.name)) {
    alert("Invalid GNPro CSO file name.");
    event.target.value = "";
    return;
  }

  selectedFiles.cso = file;
  enableProcessIfReady();
});

selectors.trackingInput.addEventListener("change", event => {
  const file = event.target.files[0];
  if (!file) return;

  if (!/^Tracking_Results_\d{4}-\d{2}-\d{2}/.test(file.name)) {
    alert("Invalid Tracking Results file name.");
    event.target.value = "";
    return;
  }

  selectedFiles.tracking = file;
  enableProcessIfReady();
});

selectors.processBtn.addEventListener("click", async () => {
  if (selectedFiles.kci) {
    startProgressContext("Processing KCI Excel...");

    const store = getStore("readwrite");
    ["Dump", "WO", "MO", "MO Items", "SO", "Closed Cases"].forEach(sheet => {
      dataTablesMap[sheet]?.clear().draw(false);
      store.put({ sheetName: sheet, rows: [], lastUpdated: new Date().toISOString() });
    });

    await processExcelFile(selectedFiles.kci, [
      "Dump",
      "WO",
      "MO",
      "MO Items",
      "SO",
      "Closed Cases"
    ]);

    selectedFiles.kci = null;
    selectors.kciInput.value = "";
    enableProcessIfReady();
    endProgressContext("KCI Excel processed");
    return;
  }

  if (selectedFiles.cso) {
    startProgressContext("Processing GNPro CSO file...");
    await processGNProCSOFile(selectedFiles.cso);
    selectedFiles.cso = null;
    selectors.csoInput.value = "";
    enableProcessIfReady();
    endProgressContext("GNPro CSO processed");
    return;
  }

  if (selectedFiles.tracking) {
    startProgressContext("Processing Tracking Results...");
    await processTrackingResultsFile(selectedFiles.tracking);
    selectedFiles.tracking = null;
    selectors.trackingInput.value = "";
    enableProcessIfReady();
    endProgressContext("Tracking Results processed");
  }
});

document.getElementById("copySoBtn").addEventListener("click", async () => {
  const output = await buildCopySOOrders();
  const lines = output ? output.split("\n").filter(line => line.trim() !== "") : [];

  selectors.soOutput.value = lines.length ? lines.join("\n") : "No eligible cases found.";
  selectors.soCount.textContent = `Total cases: ${lines.length}`;
  selectors.soModalTitle.textContent = "Copy SO Orders Preview";
  selectors.soModal.style.display = "flex";
});

document.getElementById("copyTrackingBtn").addEventListener("click", async () => {
  const output = await buildCopyTrackingURLs();
  const lines = output ? output.split("\n").filter(line => line.trim()) : [];

  selectors.soOutput.value = lines.length ? lines.join("\n") : "No eligible tracking URLs found.";
  selectors.soCount.textContent = `Total cases: ${lines.length}`;
  selectors.soModalTitle.textContent = "Copy Tracking URLs Preview";
  selectors.soModal.style.display = "flex";
});

document.getElementById("copyToClipboardBtn").addEventListener("click", async () => {
  const text = selectors.soOutput.value;
  await navigator.clipboard.writeText(text);
  alert("Copied to clipboard");
});

document.getElementById("closeModalBtn").addEventListener("click", () => {
  selectors.soModal.style.display = "none";
});

document.getElementById("sbdBtn").onclick = () => {
  const store = getStore("readonly");
  const req = store.get("SBD Cut Off Times");

  req.onsuccess = () => {
    const data = req.result || createEmptySbdData();
    renderSbdModal(data);
    document.getElementById("sbdModal").style.display = "flex";
  };
};

document.getElementById("saveSbdBtn").onclick = saveSbdData;
document.getElementById("closeSbdBtn").onclick = () => {
  document.getElementById("sbdModal").style.display = "none";
};

document.getElementById("tlBtn").onclick = () => {
  const req = getStore().get("TL_MAP");
  req.onsuccess = () => {
    renderTLModal(req.result?.data || []);
    openModal("tlModal");
  };
};

document.getElementById("marketBtn").onclick = () => {
  const req = getStore().get("MARKET_MAP");
  req.onsuccess = () => {
    renderMarketModal(req.result?.data || []);
    openModal("marketModal");
  };
};

document.getElementById("saveTlBtn").onclick = () => {
  const blocks = document.querySelectorAll("#tlContainer .mapping-block");
  const data = [];

  blocks.forEach(block => {
    const name = block.querySelector("h4 input").value.trim();
    if (!name) return;

    const agents = [...block.querySelectorAll(".mapping-row input")]
      .map(input => input.value.trim())
      .filter(Boolean);

    data.push({ name, agents });
  });

  getStore("readwrite").put({ sheetName: "TL_MAP", data });
  closeModal("tlModal");
};

document.getElementById("saveMarketBtn").onclick = () => {
  const blocks = document.querySelectorAll("#marketContainer .mapping-block");
  const data = [];

  blocks.forEach(block => {
    const name = block.querySelector("h4 input").value.trim();
    if (!name) return;

    const countries = [...block.querySelectorAll(".mapping-row input")]
      .map(input => input.value.trim())
      .filter(Boolean);

    data.push({ name, countries });
  });

  getStore("readwrite").put({ sheetName: "MARKET_MAP", data });
  closeModal("marketModal");
};

document.getElementById("addTlBtn").onclick = () => {
  const container = document.getElementById("tlContainer");
  container.querySelector(".placeholder")?.remove();
  addTLBlock();
};

document.getElementById("addMarketBtn").onclick = () => {
  const container = document.getElementById("marketContainer");
  container.querySelector(".placeholder")?.remove();
  addMarketBlock();
};

document.getElementById("processRepairBtn").addEventListener("click", async () => {
  startProgressContext("Building Repair Cases...");

  const store = getStore("readonly");
  const all = await new Promise(res => {
    const req = store.getAll();
    req.onsuccess = () => res(req.result);
  });

  const dump = all.find(row => row.sheetName === "Dump")?.rows || [];
  const validCases = dump.filter(row =>
    ["parts shipped", "onsite solution", "offsite solution"].includes(normalizeText(row[8]))
  );

  let processed = 0;
  const total = validCases.length;

  for (const _ of validCases) {
    processed += 1;
    updateProgressContext(
      processed,
      total,
      `Processing Repair Cases (${processed}/${total})`
    );
    await new Promise(requestAnimationFrame);
  }

  await buildRepairCases();
  await buildClosedCasesReport();

  endProgressContext("Repair & Closed Case processing completed");
});

const themeToggle = document.getElementById("themeToggle");

themeToggle.addEventListener("click", () => {
  const current = document.body.getAttribute("data-theme") || "dark";
  setTheme(current === "dark" ? "light" : "dark");
});

const savedTheme = localStorage.getItem("kci-theme") || "dark";
setTheme(savedTheme);

document.addEventListener("DOMContentLoaded", async () => {
  await openDB();
  initEmptyTables();
  requestAnimationFrame(() => {
    loadDataFromDB();
  });
});

document.addEventListener("keydown", event => {
  if (event.key !== "Enter") return;

  if (selectors.overlay.style.display === "flex" && selectors.overlayConfirm.style.display !== "none") {
    event.preventDefault();
    selectors.overlayConfirm.click();
  }
});

window.closeModal = closeModal;
window.openModal = openModal;
