let kciFile = null;
let csoFile = null;
let trackingFile = null;
let workbookCache = null;
let tablesMap = {};
const dataTablesMap = {};
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

  "Closed Cases Data": [
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

// Dump sheet header display overrides (UI only) --
const DUMP_HEADER_DISPLAY_MAP = {
  "Full Name (Primary Contact) (Contact)": "Customer Name",
  "Full Name (Owning User) (User)": "Case Owner",
  "ProductName": "Product Name",
  "Product Serial Number": "Serial Number"
};

const DB_NAME = "KCI_CASE_TRACKER_DB";
const DB_VERSION = 2;
const STORE_NAME = "sheets";

let db = null;

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = function (e) {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "sheetName" });
      }
    };

    request.onsuccess = function (e) {
      db = e.target.result;
      resolve(db);
    };

    request.onerror = function () {
      reject("Failed to open IndexedDB");
    };
  });
}

function openModal(id) {
  document.getElementById(id).style.display = 'flex';
}

function closeModal(id) {
  document.getElementById(id).style.display = 'none';
}

function getStore(mode = "readonly") {
  const tx = db.transaction(STORE_NAME, mode);
  return tx.objectStore(STORE_NAME);
}

function getColumnIndex(sheetName, columnName) {
  const headers = TABLE_SCHEMAS[sheetName];
  return headers ? headers.indexOf(columnName) + 1 : -1;
  // +1 because DataTables has S.No as column 0
}

function initEmptyTables() {
  const container = document.getElementById('tablesContainer');
  container.innerHTML = '';

  const tabsDiv = document.createElement('div');
  tabsDiv.className = 'sheet-tabs';
  container.appendChild(tabsDiv);
  
  const leftTabsDiv = document.createElement('div');
  leftTabsDiv.className = 'sheet-tabs-left';
  
  const rightTabsDiv = document.createElement('div');
  rightTabsDiv.className = 'sheet-tabs-right';
  
  tabsDiv.appendChild(leftTabsDiv);
  tabsDiv.appendChild(rightTabsDiv);

  let first = true;

  Object.keys(TABLE_SCHEMAS).forEach(sheetName => {
    const headers = TABLE_SCHEMAS[sheetName];

    const tableWrapper = document.createElement('div');
    tableWrapper.className = 'table-scroll-wrapper';
    tableWrapper.style.display = first ? 'block' : 'none';
    tableWrapper.dataset.sheet = sheetName;

    const table = document.createElement('table');
    table.className = 'display';

    const thead = document.createElement('thead');
    const tr = document.createElement('tr');
    
    // Serial number column
    const snTh = document.createElement('th');
    snTh.textContent = "S.No";
    tr.appendChild(snTh);
    
    // Actual headers
    headers.forEach(h => {
      const th = document.createElement('th');
    
      // UI-only header rename for Dump sheet
      if (sheetName === "Dump" && DUMP_HEADER_DISPLAY_MAP[h]) {
        th.textContent = DUMP_HEADER_DISPLAY_MAP[h];
      } else {
        th.textContent = h;
      }
    
      tr.appendChild(th);
    });

    thead.appendChild(tr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    table.appendChild(tbody);

    tableWrapper.appendChild(table);
    container.appendChild(tableWrapper);

    const columnDefs = [
      {
        targets: 0,
        searchable: false,
        orderable: false
      }
    ];
    
    // Apply fixed width for diagnostics / notes columns
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
      order: [[1, 'asc']],
    
      fixedHeader: {
        header: true,
        headerOffset: 96.6   // sheet-tabs (48) + dt-top (50)
      },
    
      dom:
        "<'dt-top'l f>" +
        "<'dt-middle't>" +
        "<'dt-bottom'i p>"
    });
    
    attachSerialNumber(dt);

    dataTablesMap[sheetName] = dt;

    tablesMap[sheetName] = tableWrapper;

    const tab = document.createElement('div');
    tab.className = 'sheet-tab' + (first ? ' active' : '');
    tab.textContent = sheetName;
    tab.onclick = () => switchSheet(sheetName);
    
    // Right-aligned tabs
    if (sheetName === "Repair Cases" || sheetName === "Closed Cases Data") {
      rightTabsDiv.appendChild(tab);
    } else {
      leftTabsDiv.appendChild(tab);
    }

    first = false;
  });
}

function attachSerialNumber(dt) {
  dt.on('order.dt search.dt draw.dt', function () {
    dt.column(0, { search: 'applied', order: 'applied' })
      .nodes()
      .each((cell, i) => {
        cell.innerHTML = i + 1;
      });
  });
}

function cleanCell(value) {
  if (value === null || value === undefined) return "";

  let str = String(value);

  // Remove non-breaking spaces
  str = str.replace(/\u00A0/g, ' ');

  // Remove hidden control characters (SAP / Excel junk)
  str = str.replace(/[\u0000-\u001F\u007F]/g, '');

  return str.trim();
}

function normalizeText(val) {
  return String(val || "")
    .trim()
    .toLowerCase();
}

function toYYYYMM(dateStr) {
  const d = new Date(dateStr);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function toDateKey(dateStr) {
  const d = new Date(dateStr);
  return d.toISOString().split("T")[0];
}

function excelDateToJSDate(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);

  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  let total_seconds = Math.floor(86400 * fractional_day);

  const seconds = total_seconds % 60;
  total_seconds -= seconds;
  const hours = Math.floor(total_seconds / 3600);
  const minutes = Math.floor(total_seconds / 60) % 60;

  date_info.setHours(hours);
  date_info.setMinutes(minutes);
  date_info.setSeconds(seconds);

  return date_info;
}

function formatDate(date) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function normalizeTrackingStatus(value) {
  if (value === null || value === undefined) return "";

  let str = String(value).trim();

  // 1️⃣ Remove exact trailing ", Last Update:"
  if (str.endsWith(", Last Update:")) {
    str = str.replace(/, Last Update:$/, "");
  }

  // 2️⃣ Excel date serial → formatted date
  if (/^\d{5}\.\d+$/.test(str)) {
    const serial = Number(str);
    if (!isNaN(serial)) {
      try {
        const d = excelDateToJSDate(serial);

        const pad = n => String(n).padStart(2, "0");
        str =
          `${pad(d.getDate())}-${pad(d.getMonth() + 1)}-${d.getFullYear()} ` +
          `${pad(d.getHours())}:${pad(d.getMinutes())}`;
      } catch {
        // fallback: keep original string
      }
    }
  }

  return str;
}

function normalizeRowToSchema(row, sheetName) {
  const headers = TABLE_SCHEMAS[sheetName];
  const normalized = new Array(headers.length).fill("");

  for (let i = 0; i < headers.length; i++) {
    if (row && row[i] !== undefined) {
      normalized[i] = row[i];
    }
  }

  return normalized;
}

document.getElementById('kciInput').addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;

  if (!file.name.startsWith("KCI - Open Repair Case Data")) {
    alert("Invalid file. Please upload 'KCI - Open Repair Case Data' file.");
    e.target.value = "";
    return;
  }

  kciFile = file;
  enableProcessIfReady();
});

document.getElementById('csoInput').addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;

  if (!/^GNPro_Case_CSO_Status_\d{4}-\d{2}-\d{2}/.test(file.name)) {
    alert("Invalid GNPro CSO file name.");
    e.target.value = "";
    return;
  }

  csoFile = file;
  enableProcessIfReady();
});

document.getElementById('trackingInput').addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;

  if (!/^Tracking_Results_\d{4}-\d{2}-\d{2}/.test(file.name)) {
    alert("Invalid Tracking Results file name.");
    e.target.value = "";
    return;
  }

  trackingFile = file;
  enableProcessIfReady();
});

document.getElementById('processBtn').addEventListener('click', async () => {

  if (kciFile) {
    startProgressContext("Processing KCI Excel...");
  
    // ✅ FULL overwrite ONLY for KCI Excel
    const store = getStore("readwrite");
    ["Dump", "WO", "MO", "MO Items", "SO", "Closed Cases"].forEach(sheet => {
      dataTablesMap[sheet]?.clear().draw(false);
      store.put({
        sheetName: sheet,
        rows: [],
        lastUpdated: new Date().toISOString()
      });
    });
  
    await processExcelFile(kciFile, [
      "Dump", "WO", "MO", "MO Items", "SO", "Closed Cases"
    ]);
  
    kciFile = null;
    document.getElementById('kciInput').value = "";
    endProgressContext("KCI Excel processed");
    return;
  }

  if (csoFile) {
    startProgressContext("Processing GNPro CSO file...");
    await processGNProCSOFile(csoFile);
    csoFile = null;
    document.getElementById('csoInput').value = "";
    endProgressContext("GNPro CSO processed");
    return;
  }

  if (trackingFile) {
    startProgressContext("Processing Tracking Results...");
    await processTrackingResultsFile(trackingFile);
    trackingFile = null;
    document.getElementById('trackingInput').value = "";
    endProgressContext("Tracking Results processed");
    return;
  }

});

function enableProcessIfReady() {
  if (kciFile || csoFile || trackingFile) {
    document.getElementById('processBtn').disabled = false;
  }
}

let progressContext = null;
let displayedProgress = 0;
let progressAnimFrame = null;

function showProgressOverlay() {
  document.getElementById("progressOverlay").style.display = "flex";
}

function hideProgressOverlay() {
  const overlay = document.getElementById("progressOverlay");
  overlay.style.display = "none";
  overlay.classList.remove("progress-complete");

  document.getElementById("overlayConfirmBtn").style.display = "none";
  document.getElementById("overlayProgressBar").style.width = "0%";
  document.getElementById("overlayProgressText").textContent = "0%";
}

function startProgressContext(label) {
  progressContext = { label, value: 0 };
  displayedProgress = 0;

  showProgressOverlay();

  document.getElementById("overlayStatusText").textContent = label;
  document.getElementById("overlayProgressBar").style.width = "0%";
  document.getElementById("overlayProgressText").textContent = "0%";
}

function animateProgressTo(targetPercent, duration = 280) {
  const bar = document.getElementById("overlayProgressBar");

  if (progressAnimFrame) {
    cancelAnimationFrame(progressAnimFrame);
    progressAnimFrame = null;
  }

  const start = performance.now();
  const startPercent = displayedProgress;
  const delta = targetPercent - startPercent;

  function step(now) {
    const elapsed = now - start;
    const progress = Math.min(elapsed / duration, 1);

    // Ease-out curve (feels natural)
    const eased = 1 - Math.pow(1 - progress, 3);
    const current = startPercent + delta * eased;

    displayedProgress = current;
    bar.style.width = current.toFixed(2) + "%";
    document.getElementById("overlayProgressText").textContent =
      Math.round(current) + "%";

    if (progress < 1) {
      progressAnimFrame = requestAnimationFrame(step);
    } else {
      displayedProgress = targetPercent;
      bar.style.width = targetPercent + "%";
      document.getElementById("overlayProgressText").textContent =
        targetPercent + "%";
      progressAnimFrame = null;
    }
  }

  progressAnimFrame = requestAnimationFrame(step);
}

function updateProgressContext(current, total, text) {
  if (!progressContext || total === 0) return;

  const percent = Math.min(
    100,
    Math.round((current / total) * 100)
  );

  animateProgressTo(percent);

  if (text) {
    document.getElementById("overlayStatusText").textContent = text;
  }
}

function endProgressContext(text = "Completed") {
  const bar = document.getElementById("overlayProgressBar");
  const status = document.getElementById("overlayStatusText");
  const confirmBtn = document.getElementById("overlayConfirmBtn");
  const overlay = document.getElementById("progressOverlay");

  animateProgressTo(100, 200);
  status.textContent = text;

  overlay.classList.add("progress-complete");
  confirmBtn.style.display = "inline-block";

  confirmBtn.onclick = () => {
    hideProgressOverlay();
    progressContext = null;
  };
}

function buildSheetTables(workbook) {
  return new Promise(async resolve => {

    const sheetNames = workbook.SheetNames;

    let processedRows = 0;
    const totalRows = sheetNames.reduce((sum, s) => {
      const sheet = workbook.Sheets[s];
      return sum + (sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1 }).length : 0);
    }, 0);

    for (let index = 0; index < sheetNames.length; index++) {
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
          const excelRow = row.slice(3); // Column D onward
      
          // ✅ Skip empty Excel rows (fixes ghost rows)
          if (!excelRow.some(cell => String(cell || "").trim() !== "")) {
            return null;
          }
      
          const cleanedRow = [];
      
          for (let i = 0; i < headers.length; i++) {
            let cell = excelRow[i];
      
            if (typeof cell === "number" && cell > 40000 && cell < 60000) {
              try {
                cleanedRow.push(formatDate(excelDateToJSDate(cell)));
              } catch {
                cleanedRow.push(cleanCell(cell));
              }
            } else {
              cleanedRow.push(cleanCell(cell));
            }
          }
      
          return cleanedRow;
        })
        .filter(Boolean); // 🔥 removes null rows

      const dataTable = dataTablesMap[sheetName];
      dataTable.clear();

      rows.forEach(r => {
        dataTable.row.add(["", ...r]);
        processedRows++;
        updateProgressContext(
          processedRows,
          totalRows,
          `Processing ${sheetName} (${processedRows}/${totalRows})`
        );
      });

      dataTable.draw(false);

      getStore("readwrite").put({
        sheetName,
        rows: rows.map(r => normalizeRowToSchema(r, sheetName)),
        lastUpdated: new Date().toISOString()
      });
    }

    resolve();
  });
}

function processExcelFile(file, allowedSheets) {
  return new Promise(resolve => {
    const reader = new FileReader();

    reader.onload = async function (evt) {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const filteredWorkbook = {
        SheetNames: workbook.SheetNames.filter(s => allowedSheets.includes(s)),
        Sheets: workbook.Sheets
      };

    await buildSheetTables(filteredWorkbook);
    resolve();
    };

    reader.readAsArrayBuffer(file);
  });
}

function formatSentenceCase(value) {
  if (!value) return "";
  const str = String(value).trim().toLowerCase();
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function processCsvFile(file, targetSheetName) {
  return new Promise(resolve => {
    const reader = new FileReader();

    reader.onload = function (e) {
      const text = e.target.result;

      // Parse CSV to rows
      const rows = XLSX.utils.sheet_to_json(
        XLSX.read(text, { type: "string" }).Sheets.Sheet1,
        { header: 1, raw: true }
      );

      if (!rows.length) {
        resolve();
        return;
      }

      const headers = TABLE_SCHEMAS[targetSheetName];
      const colCount = headers.length;
      const dataRows = rows.slice(1).map(r => {
        const cleaned = [];
      
        for (let i = 0; i < headers.length; i++) {
          let cellValue = cleanCell(r[i]);
      
          // ✅ Apply sentence case ONLY for CSO Status → Status column
          if (targetSheetName === "CSO Status" && headers[i] === "Status") {
            cellValue = formatSentenceCase(cellValue);
          }
      
          cleaned.push(cellValue);
        }
      
        return cleaned;
      });

      // Update UI
      const dt = dataTablesMap[targetSheetName];
      dt.clear();
      dataRows.forEach(r => dt.row.add(["", ...r])); // S.No placeholder
      dt.draw(false);

      // Save to IndexedDB
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
  rows.slice(1).forEach(r => {
    const caseId = cleanCell(r[0]);
    if (!caseId) return;

    map.set(caseId, {
      status: formatSentenceCase(cleanCell(r[2])),
      tracking: cleanCell(r[3]),
      repairStatus: cleanCell(r[4])
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

  const dump = allData.find(r => r.sheetName === "Dump")?.rows || [];
  const so = allData.find(r => r.sheetName === "SO")?.rows || [];
  const oldCso = allData.find(r => r.sheetName === "CSO Status")?.rows || [];

  // Build GNPro CSV map
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

  // 1️⃣ Offsite cases
  const offsiteCases = [
    ...new Set(
      dump
        .filter(r => normalizeText(r[dumpResIdx]) === "offsite solution")
        .map(r => r[dumpCaseIdx])
    )
  ];

  const finalRows = [];

  let processed = 0;
  const total = offsiteCases.length;
  
  offsiteCases.forEach(caseId => {
    processed++;
    updateProgressContext(
      processed,
      total,
      `Processing CSO cases (${processed}/${total})`
    );
    const soRows = so.filter(r => r[soCaseIdx] === caseId);
    if (!soRows.length) return;

    const latestSO = soRows.sort((a, b) =>
      new Date(b[soDateIdx]) - new Date(a[soDateIdx])
    )[0];

    const cso = stripOrderSuffix(latestSO[soOrderIdx]);

    let status = "Not Found";
    let tracking = "";

    const oldRow = oldCso.find(r => r[oldCaseIdx] === caseId);
    if (oldRow) {
      const oldStatus = normalizeText(oldRow[oldStatusIdx]);
      if (
        oldStatus === "delivered" ||
        oldStatus === "order cancelled, not to be reopened"
      ) {
        status = formatSentenceCase(oldRow[oldStatusIdx]);
        tracking = oldRow[oldTrackIdx];
        finalRows.push([
          caseId,
          cso,
          status,
          tracking,
          oldRow[4] || ""   // Repair Status (if exists)
        ]);
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
    
    finalRows.push([
      caseId,
      cso,
      status,
      tracking,
      repairStatus
    ]);
  });

  // Update UI
  const dt = dataTablesMap["CSO Status"];
  dt.clear();
  finalRows.forEach(r => dt.row.add(["", ...r]));
  dt.draw(false);

  // Save to IndexedDB
  const writeStore = getStore("readwrite");
  writeStore.put({
    sheetName: "CSO Status",
    rows: finalRows,
    lastUpdated: new Date().toISOString()
  });
}

function loadDataFromDB() {
  const store = getStore("readonly");
  const request = store.getAll();

  request.onsuccess = function () {
    const records = request.result;

    records.forEach(record => {
      const sheetName = record.sheetName;
      const dt = dataTablesMap[sheetName];
      if (!dt) return;

      dt.clear();
      record.rows.forEach(row => {
        const normalized = normalizeRowToSchema(row, sheetName);
        dt.row.add(["", ...normalized]); // S.No placeholder
      });
      dt.draw(false);
    });
  };
}

async function openClosedCasesReport() {
  const store = getStore("readonly");
  const all = await new Promise(r => {
    const q = store.getAll();
    q.onsuccess = () => r(q.result);
  });

  const closed =
    all.find(x => x.sheetName === "Closed Cases Data")?.rows || [];

  if (!closed.length) {
    alert("No Closed Cases Data available.");
    return;
  }

  buildClosedCasesMonthFilter(closed);
  buildClosedCasesAgentFilter(closed);
  buildClosedCasesSummary(closed);

  openModal("closedCasesReportModal");
}

function buildClosedCasesMonthFilter(rows) {
  const select = document.getElementById("ccMonthSelect");
  select.innerHTML = "";

  const months = [...new Set(
    rows.map(r => toYYYYMM(r[6]))   // Case Closed Date
  )].sort().reverse().slice(0, 7);

  months.forEach(m => {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    select.appendChild(opt);
  });

  // Ensure latest month is selected by default
  if (select.options.length) {
    select.selectedIndex = 0;
  }

  select.onchange = () => buildClosedCasesSummary(rows);
}

function buildClosedCasesAgentFilter(rows) {
  const select = document.getElementById("ccAgentSelect");
  select.innerHTML = "";

  const agents = [...new Set(
    rows
      .map(r => r[7])   // Closed By
      .filter(v => v && v !== "CRM Auto Closed")
  )].sort();

  agents.forEach(a => {
    const opt = document.createElement("option");
    opt.value = a;
    opt.textContent = a;
    opt.selected = true; // default = all selected
    select.appendChild(opt);
  });

  select.onchange = () => buildClosedCasesSummary(rows);
}

function buildClosedCasesSummary(rows) {
  const month = document.getElementById("ccMonthSelect").value;
  const agentSelect = document.getElementById("ccAgentSelect");

  const selectedAgents = [...agentSelect.selectedOptions].map(o => o.value);

  const filtered = rows.filter(r =>
    toYYYYMM(r[6]) === month
  );

  const byDate = {};

  filtered.forEach(r => {
    const dateKey = toDateKey(r[6]);
    const closedBy = r[7];

    if (!byDate[dateKey]) {
      byDate[dateKey] = {
        total: 0,
        kci: 0,
        crm: 0
      };
    }

    byDate[dateKey].total++;

    if (closedBy === "CRM Auto Closed") {
      byDate[dateKey].crm++;
    } else if (selectedAgents.includes(closedBy)) {
      byDate[dateKey].kci++;
    }
  });

  const table = $("#ccSummaryTable").DataTable();
  table.clear();

  Object.keys(byDate).sort().forEach(date => {
    const d = byDate[date];
    const agentRC = d.total - d.kci - d.crm;

    table.row.add([
      date,
      d.total,
      d.kci > 0
      ? `<span class="cc-kci" data-date="${date}">${d.kci}</span>`
      : "0",
      d.crm,
      agentRC
    ]);
  });

  table.draw(false);

  attachDrilldownClicks(filtered);
}

function attachDrilldownClicks(rows) {
  document.querySelectorAll(".cc-kci").forEach(el => {
    el.onclick = () => {
      const date = el.dataset.date;
      buildDrilldown(rows, date);
    };
  });
}

function buildDrilldown(rows, date) {
  const agentSelect = document.getElementById("ccAgentSelect");
  const selectedAgents = [...agentSelect.selectedOptions].map(o => o.value);

  const map = {};

const selectedMonth =
  document.getElementById("ccMonthSelect").value;

  rows.forEach(r => {
    if (
      toYYYYMM(r[6]) === selectedMonth &&
      toDateKey(r[6]) === date &&
      selectedAgents.includes(r[7])
    ) {
      map[r[7]] = (map[r[7]] || 0) + 1;
    }
  });

  const table = $("#ccDrillTable").DataTable();
  table.clear();

  Object.entries(map).forEach(([agent, count]) => {
    table.row.add([agent, count]);
  });

  table.draw(false);

  document.getElementById("ccDrillTitle").textContent =
    `KCI Closures – ${date}`;
}


function switchSheet(sheetName) {
  document.querySelectorAll('.sheet-tab').forEach(tab => {
    tab.classList.toggle('active', tab.textContent === sheetName);
  });

  Object.keys(tablesMap).forEach(name => {
    const wrapper = tablesMap[name];
    const isActive = name === sheetName;

    wrapper.style.display = isActive ? 'block' : 'none';

    if (isActive) {
      const table = wrapper.querySelector('table');
      const dataTable = dataTablesMap[sheetName];

      // CRITICAL: force DataTables to recalc columns
      dataTable.columns.adjust().draw(false);
    }
  });
}

function toDateOnly(d) {
  const x = new Date(d);
  return new Date(x.getFullYear(), x.getMonth(), x.getDate());
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
  const s = normalizeText(status);

  if (s === "closed") return 1;
  if (s === "pod") return 2;
  if (s === "shipped") return 3;
  if (s === "ordered") return 4;
  if (s === "partially ordered") return 5;
  if (s === "order pending") return 6;
  if (s === "new") return 7;
  if (s === "cancelled") return 8;

  return 9; // unknown / future statuses
}

function getLatestMO(caseId, mo) {
  const caseMOs = mo.filter(r => r[1] === caseId);
  if (!caseMOs.length) return null;

  // Sort by Created On DESC
  caseMOs.sort((a, b) => new Date(b[2]) - new Date(a[2]));
  const latestTime = new Date(caseMOs[0][2]);

  // 5-minute window
  const windowMOs = caseMOs.filter(r =>
    Math.abs(new Date(r[2]) - latestTime) <= 5 * 60 * 1000
  );

  // Priority sort
  windowMOs.sort(
    (a, b) =>
      getMOStatusPriority(a[3]) - getMOStatusPriority(b[3])
  );

  return windowMOs[0];
}

function getFirstOrderDate(caseId, wo, mo, so) {
  const dates = [];

  wo.forEach(r => r[0] === caseId && r[6] && dates.push(new Date(r[6])));
  mo.forEach(r => r[1] === caseId && r[2] && dates.push(new Date(r[2])));
  so.forEach(r => r[0] === caseId && r[2] && dates.push(new Date(r[2])));

  if (!dates.length) return null;
  return new Date(Math.min(...dates));
}

function calculateSBD(caseRow, firstOrderDate, sbdConfig) {
  if (!firstOrderDate || !sbdConfig?.periods) return "NA";

  const caseCreated = new Date(caseRow.createdOn);
  const caseDateOnly = toDateOnly(caseCreated);

  const period = sbdConfig.periods.find(p =>
    p.startDate && p.endDate &&
    caseDateOnly >= new Date(p.startDate) &&
    caseDateOnly <= new Date(p.endDate)
  );

  if (!period) return "NA";

  const countryRow = period.rows.find(
    r => normalizeText(r.country) === normalizeText(caseRow.country)
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

  const dump = allData.find(r => r.sheetName === "Dump")?.rows || [];
  const so = allData.find(r => r.sheetName === "SO")?.rows || [];
  const cso = allData.find(r => r.sheetName === "CSO Status")?.rows || [];

  const dumpIdx = TABLE_SCHEMAS["Dump"].indexOf("Case Resolution Code");
  const dumpCaseIdx = TABLE_SCHEMAS["Dump"].indexOf("Case ID");

  const soCaseIdx = TABLE_SCHEMAS["SO"].indexOf("Case ID");
  const soDateIdx = TABLE_SCHEMAS["SO"].indexOf("Date and Time Submitted");
  const soOrderIdx = TABLE_SCHEMAS["SO"].indexOf("Order Reference ID");

  const csoCaseIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Case ID");
  const csoStatusIdx = TABLE_SCHEMAS["CSO Status"].indexOf("Status");

  // Step 1: Offsite cases
  const offsiteCases = [
    ...new Set(
      dump
        .filter(r => normalizeText(r[dumpIdx]) === "offsite solution")
        .map(r => r[dumpCaseIdx])
    )
  ];

  const result = [];

  offsiteCases.forEach(caseId => {
    const soRows = so.filter(r => r[soCaseIdx] === caseId);
    if (!soRows.length) return; // skip if no SO

    // latest SO by date
    const latest = soRows.sort((a, b) => {
      const da = new Date(a[soDateIdx]);
      const db = new Date(b[soDateIdx]);
      return db - da;
    })[0];

    let orderId = stripOrderSuffix(latest[soOrderIdx]);
    if (!orderId) return;

    const csoRow = cso.find(r => r[csoCaseIdx] === caseId);
    if (csoRow) {
      const status = normalizeText(csoRow[csoStatusIdx]);
      if (
        status === "delivered" ||
        status === "order cancelled, not to be reopened"
      ) {
        return; // exclude
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

  const dump = allData.find(r => r.sheetName === "Dump")?.rows || [];
  const mo = allData.find(r => r.sheetName === "MO")?.rows || [];
  const moItems = allData.find(r => r.sheetName === "MO Items")?.rows || [];
  const cso = allData.find(r => r.sheetName === "CSO Status")?.rows || [];
  const delivery = allData.find(r => r.sheetName === "Delivery Details")?.rows || [];

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

  // ---- Stage 1: MO based tracking (primary) ----
  const partsShippedCases = [
    ...new Set(
      dump
        .filter(r => normalizeText(r[dumpResIdx]) === "parts shipped")
        .map(r => r[dumpCaseIdx])
    )
  ];

  const moTrackingMap = new Map();

  partsShippedCases.forEach(caseId => {
    const caseMOs = mo.filter(r => r[moCaseIdx] === caseId);
    if (!caseMOs.length) return;

    // Sort by created date desc
    caseMOs.sort((a, b) => new Date(b[moCreatedIdx]) - new Date(a[moCreatedIdx]));

    const latestTime = new Date(caseMOs[0][moCreatedIdx]);

    // 5-minute window
    const windowMOs = caseMOs.filter(r =>
      Math.abs(new Date(r[moCreatedIdx]) - latestTime) <= 5 * 60 * 1000
    );

    // Pick by status priority
    windowMOs.sort(
      (a, b) => getMOStatusPriority(a[moStatusIdx]) - getMOStatusPriority(b[moStatusIdx])
    );

    const selected = windowMOs[0];
    const status = normalizeText(selected[moStatusIdx]);

    if (status !== "closed" && status !== "pod") return;

    const moNumber = selected[moOrderIdx];

    const item = moItems.find(r =>
      r[moItemOrderIdx] === moNumber &&
      normalizeText(r[moItemNameIdx]).endsWith("- 1")
    );

    if (!item || !item[moItemUrlIdx]) return;

    moTrackingMap.set(caseId, item[moItemUrlIdx]);
  });

  // ---- Stage 2: CSO delivered tracking (independent) ----
  const csoTrackingMap = new Map();
  
  cso.forEach(r => {
    const caseId = r[csoCaseIdx];
    const status = normalizeText(r[csoStatusIdx]);
    const tn = r[csoTrackIdx];
  
    if (status === "delivered" && tn) {
      const url =
        "http://wwwapps.ups.com/WebTracking/processInputRequest" +
        "?TypeOfInquiryNumber=T&InquiryNumber1=" + tn;
  
      csoTrackingMap.set(caseId, url);
    }
  });
  
  // ---- Stage 3: Merge MO + CSO tracking ----
  const combinedTrackingMap = new Map();
  
  // MO tracking first
  moTrackingMap.forEach((url, caseId) => {
    combinedTrackingMap.set(caseId, url);
  });
  
  // CSO tracking next
  csoTrackingMap.forEach((url, caseId) => {
    combinedTrackingMap.set(caseId, url);
  });
  
  // ---- Stage 4: Remove processed cases ----
  delivery.forEach(r => {
    const caseId = r[delCaseIdx];
    const status = normalizeText(r[delStatusIdx]);
  
    if (
      combinedTrackingMap.has(caseId) &&
      status &&
      status !== "no status found"
    ) {
      combinedTrackingMap.delete(caseId);
    }
  });
  
  // ---- Stage 5: Final output ----
  return [...combinedTrackingMap.entries()]
    .map(([caseId, url]) => `${caseId} | ${url}`)
    .join("\n");
}

function parseTrackingResultsCSV(text) {
  const rows = XLSX.utils.sheet_to_json(
    XLSX.read(text, { type: "string" }).Sheets.Sheet1,
    { header: 1, raw: true }
  );

  const map = new Map();

  // Skip header
  rows.slice(1).forEach(r => {
    const caseId = cleanCell(r[0]);   // Column A
    const status = cleanCell(r[1]);   // Column B (CurrentStatus)
    if (!caseId) return;
    map.set(caseId, status);
  });

  return map;
}

async function processTrackingResultsFile(file) {
  // 1️⃣ Load existing data
  const store = getStore("readonly");
  const allData = await new Promise(res => {
    const req = store.getAll();
    req.onsuccess = () => res(req.result);
  });

  const dump =
    allData.find(r => r.sheetName === "Dump")?.rows || [];

  const oldDelivery =
    allData.find(r => r.sheetName === "Delivery Details")?.rows || [];

  const dumpCaseIdx =
    TABLE_SCHEMAS["Dump"].indexOf("Case ID");

  const delCaseIdx =
    TABLE_SCHEMAS["Delivery Details"].indexOf("CaseID");

  const delStatusIdx =
    TABLE_SCHEMAS["Delivery Details"].indexOf("CurrentStatus");

  // 2️⃣ Parse Tracking Results CSV
  const csvText = await file.text();
  const trackingCSVMap = parseTrackingResultsCSV(csvText);
  // Map<CaseID, CurrentStatus>

  // 3️⃣ Build a map from existing Delivery Details
  const deliveryMap = new Map();

  oldDelivery.forEach(r => {
    const caseId = r[delCaseIdx];
    const status = r[delStatusIdx];
    if (!caseId) return;
    deliveryMap.set(caseId, status);
  });

  // 4️⃣ Merge / update using Tracking Results CSV
  let processed = 0;
  const total = trackingCSVMap.size;

  trackingCSVMap.forEach((status, caseId) => {
    processed++;
    updateProgressContext(
      processed,
      total,
      `Updating Delivery Details (${processed}/${total})`
    );

    // Update existing or add new
    deliveryMap.set(caseId, normalizeTrackingStatus(status));
  });

  // 5️⃣ Cleanup: remove cases NOT present in Dump
  const validDumpCases = new Set(
    dump.map(r => r[dumpCaseIdx])
  );

  [...deliveryMap.keys()].forEach(caseId => {
    if (!validDumpCases.has(caseId)) {
      deliveryMap.delete(caseId);
    }
  });

  // 6️⃣ Build final Delivery Details rows
  const finalRows = [...deliveryMap.entries()].map(
    ([caseId, status]) => [caseId, status]
  );

  // 7️⃣ Update UI table
  const dt = dataTablesMap["Delivery Details"];
  dt.clear();
  finalRows.forEach(r => dt.row.add(["", ...r]));
  dt.draw(false);

  // 8️⃣ Save to IndexedDB
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
  return aStart && aEnd && bStart && bEnd &&
    !(aEnd < bStart || bEnd < aStart);
}

function renderSbdModal(data) {
  const container = document.querySelector(".sbd-periods");
  container.innerHTML = "";

  data.periods.forEach((p, idx) => {
    const div = document.createElement("div");
    div.className = "sbd-period";

    div.innerHTML = `
      <h4>Date Period ${idx + 1}</h4>
      <div class="sbd-dates">
        <input type="date" class="sbd-start" value="${p.startDate}">
        <input type="date" class="sbd-end" value="${p.endDate}">
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
      <div class="sbd-add">+ Add Row</div>
    `;

    const tbody = div.querySelector("tbody");

    function addRow(row = { country: "", time: "" }) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><input value="${row.country}"></td>
        <td><input type="time" value="${row.time}"></td>
        <td><button class="del">✕</button></td>
      `;
      tr.querySelector(".del").onclick = () => tr.remove();
      tbody.appendChild(tr);
    }

    p.rows.forEach(addRow);
    div.querySelector(".sbd-add").onclick = () => addRow();

    container.appendChild(div);
  });
}

async function saveSbdData() {
  const periods = [];
  const blocks = document.querySelectorAll(".sbd-period");

  blocks.forEach(b => {
    const startDate = b.querySelector(".sbd-start").value;
    const endDate = b.querySelector(".sbd-end").value;

    const rows = [];
    b.querySelectorAll("tbody tr").forEach(tr => {
      const country = tr.children[0].querySelector("input").value.trim();
      const time = tr.children[1].querySelector("input").value;
      if (country && time) rows.push({ country, time });
    });

    periods.push({ startDate, endDate, rows });
  });

  // Date overlap validation
  for (let i = 0; i < periods.length; i++) {
    for (let j = i + 1; j < periods.length; j++) {
      if (datesOverlap(
        periods[i].startDate,
        periods[i].endDate,
        periods[j].startDate,
        periods[j].endDate
      )) {
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

document.getElementById("sbdBtn").onclick = async () => {
  const store = getStore("readonly");
  const req = store.get("SBD Cut Off Times");

  req.onsuccess = () => {
    const data = req.result || createEmptySbdData();
    renderSbdModal(data);
    document.getElementById("sbdModal").style.display = "flex";
  };
};

document.getElementById("saveSbdBtn").onclick = saveSbdData;
document.getElementById("closeSbdBtn").onclick =
  () => document.getElementById("sbdModal").style.display = "none";

function adjustMappingModalWidth(containerId, modalId) {
  const container = document.getElementById(containerId);
  const modal = document.querySelector(`#${modalId} .modal`);

  const blocks = container.querySelectorAll(".mapping-block:not(.placeholder)");
  const count = Math.max(1, Math.min(blocks.length, 3)); // 1–3 per row

  const blockWidth = 320;
  const gap = 12;

  const width = count * blockWidth + (count - 1) * gap;
  modal.style.width = width + "px";
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
      <button onclick="this.closest('.mapping-block').remove()">✕</button>
    </h4>
    <div class="agents"></div>
    <button class="add-btn">+ Add Agent</button>
  `;

  const agentsDiv = block.querySelector(".agents");
  const addAgentBtn = block.querySelector(".add-btn");

  function addAgent(val = "") {
    const row = document.createElement("div");
    row.className = "mapping-row";
    row.innerHTML = `
      <input value="${val}">
      <button onclick="this.parentElement.remove()">✕</button>
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
    addMarketBlock();                   // real empty block
    container.firstChild.classList.add("placeholder");
  } else {
    data.forEach(m => addMarketBlock(m.name, m.countries));
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
      <button onclick="this.closest('.mapping-block').remove()">✕</button>
    </h4>
    <div class="countries"></div>
    <button class="add-btn">+ Add Country</button>
  `;

  const div = block.querySelector(".countries");
  const addBtn = block.querySelector(".add-btn");

  function addCountry(val = "") {
    const row = document.createElement("div");
    row.className = "mapping-row";
    row.innerHTML = `
      <input value="${val}">
      <button onclick="this.parentElement.remove()">✕</button>
    `;
    div.appendChild(row);
  }

  countries.forEach(addCountry);
  addBtn.onclick = () => addCountry();

  container.appendChild(block);
  adjustMappingModalWidth("marketContainer", "marketModal");
}

document.getElementById("saveTlBtn").onclick = () => {
  const blocks = document.querySelectorAll("#tlContainer .mapping-block");
  const data = [];

  blocks.forEach(b => {
    const name = b.querySelector("h4 input").value.trim();
    if (!name) return;

    const agents = [...b.querySelectorAll(".mapping-row input")]
      .map(i => i.value.trim())
      .filter(Boolean);

    data.push({ name, agents });
  });

  getStore("readwrite").put({ sheetName: "TL_MAP", data });
  closeModal("tlModal");
};

document.getElementById("saveMarketBtn").onclick = () => {
  const blocks = document.querySelectorAll("#marketContainer .mapping-block");
  const data = [];

  blocks.forEach(b => {
    const name = b.querySelector("h4 input").value.trim();
    if (!name) return;

    const countries = [...b.querySelectorAll(".mapping-row input")]
      .map(i => i.value.trim())
      .filter(Boolean);

    data.push({ name, countries });
  });

  getStore("readwrite").put({ sheetName: "MARKET_MAP", data });
  closeModal("marketModal");
};

document.getElementById("tlBtn").onclick = async () => {
  const req = getStore().get("TL_MAP");
  req.onsuccess = () => {
    renderTLModal(req.result?.data || []);
    openModal("tlModal");
  };
};

document.getElementById("marketBtn").onclick = async () => {
  const req = getStore().get("MARKET_MAP");
  req.onsuccess = () => {
    renderMarketModal(req.result?.data || []);
    openModal("marketModal");
  };
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

document.getElementById("closedCasesReportBtn")
  .addEventListener("click", openClosedCasesReport);

document.getElementById("copySoBtn").addEventListener("click", async () => {
  const output = await buildCopySOOrders();

  const lines = output
    ? output.split("\n").filter(l => l.trim() !== "")
    : [];

  document.getElementById("soOutput").value =
    lines.length ? lines.join("\n") : "No eligible cases found.";

  document.getElementById("soCount").textContent =
    `Total cases: ${lines.length}`;

  document.getElementById("soModalTitle").textContent =
    "Copy SO Orders Preview";

  document.getElementById("soModal").style.display = "flex";
});

document.getElementById("copyTrackingBtn").addEventListener("click", async () => {
  const output = await buildCopyTrackingURLs();

  const lines = output
    ? output.split("\n").filter(l => l.trim())
    : [];

  document.getElementById("soOutput").value =
    lines.length ? lines.join("\n") : "No eligible tracking URLs found.";

  document.getElementById("soCount").textContent =
    `Total cases: ${lines.length}`;

  document.getElementById("soModalTitle").textContent =
    "Copy Tracking URLs Preview";

  document.getElementById("soModal").style.display = "flex";
});

document.getElementById("closeModalBtn").addEventListener("click", () => {
  document.getElementById("soModal").style.display = "none";
});

document.getElementById("copyToClipboardBtn").addEventListener("click", async () => {
  const text = document.getElementById("soOutput").value;
  await navigator.clipboard.writeText(text);
  alert("Copied to clipboard");
});

async function buildRepairCases() {
  const store = getStore("readonly");
  const all = await new Promise(r => {
    const q = store.getAll();
    q.onsuccess = () => r(q.result);
  });

  const dump = all.find(x => x.sheetName === "Dump")?.rows || [];
  const wo = all.find(x => x.sheetName === "WO")?.rows || [];
  const mo = all.find(x => x.sheetName === "MO")?.rows || [];
  const moItems = all.find(x => x.sheetName === "MO Items")?.rows || [];
  const so = all.find(x => x.sheetName === "SO")?.rows || [];
  const cso = all.find(x => x.sheetName === "CSO Status")?.rows || [];
  const delivery = all.find(x => x.sheetName === "Delivery Details")?.rows || [];

  const tlMap = all.find(x => x.sheetName === "TL_MAP")?.data || [];
  const marketMap = all.find(x => x.sheetName === "MARKET_MAP")?.data || [];
  const sbdConfig = all.find(x => x.sheetName === "SBD Cut Off Times");

  const validCases = dump.filter(r =>
    ["parts shipped", "onsite solution", "offsite solution"]
      .includes(normalizeText(r[8]))
  );

  const rows = [];

  validCases.forEach(d => {
    const caseId = d[0];
    const resolution = normalizeText(d[8]);

    const firstOrder = getFirstOrderDate(caseId, wo, mo, so);

    const tl =
      tlMap.find(t =>
        t.agents.some(a => normalizeText(a) === normalizeText(d[9]))
      )?.name || "";

    const market =
      marketMap.find(m =>
        m.countries.some(c => normalizeText(c) === normalizeText(d[6]))
      )?.name || "";

    const onsiteRFC =
      resolution === "onsite solution"
        ? (wo.filter(w => w[0] === caseId)
            .sort((a, b) => new Date(b[6]) - new Date(a[6]))[0]?.[5] || "")
        : "Not Found";

    const csrRFC =
      resolution === "parts shipped"
        ? (getLatestMO(caseId, mo)?.[3] || "")
        : "Not Found";

    const benchRFC =
      resolution === "offsite solution"
        ? (cso.find(c => c[0] === caseId)?.[2] || "")
        : "Not Found";

    const dnap =
      resolution === "offsite solution" &&
      normalizeText(cso.find(c => c[0] === caseId)?.[4])
        .includes("product returned unrepaired to customer")
        ? "True"
        : "";

    let partNumber = "";
    let partName = "";
    
    if (resolution === "parts shipped") {
      const latestMO = getLatestMO(caseId, mo);
    
      if (latestMO) {
        const moNumber = latestMO[0];
    
        const item = moItems.find(r =>
          r[0] === moNumber &&
          normalizeText(r[1]).endsWith("- 1")
        );
    
        if (item) {
          partNumber = item[2] || "";
    
          if (item[3]) {
            const idx = item[3].indexOf("-");
            partName = idx >= 0
              ? item[3].substring(idx + 1).trim()
              : item[3].trim();
          }
        }
      }
    }

    rows.push([
      caseId,
      d[1], d[2], d[3], d[6], d[8], d[9], d[15],
      getCAGroup(d[2]),
      tl,
      calculateSBD({ createdOn: d[2], country: d[6] }, firstOrder, sbdConfig),
      onsiteRFC,
      csrRFC,
      benchRFC,
      market,
      onsiteRFC !== "Not Found"
        ? (
            wo.filter(w => w[0] === caseId)
              .sort((a, b) => new Date(b[6]) - new Date(a[6]))[0]
              ?.[9] || ""
          )
        : "",
      delivery.find(x => x[0] === caseId)?.[1] || "",
      partNumber,
      partName,
      d[16],
      d[17],
      d[10],
      dnap
    ]);
  });

  const dt = dataTablesMap["Repair Cases"];
  dt.clear();
  rows.forEach(r => dt.row.add(["", ...r]));
  dt.draw(false);

  getStore("readwrite").put({
    sheetName: "Repair Cases",
    rows,
    lastUpdated: new Date().toISOString()
  });
}

async function buildClosedCasesReport() {
  const store = getStore("readonly");
  const all = await new Promise(r => {
    const q = store.getAll();
    q.onsuccess = () => r(q.result);
  });

  const closed = all.find(x => x.sheetName === "Closed Cases")?.rows || [];
  const repair = all.find(x => x.sheetName === "Repair Cases")?.rows || [];

  const now = new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth() - 6, 1);

  const rows = [];

  closed.forEach(c => {
    const closedDate = new Date(c[4]);
    if (closedDate < cutoff) return;

    const repairRow = repair.find(r => r[0] === c[0]);

    let closedBy = c[2];
    if (closedBy === "# CrmWebJobUser-Prod") closedBy = "CRM Auto Closed";
    else if (
      ["# MSFT-ServiceSystemAdmin",
       "# CrmEEGUser-Prod",
       "# MSFT-ServiceSystemAdminDev",
       "SYSTEM"].includes(closedBy)
    ) closedBy = c[8];

    rows.push([
      c[0],
      repairRow?.[1] || "",
      c[1], c[6], c[2], c[3], c[4],
      closedBy,
      c[9], c[5], c[8], c[10],
      repairRow?.[9] || "",
      repairRow?.[10] || "",
      repairRow?.[14] || ""
    ]);
  });

  const dt = dataTablesMap["Closed Cases Data"];
  dt.clear();
  rows.forEach(r => dt.row.add(["", ...r]));
  dt.draw(false);

  const write = getStore("readwrite");
  write.put({
    sheetName: "Closed Cases Data",
    rows,
    lastUpdated: new Date().toISOString()
  });

  // Remove closed cases from Repair Cases
  const remaining = repair.filter(
    r => !rows.some(c => c[0] === r[0])
  );

  write.put({ sheetName: "Repair Cases", rows: remaining });
}

document.getElementById("processRepairBtn")
  .addEventListener("click", async () => {

    startProgressContext("Building Repair Cases...");

    const store = getStore("readonly");
    const all = await new Promise(r => {
      const q = store.getAll();
      q.onsuccess = () => r(q.result);
    });

    const dump = all.find(x => x.sheetName === "Dump")?.rows || [];
    const validCases = dump.filter(r =>
      ["parts shipped", "onsite solution", "offsite solution"]
        .includes(normalizeText(r[8]))
    );

    let processed = 0;
    const total = validCases.length;

    for (const _ of validCases) {
      processed++;
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

document.addEventListener('DOMContentLoaded', async () => {
  await openDB();
  initEmptyTables();

  // IMPORTANT: wait for DataTables to fully initialize
  requestAnimationFrame(() => {
    loadDataFromDB();
    $('#ccSummaryTable').DataTable({
      paging: false,
      searching: false,
      info: false
    });
    
    $('#ccDrillTable').DataTable({
      paging: false,
      searching: false,
      info: false
    });
  });
});

const themeToggle = document.getElementById('themeToggle');

function setTheme(theme) {
  document.body.setAttribute('data-theme', theme);
  localStorage.setItem('kci-theme', theme);
  themeToggle.textContent = theme === 'dark' ? '🌙' : '☀️';
}

themeToggle.addEventListener('click', () => {
  const current = document.body.getAttribute('data-theme') || 'dark';
  setTheme(current === 'dark' ? 'light' : 'dark');
});

// Init theme on load
const savedTheme = localStorage.getItem('kci-theme') || 'dark';
setTheme(savedTheme);

document.addEventListener("keydown", (e) => {
  if (e.key !== "Enter") return;

  const overlay = document.getElementById("progressOverlay");
  const confirmBtn = document.getElementById("overlayConfirmBtn");

  if (
    overlay.style.display === "flex" &&
    confirmBtn.style.display !== "none"
  ) {
    e.preventDefault();
    confirmBtn.click();
  }
});














