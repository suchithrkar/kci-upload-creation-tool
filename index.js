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
  ]
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

function getStore(mode = "readonly") {
  const tx = db.transaction(STORE_NAME, mode);
  return tx.objectStore(STORE_NAME);
}

function initEmptyTables() {
  const container = document.getElementById('tablesContainer');
  container.innerHTML = '';

  const tabsDiv = document.createElement('div');
  tabsDiv.className = 'sheet-tabs';
  container.appendChild(tabsDiv);

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
      th.textContent = h;
      tr.appendChild(th);
    });

    thead.appendChild(tr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    table.appendChild(tbody);

    tableWrapper.appendChild(table);
    container.appendChild(tableWrapper);

    const dt = $(table).DataTable({
      pageLength: 25,
      autoWidth: true,
      columnDefs: [
        {
          targets: 0,
          searchable: false,
          orderable: false
        }
      ],
      order: [[1, 'asc']]
    });
    attachSerialNumber(dt);

    dataTablesMap[sheetName] = dt;

    tablesMap[sheetName] = tableWrapper;

    const tab = document.createElement('div');
    tab.className = 'sheet-tab' + (first ? ' active' : '');
    tab.textContent = sheetName;
    tab.onclick = () => switchSheet(sheetName);
    tabsDiv.appendChild(tab);

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

  // - Remove hidden control characters (SAP / Excel junk)
  str = str.replace(/[\u0000-\u001F\u007F]/g, '');

  return str.trim();
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

document.getElementById('processBtn').addEventListener('click', async function () {
  document.getElementById('progressBarContainer').style.display = 'block';
  document.getElementById('progressBar').style.width = '0%';

  if (kciFile) {
    await processExcelFile(
      kciFile,
      ["Dump", "WO", "MO", "MO Items", "SO", "Closed Cases"]
    );
    kciFile = null;
    document.getElementById('kciInput').value = "";
  }

  if (csoFile) {
    await processGNProCSOFile(csoFile);
    csoFile = null;
    document.getElementById('csoInput').value = "";
  }
  
  if (trackingFile) {
    await processTrackingResultsFile(trackingFile);
    trackingFile = null;
    document.getElementById('trackingInput').value = "";
  }

  document.getElementById('statusText').textContent = "Upload processed successfully.";
  document.getElementById('processBtn').disabled = true;
});

function enableProcessIfReady() {
  if (kciFile || csoFile || trackingFile) {
    document.getElementById('processBtn').disabled = false;
    document.getElementById('statusText').textContent = "Files ready to process.";
  }
}

function buildSheetTables(workbook) {

  const progressBar = document.getElementById('progressBar');
  const statusText = document.getElementById('statusText');

  const sheetNames = workbook.SheetNames;
  let processed = 0;

  (async function processSheets() {
    for (let index = 0; index < sheetNames.length; index++) {
      const sheetName = sheetNames[index];
      await new Promise(requestAnimationFrame);
  
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  
      if (json.length === 0) continue;
  
      const headers = TABLE_SCHEMAS[sheetName];
      if (!headers) continue; // ignore unknown sheets safely
      const rows = json.slice(1).map(row => {
        const cleanedRow = [];
        const excelRow = row.slice(3);
      
        for (let i = 0; i < headers.length; i++) {
          let cell = excelRow[i];
      
          if (typeof cell === 'number' && cell > 40000 && cell < 60000) {
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
      });
  
      const tableWrapper = tablesMap[sheetName];
      if (!tableWrapper) continue;
      
      const table = tableWrapper.querySelector('table');
      const dataTable = dataTablesMap[sheetName];
      dataTable.clear();
      
      rows.forEach(r => {
        dataTable.row.add(["", ...r]); // S.No placeholder
      });
      
      dataTable.draw(false);

      const store = getStore("readwrite");

      store.put({
        sheetName: sheetName,
        rows: rows.map(r => normalizeRowToSchema(r, sheetName)),
        lastUpdated: new Date().toISOString()
      });
      
      processed++;
      const percent = Math.round((processed / sheetNames.length) * 100);
      progressBar.style.width = percent + '%';
    }
  
    statusText.textContent = 'Processing complete';
    document.getElementById('processBtn').disabled = true;
  })();
}

function processExcelFile(file, allowedSheets) {
  return new Promise(resolve => {
    const reader = new FileReader();

    reader.onload = function (evt) {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const filteredWorkbook = {
        SheetNames: workbook.SheetNames.filter(s => allowedSheets.includes(s)),
        Sheets: workbook.Sheets
      };

      buildSheetTables(filteredWorkbook);
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

  offsiteCases.forEach(caseId => {
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
    if (csvRow) {
      status = csvRow.status;
      tracking = csvRow.tracking;
    }

    finalRows.push([
      caseId,
      cso,
      status,
      tracking,
      csvRow.repairStatus || ""
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
        dt.row.add(["", ...row]); // S.No placeholder
      });
      dt.draw(false);
    });
  };
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

function normalizeText(val) {
  return String(val || "").trim().toLowerCase();
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

  const trackingMap = new Map();

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

    trackingMap.set(caseId, item[moItemUrlIdx]);
  });

  // ---- Stage 2: CSO delivered (secondary) ----
  cso.forEach(r => {
    const caseId = r[csoCaseIdx];
    if (trackingMap.has(caseId)) return;

    if (normalizeText(r[csoStatusIdx]) === "delivered") {
      const tn = r[csoTrackIdx];
      if (!tn) return;

      const url =
        "http://wwwapps.ups.com/WebTracking/processInputRequest" +
        "?TypeOfInquiryNumber=T&InquiryNumber1=" + tn;

      trackingMap.set(caseId, url);
    }
  });

  // ---- Stage 3: remove processed cases ----
  delivery.forEach(r => {
    const caseId = r[delCaseIdx];
    const status = normalizeText(r[delStatusIdx]);

    if (
      trackingMap.has(caseId) &&
      status &&
      status !== "no status found"
    ) {
      trackingMap.delete(caseId);
    }
  });

  return [...trackingMap.entries()]
    .map(([k, v]) => `${k} | ${v}`)
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
  const store = getStore("readonly");
  const allData = await new Promise(res => {
    const req = store.getAll();
    req.onsuccess = () => res(req.result);
  });

  const dump = allData.find(r => r.sheetName === "Dump")?.rows || [];
  const mo = allData.find(r => r.sheetName === "MO")?.rows || [];
  const moItems = allData.find(r => r.sheetName === "MO Items")?.rows || [];
  const cso = allData.find(r => r.sheetName === "CSO Status")?.rows || [];
  const oldDelivery = allData.find(r => r.sheetName === "Delivery Details")?.rows || [];

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

  // --- Parse Tracking Results CSV ---
  const csvText = await file.text();
  const trackingCSVMap = parseTrackingResultsCSV(csvText);

  // --- STEP 1A: Parts Shipped cases (MO Stage-1 logic) ---
  const partsShippedCases = [
    ...new Set(
      dump
        .filter(r => normalizeText(r[dumpResIdx]) === "parts shipped")
        .map(r => r[dumpCaseIdx])
    )
  ];

  const finalCaseSet = new Set();

  partsShippedCases.forEach(caseId => {
    const caseMOs = mo.filter(r => r[moCaseIdx] === caseId);
    if (!caseMOs.length) return;

    caseMOs.sort((a, b) => new Date(b[moCreatedIdx]) - new Date(a[moCreatedIdx]));
    const latestTime = new Date(caseMOs[0][moCreatedIdx]);

    const windowMOs = caseMOs.filter(r =>
      Math.abs(new Date(r[moCreatedIdx]) - latestTime) <= 5 * 60 * 1000
    );

    windowMOs.sort(
      (a, b) => getMOStatusPriority(a[moStatusIdx]) - getMOStatusPriority(b[moStatusIdx])
    );

    const selected = windowMOs[0];
    const status = normalizeText(selected[moStatusIdx]);

    if (status !== "closed" && status !== "pod") return;

    // Ensure MO has tracking URL (-1 only)
    const moNumber = selected[moOrderIdx];
    const item = moItems.find(r =>
      r[moItemOrderIdx] === moNumber &&
      normalizeText(r[moItemNameIdx]).endsWith("- 1") &&
      r[moItemUrlIdx]
    );

    if (!item) return;

    finalCaseSet.add(caseId);
  });

  // --- STEP 1B: CSO Delivered cases ---
  cso.forEach(r => {
    if (normalizeText(r[csoStatusIdx]) === "delivered") {
      finalCaseSet.add(r[csoCaseIdx]);
    }
  });

  // --- STEP 2: Build Delivery Details rows ---
  const finalRows = [];

  finalCaseSet.forEach(caseId => {
    // Priority 1: existing Delivery Details
    const oldRow = oldDelivery.find(r => r[oldDelCaseIdx] === caseId);
    if (oldRow && oldRow[oldDelStatusIdx]) {
      finalRows.push([caseId, oldRow[oldDelStatusIdx]]);
      return;
    }

    // Priority 2: Tracking Results CSV
    if (trackingCSVMap.has(caseId)) {
      finalRows.push([caseId, trackingCSVMap.get(caseId)]);
      return;
    }

    // Fallback
    finalRows.push([caseId, "No Status Found"]);
  });

  // --- Update UI ---
  const dt = dataTablesMap["Delivery Details"];
  dt.clear();
  finalRows.forEach(r => dt.row.add(["", ...r]));
  dt.draw(false);

  // --- Save to IndexedDB ---
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
      <h4>Period ${idx + 1}</h4>
      <div class="sbd-dates">
        <input type="date" class="sbd-start" value="${p.startDate}">
        <input type="date" class="sbd-end" value="${p.endDate}">
      </div>
      <table class="sbd-table">
        <thead>
          <tr><th>Country</th><th>Cut Off Time</th><th></th></tr>
        </thead>
        <tbody></tbody>
      </table>
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

document.getElementById("copySoBtn").addEventListener("click", async () => {
  const output = await buildCopySOOrders();

  const lines = output
    ? output.split("\n").filter(l => l.trim() !== "")
    : [];

  document.getElementById("soOutput").value =
    lines.length ? lines.join("\n") : "No eligible cases found.";

  document.getElementById("soCount").textContent =
    `Total cases: ${lines.length}`;

  document.querySelector("#soModal h3").textContent =
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

  document.querySelector("#soModal h3").textContent =
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

document.addEventListener('DOMContentLoaded', async () => {
  await openDB();
  initEmptyTables();

  // IMPORTANT: wait for DataTables to fully initialize
  requestAnimationFrame(() => {
    loadDataFromDB();
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




























