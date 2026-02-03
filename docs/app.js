import { loadExcelFile } from "../src/loaders/excelLoader.js";
import { loadCsvFile } from "../src/loaders/csvLoader.js";
import { buildRepairCases } from "../src/engine/repairCases.js";
import { mapDumpRow } from "../src/schemas/dump.js";
import { mapWoRow } from "../src/schemas/wo.js";
import { mapMoRow } from "../src/schemas/mo.js";
import { mapMoItemsRow } from "../src/schemas/moItems.js";
import { mapSoRow } from "../src/schemas/so.js";
import { mapCsoRow } from "../src/schemas/cso.js";
import { mapDeliveryRow } from "../src/schemas/delivery.js";

const output = document.querySelector("#output");
const state = window.appState;

const buttons = {
  processRepairCases: document.querySelector("#processRepairCases"),
  processClosedCases: document.querySelector("#processClosedCases"),
  copySoOrders: document.querySelector("#copySoOrders"),
  copyTrackingUrls: document.querySelector("#copyTrackingUrls"),
};

const setOutput = (message) => {
  output.textContent = message;
};

const schemaBySheet = {
  dump: mapDumpRow,
  wo: mapWoRow,
  mo: mapMoRow,
  moitems: mapMoItemsRow,
  so: mapSoRow,
  cso: mapCsoRow,
  delivery: mapDeliveryRow,
};

const normalizeSheetKey = (sheetName) =>
  sheetName.toLowerCase().replace(/\s+/g, "");

const mapSheetRows = (sheetName, rows) => {
  const mapper = schemaBySheet[normalizeSheetKey(sheetName)];

  if (!mapper) {
    return rows;
  }

  return rows.map(mapper);
};

const mapWorkbookData = (workbookData) =>
  Object.entries(workbookData).reduce((acc, [sheetName, rows]) => {
    acc[sheetName] = mapSheetRows(sheetName, rows);
    return acc;
  }, {});

const getErrorMessage = (error) =>
  error instanceof Error ? error.message : String(error || "Unknown error");

const hasRows = (rows) => Array.isArray(rows) && rows.length > 0;

const hasNormalizedExcel = () =>
  state.kciExcel?.normalized &&
  Object.keys(state.kciExcel.normalized).length > 0;

const updateButtonState = () => {
  const excelReady = hasNormalizedExcel();
  const csoReady = hasRows(state.csoCsv.data);
  const trackingReady = hasRows(state.trackingCsv.data);
  const casesReady =
    hasRows(state.repairCases) || hasRows(state.closedCases);

  buttons.processRepairCases.disabled = !(excelReady && csoReady);
  buttons.processClosedCases.disabled = !(excelReady && csoReady);
  buttons.copySoOrders.disabled = !casesReady;
  buttons.copyTrackingUrls.disabled = !(casesReady && trackingReady);
};

const handleExcelLoad = async (file) => {
  state.kciExcel.file = file || null;
  const rawData = await loadExcelFile(file);
  state.kciExcel.raw = rawData;
  state.kciExcel.normalized = mapWorkbookData(rawData);
  const sheetCounts = Object.values(state.kciExcel.normalized).map(
    (rows) => rows.length
  );
  const totalRows = sheetCounts.reduce((sum, count) => sum + count, 0);
  console.log("KCI Excel loaded.", {
    sheets: Object.keys(state.kciExcel.normalized).length,
    totalRows,
  });
  setOutput("KCI Excel loaded.");
  updateButtonState();
};

const handleCsvLoad = async (file, targetKey, label, mapper) => {
  state[targetKey].file = file || null;
  const rawData = await loadCsvFile(file);
  state[targetKey].data = mapper ? rawData.map(mapper) : rawData;
  console.log(`${label} loaded.`, { rows: state[targetKey].data.length });
  setOutput(`${label} loaded.`);
  updateButtonState();
};

const buildClosedCases = (_normalizedData, csoData) =>
  Array.isArray(csoData) ? csoData.map((row) => ({ ...row })) : [];

const collectSoOrders = (repairCases, closedCases) => {
  const orders = new Set();
  const addOrder = (value) => {
    if (value === null || value === undefined) {
      return;
    }
    const normalized = String(value).trim();
    if (normalized) {
      orders.add(normalized);
    }
  };

  repairCases.forEach((repairCase) => {
    (repairCase.relatedOrders?.so || []).forEach(addOrder);
    addOrder(repairCase.so);
    addOrder(repairCase.soId);
  });

  closedCases.forEach((row) => {
    addOrder(row.so);
    addOrder(row.soId);
    addOrder(row.salesOrder);
    addOrder(row.salesOrderNumber);
  });

  return Array.from(orders);
};

const collectTrackingUrls = (trackingRows) => {
  const urls = new Set();
  const keys = ["trackingUrl", "trackingURL", "url", "trackingLink", "link"];

  trackingRows.forEach((row) => {
    const match = keys.find((key) => row && row[key]);
    if (!match) {
      return;
    }
    const normalized = String(row[match]).trim();
    if (normalized) {
      urls.add(normalized);
    }
  });

  return Array.from(urls);
};

document.querySelector("#kciExcel").addEventListener("change", (event) => {
  handleExcelLoad(event.target.files?.[0]).catch((error) => {
    setOutput(`Failed to load KCI Excel: ${getErrorMessage(error)}`);
  });
});

document.querySelector("#csoCsv").addEventListener("change", (event) => {
  handleCsvLoad(event.target.files?.[0], "csoCsv", "CSO CSV", mapCsoRow).catch(
    (error) => {
      setOutput(`Failed to load CSO CSV: ${getErrorMessage(error)}`);
    }
  );
});

document
  .querySelector("#trackingCsv")
  .addEventListener("change", (event) => {
    handleCsvLoad(
      event.target.files?.[0],
      "trackingCsv",
      "Tracking CSV",
      mapDeliveryRow
    ).catch((error) => {
      setOutput(`Failed to load Tracking CSV: ${getErrorMessage(error)}`);
    });
  });

document
  .querySelector("#processRepairCases")
  .addEventListener("click", () => {
    if (!hasNormalizedExcel()) {
      setOutput("Please load the KCI Excel file first.");
      return;
    }
    if (!hasRows(state.csoCsv.data)) {
      setOutput("Please load the CSO CSV file first.");
      return;
    }

    try {
      const repairCases = buildRepairCases(state.kciExcel.normalized);
      state.repairCases = repairCases;
      setOutput(`Repair cases processed: ${repairCases.length}`);
      updateButtonState();
    } catch (error) {
      setOutput(`Failed to process repair cases: ${getErrorMessage(error)}`);
    }
  });

document
  .querySelector("#processClosedCases")
  .addEventListener("click", () => {
    if (!hasNormalizedExcel()) {
      setOutput("Please load the KCI Excel file first.");
      return;
    }
    if (!hasRows(state.csoCsv.data)) {
      setOutput("Please load the CSO CSV file first.");
      return;
    }

    try {
      const closedCases = buildClosedCases(
        state.kciExcel.normalized,
        state.csoCsv.data
      );
      state.closedCases = closedCases;
      setOutput(`Closed cases processed: ${closedCases.length}`);
      updateButtonState();
    } catch (error) {
      setOutput(`Failed to process closed cases: ${getErrorMessage(error)}`);
    }
  });

document.querySelector("#copySoOrders").addEventListener("click", () => {
  const orders = collectSoOrders(state.repairCases, state.closedCases);
  state.copySoOrders = orders;
  setOutput(`SO orders ready: ${orders.length}`);
  console.log("SO orders prepared.", { rows: orders.length });
  updateButtonState();
});

document.querySelector("#copyTrackingUrls").addEventListener("click", () => {
  if (!hasRows(state.trackingCsv.data)) {
    setOutput("Please load the Tracking CSV file first.");
    return;
  }

  const urls = collectTrackingUrls(state.trackingCsv.data);
  state.copyTrackingUrls = urls;
  setOutput(`Tracking URLs ready: ${urls.length}`);
  console.log("Tracking URLs prepared.", { rows: urls.length });
  updateButtonState();
});

updateButtonState();
