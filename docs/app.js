import { loadExcelFile } from "../src/loaders/excelLoader.js";
import { loadCsvFile } from "../src/loaders/csvLoader.js";
import { mapDumpRow } from "../src/schemas/dump.js";
import { mapWoRow } from "../src/schemas/wo.js";
import { mapMoRow } from "../src/schemas/mo.js";
import { mapMoItemsRow } from "../src/schemas/moItems.js";
import { mapSoRow } from "../src/schemas/so.js";
import { mapCsoRow } from "../src/schemas/cso.js";
import { mapDeliveryRow } from "../src/schemas/delivery.js";

const output = document.querySelector("#output");
const state = window.appState;

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

const handleExcelLoad = async (file) => {
  state.kciExcelFile = file || null;
  const rawData = await loadExcelFile(file);
  state.kciExcelData = mapWorkbookData(rawData);
  setOutput("KCI Excel loaded.");
};

const handleCsvLoad = async (file, targetKey, label, mapper) => {
  state[targetKey].file = file || null;
  const rawData = await loadCsvFile(file);
  state[targetKey].data = mapper ? rawData.map(mapper) : rawData;
  setOutput(`${label} loaded.`);
};

document.querySelector("#kciExcel").addEventListener("change", (event) => {
  handleExcelLoad(event.target.files?.[0]).catch((error) => {
    setOutput(`Failed to load KCI Excel: ${error.message}`);
  });
});

document.querySelector("#csoCsv").addEventListener("change", (event) => {
  handleCsvLoad(event.target.files?.[0], "csoCsv", "CSO CSV", mapCsoRow).catch(
    (error) => {
      setOutput(`Failed to load CSO CSV: ${error.message}`);
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
      setOutput(`Failed to load Tracking CSV: ${error.message}`);
    });
  });

document
  .querySelector("#processRepairCases")
  .addEventListener("click", () => {
    setOutput("Process Repair Cases clicked.");
  });

document
  .querySelector("#processClosedCases")
  .addEventListener("click", () => {
    setOutput("Process Closed Cases clicked.");
  });

document.querySelector("#copySoOrders").addEventListener("click", () => {
  setOutput("Copy SO Orders clicked.");
});

document
  .querySelector("#copyTrackingUrls")
  .addEventListener("click", () => {
    setOutput("Copy Tracking URLs clicked.");
  });
