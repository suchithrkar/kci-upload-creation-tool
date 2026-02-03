import { loadExcelFile } from "../src/loaders/excelLoader.js";
import { loadCsvFile } from "../src/loaders/csvLoader.js";

const output = document.querySelector("#output");
const state = window.appState;

const setOutput = (message) => {
  output.textContent = message;
};

const handleExcelLoad = async (file) => {
  state.kciExcelFile = file || null;
  state.kciExcelData = await loadExcelFile(file);
  setOutput("KCI Excel loaded.");
};

const handleCsvLoad = async (file, targetKey, label) => {
  state[targetKey].file = file || null;
  state[targetKey].data = await loadCsvFile(file);
  setOutput(`${label} loaded.`);
};

document.querySelector("#kciExcel").addEventListener("change", (event) => {
  handleExcelLoad(event.target.files?.[0]).catch((error) => {
    setOutput(`Failed to load KCI Excel: ${error.message}`);
  });
});

document.querySelector("#csoCsv").addEventListener("change", (event) => {
  handleCsvLoad(event.target.files?.[0], "csoCsv", "CSO CSV").catch(
    (error) => {
      setOutput(`Failed to load CSO CSV: ${error.message}`);
    }
  );
});

document
  .querySelector("#trackingCsv")
  .addEventListener("change", (event) => {
    handleCsvLoad(event.target.files?.[0], "trackingCsv", "Tracking CSV").catch(
      (error) => {
        setOutput(`Failed to load Tracking CSV: ${error.message}`);
      }
    );
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
