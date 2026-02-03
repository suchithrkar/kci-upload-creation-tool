import { getCaGroup } from "./caGroup.js";

const normalizeString = (value) => {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
};

const getFirstValue = (row, keys) => {
  for (const key of keys) {
    if (row && row[key] !== undefined && row[key] !== null) {
      const normalized = normalizeString(row[key]);
      if (normalized) {
        return normalized;
      }
    }
  }
  return "";
};

const getCaseId = (row) =>
  getFirstValue(row, ["caseId", "caseNumber", "caseNo", "case"]);

const getOwner = (row) => getFirstValue(row, ["owner", "caseOwner"]);

const getStatus = (row) => getFirstValue(row, ["status", "caseStatus"]);

const getSbd = (row) => getFirstValue(row, ["sbd", "serviceByDate"]);

const getCreatedOn = (row) =>
  getFirstValue(row, ["createdOn", "createdDate", "createdAt"]);

const appendUnique = (list, value) => {
  if (!value) {
    return;
  }
  if (!list.includes(value)) {
    list.push(value);
  }
};

const toRows = (sheetData) => (Array.isArray(sheetData) ? sheetData : []);

const addOrderRefs = (caseEntry, row) => {
  appendUnique(caseEntry.relatedOrders.wo, getFirstValue(row, ["wo", "woId"]));
  appendUnique(caseEntry.relatedOrders.mo, getFirstValue(row, ["mo", "moId"]));
  appendUnique(caseEntry.relatedOrders.so, getFirstValue(row, ["so", "soId"]));
};

export const buildRepairCases = (normalizedData = {}) => {
  const casesById = new Map();

  Object.values(normalizedData).forEach((sheetRows) => {
    toRows(sheetRows).forEach((row) => {
      const caseId = normalizeString(getCaseId(row));
      if (!caseId) {
        return;
      }

      // Group rows by caseId to build a single repair case record per case.
      if (!casesById.has(caseId)) {
        casesById.set(caseId, {
          caseId,
          owner: "",
          status: "",
          caGroup: "",
          sbd: "",
          relatedOrders: {
            wo: [],
            mo: [],
            so: [],
          },
        });
      }

      const caseEntry = casesById.get(caseId);
      caseEntry.owner = caseEntry.owner || getOwner(row);
      caseEntry.status = caseEntry.status || getStatus(row);
      caseEntry.sbd = caseEntry.sbd || getSbd(row);

      if (!caseEntry.caGroup) {
        const createdOn = getCreatedOn(row);
        caseEntry.caGroup = createdOn ? getCaGroup(createdOn) : "";
      }

      // Collect related order references from any sheet rows tied to this case.
      addOrderRefs(caseEntry, row);
    });
  });

  return Array.from(casesById.values());
};
