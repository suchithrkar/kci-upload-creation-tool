const normalizeValue = (value) => {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value);
};

const escapeCsvValue = (value) => {
  const normalized = normalizeValue(value);
  const needsQuotes = /[",\n\r]/.test(normalized);
  const escaped = normalized.replace(/"/g, '""');
  return needsQuotes ? `"${escaped}"` : escaped;
};

const buildCsvFromObjects = (rows) => {
  const headers = [];
  const headerSet = new Set();

  rows.forEach((row) => {
    if (!row || typeof row !== "object") {
      return;
    }
    Object.keys(row).forEach((key) => {
      if (!headerSet.has(key)) {
        headerSet.add(key);
        headers.push(key);
      }
    });
  });

  if (headers.length === 0) {
    return "";
  }

  const lines = [headers.map(escapeCsvValue).join(",")];

  rows.forEach((row) => {
    const line = headers.map((header) =>
      escapeCsvValue(row ? row[header] : "")
    );
    lines.push(line.join(","));
  });

  return lines.join("\n");
};

const buildCsvFromValues = (rows) =>
  rows.map((value) => escapeCsvValue(value)).join("\n");

const buildCsv = (rows) => {
  if (!Array.isArray(rows) || rows.length === 0) {
    return "";
  }

  const hasObjectRow = rows.some(
    (row) => row && typeof row === "object" && !Array.isArray(row)
  );

  return hasObjectRow ? buildCsvFromObjects(rows) : buildCsvFromValues(rows);
};

export const exportCsv = (rows, filename) => {
  const csvContent = buildCsv(rows);

  if (!csvContent) {
    return false;
  }

  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  return true;
};
