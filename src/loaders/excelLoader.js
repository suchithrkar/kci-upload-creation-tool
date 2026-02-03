const normalizeCellValue = (value) => {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "string") {
    return value.trim();
  }
  return value;
};

const normalizeRow = (row) => {
  const normalized = {};

  Object.entries(row).forEach(([key, value]) => {
    normalized[key] = normalizeCellValue(value);
  });

  return normalized;
};

export const loadExcelFile = async (file) => {
  if (!file) {
    return {};
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  return workbook.SheetNames.reduce((acc, sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      defval: "",
    });

    acc[sheetName] = rows.map(normalizeRow);
    return acc;
  }, {});
};
