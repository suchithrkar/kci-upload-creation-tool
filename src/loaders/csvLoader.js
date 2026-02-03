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

export const loadCsvFile = async (file) => {
  if (!file) {
    return [];
  }

  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        const data = Array.isArray(results.data) ? results.data : [];
        resolve(data.map(normalizeRow));
      },
      error: (error) => {
        reject(error);
      },
    });
  });
};
