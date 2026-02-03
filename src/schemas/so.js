const normalizeValue = (value) => {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "string") {
    return value.trim();
  }
  return value;
};

const normalizeKey = (key) => {
  const trimmed = String(key ?? "").trim();
  if (!trimmed) {
    return "";
  }

  return trimmed
    .replace(/[_-]+/g, " ")
    .replace(/[^a-zA-Z0-9 ]+/g, " ")
    .split(" ")
    .filter(Boolean)
    .map((word, index) => {
      const lower = word.toLowerCase();
      if (index === 0) {
        return lower;
      }
      return `${lower.charAt(0).toUpperCase()}${lower.slice(1)}`;
    })
    .join("");
};

export const mapSoRow = (row = {}) => {
  const normalized = {};

  Object.entries(row).forEach(([key, value]) => {
    const normalizedKey = normalizeKey(key) || "field";
    normalized[normalizedKey] = normalizeValue(value);
  });

  return normalized;
};
