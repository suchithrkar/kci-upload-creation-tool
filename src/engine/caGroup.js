const toValidDate = (value) => {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
};

const MS_PER_DAY = 1000 * 60 * 60 * 24;

export const getCaGroup = (createdOn) => {
  const createdDate = toValidDate(createdOn);
  if (!createdDate) {
    return "";
  }

  const now = new Date();
  const diffDays = (now - createdDate) / MS_PER_DAY;

  if (diffDays <= 3) {
    return "0–3 Days";
  }
  if (diffDays <= 5) {
    return "3–5 Days";
  }
  if (diffDays <= 10) {
    return "5–10 Days";
  }
  if (diffDays <= 15) {
    return "10–15 Days";
  }
  if (diffDays <= 30) {
    return "15–30 Days";
  }
  if (diffDays <= 60) {
    return "30–60 Days";
  }
  if (diffDays <= 90) {
    return "60–90 Days";
  }
  return "> 90 Days";
};
