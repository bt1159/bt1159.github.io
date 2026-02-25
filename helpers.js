const helpers = {
    
/**
 * Convert an Excel serial date number into a JavaScript Date.
 *
 * Excel’s default (1900) date system stores dates as the number of days since
 * 1899‑12‑30 (because of the historical “1900 is a leap year” quirk).
 *
 * @param {number} excelDate - Days since 1899‑12‑30 in the Excel 1900 system.
 * @returns {Date} A JavaScript Date representing the same calendar day.
 */
excelDateToJS: function (excelDate) {
  let excelDateFixed = excelDate ?? 0;
  const date = new Date(Math.round((excelDateFixed - 25569) * 86400 * 1000));
  return date;
}
}
