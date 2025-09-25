/***********************
 * SHEET HELPERS
 ***********************/
function getJobsSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Jobs");
}
function getResultsSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Results");
}
function getAggregatedSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Aggregated");
}

/**
 * Convert a sheet into an array of objects using header row.
 */
function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map((r) => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = r[i]));
    return obj;
  });
}

/**
 * Find job row by ID with sheet + headers.
 */
function getJobRowInfo(jobId) {
  const sheet = getJobsSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const jobIdIdx = headers.indexOf("Job ID");
  if (jobIdIdx === -1) throw new Error("Jobs sheet missing 'Job ID' column");
  const rowIndex = values.findIndex((r) => r[jobIdIdx] === jobId);
  if (rowIndex <= 0) return null;
  const job = {};
  headers.forEach((h, i) => (job[h] = values[rowIndex][i]));
  return { sheet, headers, rowIndex, job };
}

/**
 * Save a new job row.
 */
function saveJob(job) {
  const sheet = getJobsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map((h) => (h in job ? job[h] : ""));
  sheet.appendRow(row);
}

/**
 * Update an existing job row.
 */
function updateJob(job) {
  const sheet = getJobsSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const jobIdIdx = headers.indexOf("Job ID");
  const rowIdx = values.findIndex((r) => r[jobIdIdx] === job["Job ID"]);
  if (rowIdx <= 0) {
    Logger.log(`updateJob: Job ID ${job["Job ID"]} not found`);
    return;
  }
  headers.forEach((h, i) => {
    if (h in job) {
      sheet.getRange(rowIdx + 1, i + 1).setValue(job[h]);
    }
  });
}
