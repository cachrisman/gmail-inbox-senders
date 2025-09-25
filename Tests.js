/***********************
 * TEST / SANITY CHECKS
 ***********************/
function testEnvironment() {
  Logger.log("=== Running Environment Sanity Check ===");

  // Spreadsheet access
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log(`Spreadsheet opened: ${ss.getName()}`);
  } catch (e) {
    Logger.log(`❌ Spreadsheet error: ${e.message}`);
    return;
  }

  // Jobs headers
  try {
    const jobs = getJobsSheet();
    const jobHeaders = jobs
      .getRange(1, 1, 1, jobs.getLastColumn())
      .getValues()[0];
    Logger.log("Jobs headers: " + jobHeaders.join(", "));
    const requiredJobs = [
      "Job ID",
      "Type",
      "Search",
      "Target",
      "Query",
      "Status",
      "Processed",
      "Total",
      "PageToken",
      "StartedAt",
      "UpdatedAt",
      "Error",
      "FinishedAt",
    ];
    requiredJobs.forEach((h) => {
      if (jobHeaders.indexOf(h) === -1)
        Logger.log(`❌ Missing Jobs header: ${h}`);
    });
  } catch (e) {
    Logger.log(`❌ Jobs sheet error: ${e.message}`);
  }

  // Results headers
  try {
    const results = getResultsSheet();
    const resultHeaders = results
      .getRange(1, 1, 1, results.getLastColumn())
      .getValues()[0];
    Logger.log("Results headers: " + resultHeaders.join(", "));
    const requiredResults = [
      "Job ID",
      "Type",
      "Search",
      "Target",
      "Address",
      "Total",
      "Unread",
      "Threads",
      "LastDate",
      "FinishedAt",
    ];
    requiredResults.forEach((h) => {
      if (resultHeaders.indexOf(h) === -1)
        Logger.log(`❌ Missing Results header: ${h}`);
    });
  } catch (e) {
    Logger.log(`❌ Results sheet error: ${e.message}`);
  }

  // Aggregated headers
  try {
    const agg = getAggregatedSheet();
    const aggHeaders = agg
      .getRange(1, 1, 1, agg.getLastColumn())
      .getValues()[0];
    Logger.log("Aggregated headers: " + aggHeaders.join(", "));
    const requiredAgg = [
      "Job ID",
      "Address",
      "Total",
      "Unread",
      "Threads",
      "LastDate",
      "UpdatedAt",
    ];
    requiredAgg.forEach((h) => {
      if (aggHeaders.indexOf(h) === -1)
        Logger.log(`❌ Missing Aggregated header: ${h}`);
    });
  } catch (e) {
    Logger.log(`❌ Aggregated sheet error: ${e.message}`);
  }

  // Gmail API connectivity
  try {
    const estimate = getThreadEstimate("in:inbox");
    Logger.log(`Gmail API estimate for in:inbox = ${estimate}`);
  } catch (e) {
    Logger.log(`❌ Gmail API error: ${e.message}`);
  }

  // Backend function presence
  testBackendFunctions();

  Logger.log("=== Sanity Check Complete ===");
}

/**
 * Verify that all expected backend functions are present.
 */
function testBackendFunctions() {
  const names = [
    // Sheets
    "getJobsSheet",
    "getResultsSheet",
    "getAggregatedSheet",
    "sheetToObjects",
    "getJobRowInfo",
    "saveJob",
    "updateJob",

    // Jobs & scheduler
    "listAllJobs",
    "listRunningJobs",
    "startMarkReadAndArchiveJob",
    "startFetchSendersJob",
    "cancelJob",
    "getJobStatus",
    "markJob",
    "ensureJobTrigger",
    "processBackgroundJobs",
    "writeJobResults",

    // Processors
    "processArchiveBatchJob",
    "processFetchSendersBatchJob",
    "getJobResult",
    "computeExactCountForJob",

    // Gmail helpers
    "getThreadCountExact",
    "getThreadEstimate",
    "getSendersPage",
    "getSubjects",
    "safeGetThread",
    "updateSenderStats",
  ];

  const missing = [];
  names.forEach((n) => {
    try {
      if (typeof this[n] !== "function") missing.push(n);
    } catch (e) {
      missing.push(n);
    }
  });

  if (missing.length) {
    Logger.log("❌ Missing functions: " + missing.join(", "));
  } else {
    Logger.log("✅ All expected backend functions are present.");
  }
}
