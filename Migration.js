/***********************
 * MIGRATION HELPERS
 ***********************/
function migrateJobsFromProperties() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  if (!allProps || Object.keys(allProps).length === 0) return;

  const sheet = getJobsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = [];

  Object.keys(allProps).forEach((jobId) => {
    try {
      const job = JSON.parse(allProps[jobId]);
      const rowObj = {
        "Job ID": jobId || job.jobId,
        Type: job.Type || job.type,
        Search: job.Search || job.search,
        Target: job.Target || job.target,
        Query: job.Query || job.query,
        Status: job.Status || job.status,
        Processed: job.Processed || job.processed || 0,
        Total: job.Total || job.total || 0,
        PageToken: job.PageToken || job.pageToken || "",
        StartedAt: job.StartedAt || job.startedAt || "",
        UpdatedAt: job.UpdatedAt || job.updatedAt || new Date().toISOString(),
        Error: job.Error || job.error || "",
        FinishedAt: job.FinishedAt || job.finishedAt || "",
      };
      rows.push(headers.map((h) => rowObj[h] || ""));
    } catch (e) {
      Logger.log(`Skipping property ${jobId}: ${e.message}`);
    }
  });

  if (rows.length > 0)
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length)
      .setValues(rows);
}

function migrateResultsFromProperties() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  if (!allProps || Object.keys(allProps).length === 0) return;

  const sheet = getResultsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = [];

  Object.keys(allProps).forEach((key) => {
    try {
      const val = JSON.parse(allProps[key]);
      if (
        val &&
        (val.Type === "fetchSenders" || val.type === "fetchSenders") &&
        val.senders
      ) {
        Object.entries(val.senders).forEach(([address, stats]) => {
          const rowObj = {
            "Job ID": key,
            Type: val.Type || val.type,
            Search: val.Search || val.search || "",
            Target: val.Target || val.target || "",
            Address: address,
            Total: stats.total || 0,
            Unread: stats.unread || 0,
            Threads: stats.threads || 0,
            LastDate: stats.lastDate || "",
            FinishedAt:
              val.FinishedAt || val.finishedAt || new Date().toISOString(),
          };
          rows.push(headers.map((h) => rowObj[h] || ""));
        });
      }
      if (
        val &&
        (val.Type === "markReadAndArchive" || val.type === "markReadAndArchive")
      ) {
        const rowObj = {
          "Job ID": key,
          Type: val.Type || val.type,
          Search: val.Search || val.search || "",
          Target: val.Target || val.target || "",
          Address: "",
          Total: val.Processed || 0,
          Unread: "",
          Threads: "",
          LastDate: "",
          FinishedAt:
            val.FinishedAt || val.finishedAt || new Date().toISOString(),
        };
        rows.push(headers.map((h) => rowObj[h] || ""));
      }
    } catch (e) {
      Logger.log(`Skipping property ${key}: ${e.message}`);
    }
  });

  if (rows.length > 0)
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length)
      .setValues(rows);
}
