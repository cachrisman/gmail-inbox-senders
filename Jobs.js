/***********************
 * JOB MANAGEMENT
 ***********************/
function listAllJobs() {
  return sheetToObjects(getJobsSheet());
}

function listRunningJobs() {
  return listAllJobs().filter(j => j.Status === "running");
}

/**
 * Bulk-write job results (sender rows) to the Results sheet.
 * Expects rows: [{address,total,unread,threads,lastDate}, ...]
 * Returns number of rows written.
 */
function writeJobResults(jobId, job, rows) {
  if (!rows || !rows.length) return 0;

  const sheet = getResultsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Required headers in Results sheet
  const REQUIRED = ["Job ID","Type","Search","Target","Address","Total","Unread","Threads","LastDate","FinishedAt"];
  REQUIRED.forEach(h => {
    if (headers.indexOf(h) === -1) {
      throw new Error("Results sheet missing required header: " + h);
    }
  });

  const startRow = sheet.getLastRow() + 1;
  const values = rows.map(r => ([
    jobId,
    job.Type || "",
    job.Search || "",
    job.Target || "",
    r.address || "",
    Number(r.total || 0),
    Number(r.unread || 0),
    Number(r.threads || 0),
    r.lastDate || "",
    job.FinishedAt || new Date().toISOString()
  ]));

  sheet.getRange(startRow, 1, values.length, headers.length).setValues(values);
  return values.length;
}

/***********************
 * JOB CREATION / CONTROL
 ***********************/
function startMarkReadAndArchiveJob(search, target) {
  const jobId = Utilities.getUuid();
  // Handle array of targets (domains)
  const targets = Array.isArray(target) ? target : [target];
  // Build OR query for all domains
  const queryTarget = targets
    .map(t => t && t.indexOf("@") !== -1 ? `from:${t}` : `from:*@${t}`)
    .join(" OR ");

  const job = {
    "Job ID": jobId,
    "Type": "markReadAndArchive",
    "Search": search,
    "Target": Array.isArray(target) ? target.join(", ") : target,
    "Query": queryTarget,
    "Status": "queued",
    "Processed": 0,
    "Total": 0,
    "PageToken": "",
    "StartedAt": "",
    "UpdatedAt": new Date().toISOString(),
    "Error": ""
  };
  saveJob(job);
  ensureJobTrigger();
  return jobId;
}

function startFetchSendersJob(search) {
  const jobId = Utilities.getUuid();
  const job = {
    "Job ID": jobId,
    "Type": "fetchSenders",
    "Search": search,
    "Target": search,
    "Status": "queued",
    "Processed": 0,
    "Total": 0,
    "PageToken": "",
    "StartedAt": "",
    "UpdatedAt": new Date().toISOString(),
    "Error": ""
  };
  saveJob(job);
  ensureJobTrigger();
  return jobId;
}

function cancelJob(jobId) {
  markJob({ "Job ID": jobId }, { Status: "cancelled" });
  return { success: true, jobId: jobId };
}

function getJobStatus(jobId) {
  const info = getJobRowInfo(jobId);
  return info ? info.job : { status: "unknown", id: jobId };
}

/**
 * Update a job with fields + UpdatedAt timestamp.
 */
function markJob(job, fields) {
  Object.assign(job, fields, { UpdatedAt: new Date().toISOString() });
  updateJob(job);
}

/***********************
 * JOB TRIGGER / SCHEDULER
 ***********************/
function ensureJobTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(
    t => t.getHandlerFunction() === "processBackgroundJobs"
  );
  if (!exists) {
    ScriptApp.newTrigger("processBackgroundJobs").timeBased().everyMinutes(1).create();
  }
}

/**
 * One-at-a-time job runner: picks queued → running, then processes batch.
 */
function processBackgroundJobs() {
  const jobs = listAllJobs();
  if (!jobs || jobs.length === 0) {
    Logger.log("No jobs in Jobs sheet.");
    return;
  }

  // Find running, else promote one queued
  let running = jobs.find(j => j.Status === "running");
  if (!running) {
    const queued = jobs.find(j => j.Status === "queued");
    if (queued) {
      queued.Status = "running";
      queued.StartedAt = new Date().toISOString();
      queued.UpdatedAt = new Date().toISOString();
      updateJob(queued);
      running = queued;
      Logger.log(`Started queued job ${running["Job ID"]} (${running.Type})`);
    }
  }
  if (!running) {
    Logger.log("No queued or running jobs to process.");
    return;
  }

  Logger.log(`Processing job ${running["Job ID"]} type=${running.Type}`);
  try {
    if (running.Type === "markReadAndArchive") {
      processArchiveBatchJob(running);            // in Processors.gs
    } else if (running.Type === "fetchSenders") {
      processFetchSendersBatchJob(running);       // in Processors.gs
    } else {
      markJob(running, { Status: "error", Error: `Unknown job type: ${running.Type}` });
    }
  } catch (e) {
    markJob(running, { Status: "error", Error: e.message });
  }
}