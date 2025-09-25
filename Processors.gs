/***********************
 * JOB PROCESSORS
 ***********************/
function computeExactCountForJob(jobId) {
  const info = getJobRowInfo(jobId);
  if (!info) return { ok: false, error: "Job not found" };

  const currentTotal = Number(info.job.Total || 0);
  if (currentTotal > 201)
    return { ok: true, total: currentTotal, skipped: true };

  const exact = getThreadCountExact(info.job.Search);
  const { sheet, headers, rowIndex } = info;
  const totalCol = headers.indexOf("Total") + 1;
  const updatedCol = headers.indexOf("UpdatedAt") + 1;
  if (totalCol > 0) sheet.getRange(rowIndex + 1, totalCol).setValue(exact);
  if (updatedCol > 0)
    sheet.getRange(rowIndex + 1, updatedCol).setValue(new Date().toISOString());
  return { ok: true, total: exact };
}

function getJobResult(jobId) {
  const job = listAllJobs().find((j) => j["Job ID"] === jobId);
  if (!job) return { error: `Job ${jobId} not found` };

  const sheet = getResultsSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const jobIdIdx = headers.indexOf("Job ID");

  if (job.Type === "fetchSenders") {
    const senders = {};
    values.slice(1).forEach((r) => {
      if (r[jobIdIdx] === jobId) {
        const addr = r[headers.indexOf("Address")];
        if (!addr) return;
        senders[addr] = {
          total: r[headers.indexOf("Total")] || 0,
          unread: r[headers.indexOf("Unread")] || 0,
          threads: r[headers.indexOf("Threads")] || 0,
          lastDate: r[headers.indexOf("LastDate")] || null,
        };
      }
    });
    return { ...job, senders: senders };
  }
  if (job.Type === "markReadAndArchive") return job;
  return { error: `Unsupported job type: ${job.Type}` };
}

function processArchiveBatchJob(job) {
  const q = job.Search ? `${job.Search} ${job.Query}` : job.Query;
  try {
    const resp = Gmail.Users.Threads.list("me", {
      q,
      maxResults: PAGE_SIZE_DEFAULT,
      pageToken: job.PageToken || null,
    });
    if (job.Total === 0) job.Total = resp.resultSizeEstimate || 0;

    (resp.threads || []).forEach((t) => {
      try {
        Gmail.Users.Threads.modify(
          { removeLabelIds: ["UNREAD", "INBOX"] },
          "me",
          t.id
        );
        job.Processed++;
      } catch (e) {
        job.Error = e.message;
      }
    });

    job.PageToken = resp.nextPageToken || "";
    job.UpdatedAt = new Date().toISOString();

    if (!resp.nextPageToken) {
      job.Status = "done";
      job.FinishedAt = new Date().toISOString();
      updateJob(job);
      writeJobResults(job["Job ID"], job, [
        { address: "", total: job.Processed, unread: "", threads: "" },
      ]);
    } else updateJob(job);
  } catch (e) {
    markJob(job, { Status: "error", Error: e.message });
  }
}

function processFetchSendersBatchJob(job) {
  try {
    const resp = Gmail.Users.Threads.list("me", {
      q: job.Search,
      maxResults: PAGE_SIZE_DEFAULT,
      pageToken: job.PageToken || null,
    });
    if (!job.Total) job.Total = resp.resultSizeEstimate || 0;

    const threads = resp.threads || [];
    const senders = {};

    threads.forEach((t) => {
      const thread = safeGetThread(t.id, ["From", "Date"]);
      if (!thread || !thread.messages) return;
      thread.messages.forEach((m, i) => updateSenderStats(senders, m, i === 0));
    });

    // ✅ Update Aggregated sheet incrementally
    const aggSheet = getAggregatedSheet();
    const aggValues = aggSheet.getDataRange().getValues();
    const aggHeaders = aggValues[0];
    const jobIdIdx = aggHeaders.indexOf("Job ID");
    const addrIdx = aggHeaders.indexOf("Address");

    const existing = {};
    aggValues.slice(1).forEach((r, i) => {
      if (r[jobIdIdx] === job["Job ID"]) {
        const addr = r[addrIdx];
        if (addr) existing[addr] = i + 2; // row number
      }
    });

    Object.entries(senders).forEach(([address, stats]) => {
      if (existing[address]) {
        // Update existing row
        const row = existing[address];
        aggSheet
          .getRange(row, aggHeaders.indexOf("Total") + 1)
          .setValue(stats.total);
        aggSheet
          .getRange(row, aggHeaders.indexOf("Unread") + 1)
          .setValue(stats.unread);
        aggSheet
          .getRange(row, aggHeaders.indexOf("Threads") + 1)
          .setValue(stats.threads);
        aggSheet
          .getRange(row, aggHeaders.indexOf("LastDate") + 1)
          .setValue(stats.lastDate || "");
        aggSheet
          .getRange(row, aggHeaders.indexOf("UpdatedAt") + 1)
          .setValue(new Date().toISOString());
      } else {
        // Append new row
        aggSheet.appendRow([
          job["Job ID"],
          address,
          stats.total,
          stats.unread,
          stats.threads,
          stats.lastDate || "",
          new Date().toISOString(),
        ]);
      }
    });

    // ✅ Update job progress
    job.Processed += threads.length;
    job.PageToken = resp.nextPageToken || "";
    job.UpdatedAt = new Date().toISOString();

    if (!resp.nextPageToken) {
      // Job finished
      job.Status = "done";
      job.FinishedAt = new Date().toISOString();
      updateJob(job);

      // Flush Aggregated → Results
      const allAgg = aggSheet.getDataRange().getValues();
      const headers = allAgg[0];
      const rows = allAgg
        .filter((r) => r[jobIdIdx] === job["Job ID"])
        .map((r) => ({
          address: r[addrIdx],
          total: r[headers.indexOf("Total")],
          unread: r[headers.indexOf("Unread")],
          threads: r[headers.indexOf("Threads")],
          lastDate: r[headers.indexOf("LastDate")],
        }));

      writeJobResults(job["Job ID"], job, rows);

      // (Optional) cleanup Aggregated rows for this job
      // cleanupAggregated(job["Job ID"]);
    } else {
      updateJob(job);
    }
  } catch (e) {
    markJob(job, { Status: "error", Error: e.message });
  }
}

function cleanupAggregated(jobId) {
  const sheet = getAggregatedSheet();
  const values = sheet.getDataRange().getValues();
  const jobIdIdx = values[0].indexOf("Job ID");
  const rowsToDelete = [];

  values.slice(1).forEach((r, i) => {
    if (r[jobIdIdx] === jobId) rowsToDelete.push(i + 2); // +2 = account for header
  });

  rowsToDelete.reverse().forEach((row) => sheet.deleteRow(row));
  Logger.log(
    `Cleaned up ${rowsToDelete.length} Aggregated rows for job ${jobId}`
  );
}
