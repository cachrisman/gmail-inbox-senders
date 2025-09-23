
/***********************
 * Config & Utilities
 ***********************/
var email_regex = /([A-z0-9._%+-]+)@([A-z0-9.-]+\.[A-z]{2,})/g;
var page_size_default = 50; // safer batch size

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Gmail Inbox Senders");
}

/***********************
 * Counts
 ***********************/
function getThreadEstimate(search) {
  var resp = Gmail.Users.Threads.list("me", { q: search, maxResults: 1 });
  return resp.resultSizeEstimate || 0;
}

function getThreadCountExact(search) {
  var total = 0, token = null;
  var pageSize = page_size_default;
  do {
    var resp = Gmail.Users.Threads.list("me", {
      q: search,
      maxResults: pageSize,
      pageToken: token
    });
    total += (resp.threads ? resp.threads.length : 0);
    token = resp.nextPageToken || null;
  } while (token);
  return total;
}

/***********************
 * Senders (interactive paging)
 ***********************/
function getSendersPage(search, pageToken, pageSize) {
  var senders = {};
  var resp;
  try {
    resp = Gmail.Users.Threads.list("me", {
      q: search,
      maxResults: pageSize || page_size_default,
      pageToken: pageToken || null
    });
  } catch (e) {
    Logger.log("getSendersPage list error: " + e.message);
    return { senders: [], nextPageToken: null };
  }

  var threads = resp.threads || [];
  threads.forEach(function(t) {
    var thread;
    try {
      thread = Gmail.Users.Threads.get("me", t.id, {
        format: "metadata",
        metadataHeaders: ["From"]
      });
    } catch (e) {
      Logger.log("Skipping thread " + t.id + " (get error): " + e.message);
      return;
    }
    if (!thread || !thread.messages) return;

    thread.messages.forEach(function(m, i) {
      var headers = m.payload && m.payload.headers ? m.payload.headers : [];
      var from = headers.find(function(h){ return h.name === "From"; });
      if (!from) return;
      var match = from.value.match(email_regex);
      if (!match) return;
      var address = match[0];
      if (!senders[address]) senders[address] = { total:0, unread:0, threads:0 };
      senders[address].total++;
      if (m.labelIds && m.labelIds.indexOf("UNREAD") !== -1) {
        senders[address].unread++;
      }
      if (i === 0) senders[address].threads++;
    });
  });

  return {
    senders: Object.entries(senders),
    nextPageToken: resp.nextPageToken || null
  };
}

/***********************
 * Subjects
 ***********************/
function getSubjects(sender, search) {
  var q = search + " from:" + sender;
  var resp;
  try {
    resp = Gmail.Users.Threads.list("me", { q: q, maxResults: 50 });
  } catch (e) {
    Logger.log("getSubjects list error: " + e.message);
    return { sender: sender, subjects: [] };
  }

  var threads = resp.threads || [];
  var subjects = threads.map(function(t) {
    var thread;
    try {
      thread = Gmail.Users.Threads.get("me", t.id, {
        format: "metadata",
        metadataHeaders: ["Subject", "Date"]
      });
    } catch (e) {
      Logger.log("getSubjects get error for " + t.id + ": " + e.message);
      return null;
    }
    if (!thread || !thread.messages || !thread.messages.length) return null;

    var first = thread.messages[0];
    var last  = thread.messages[thread.messages.length - 1];

    var subjH = first.payload.headers.find(function(h){ return h.name === "Subject"; });
    var dateH = last.payload.headers.find(function(h){ return h.name === "Date"; });

    var subject = subjH ? subjH.value : "(no subject)";
    var dateStr = dateH ? new Date(dateH.value).toDateString() : "";

    return { subject: subject, date: dateStr };
  }).filter(function(x){ return x !== null; });

  return { sender: sender, subjects: subjects };
}

/***********************
 * Mark Read + Archive (interactive batch)
 ***********************/
function markReadAndArchiveBatch(sender, search, pageToken, batchSize) {
  var q = search + " from:" + sender;
  var processed = 0;
  var total = 0;
  var resp;
  try {
    resp = Gmail.Users.Threads.list("me", {
      q: q,
      maxResults: batchSize || page_size_default,
      pageToken: pageToken || null
    });
  } catch (e) {
    Logger.log("markReadAndArchiveBatch list error: " + e.message);
    return { sender: sender, processed: 0, total: 0, nextPageToken: null, error: e.message };
  }

  total = resp.resultSizeEstimate || 0;
  var threads = resp.threads || [];
  threads.forEach(function(t) {
    try {
      Gmail.Users.Threads.modify(
        { removeLabelIds: ["UNREAD", "INBOX"] },
        "me",
        t.id
      );
      processed++;
    } catch (e) {
      Logger.log("Modify error for " + t.id + ": " + e.message);
    }
  });

  return {
    sender: sender,
    processed: processed,
    total: total,
    nextPageToken: resp.nextPageToken || null
  };
}

/***********************
 * Background Jobs
 ***********************/
function ensureJobTrigger() {
  var exists = ScriptApp.getProjectTriggers().some(function(t){
    return t.getHandlerFunction() === "processBackgroundJobs";
  });
  if (!exists) {
    ScriptApp.newTrigger("processBackgroundJobs").timeBased().everyMinutes(1).create();
  }
}

function listRunningJobs() {
  const props = PropertiesService.getScriptProperties();
  const jobs = props.getProperties();
  return Object.values(jobs).map(j => JSON.parse(j)).filter(j => j.status === "running");
}

function listAllJobs() {
  const props = PropertiesService.getScriptProperties();
  const jobs = props.getProperties();
  return Object.values(jobs).map(j => JSON.parse(j));
}

function cancelJob(jobId) {
  const props = PropertiesService.getScriptProperties();
  const jobStr = props.getProperty(jobId);
  if (!jobStr) return { error: "Job not found" };
  const job = JSON.parse(jobStr);
  job.status = "cancelled";
  job.updatedAt = new Date().toISOString();
  props.setProperty(jobId, JSON.stringify(job));
  Logger.log(`cancelJob: job ${jobId} cancelled`);
  return { success: true, jobId: jobId };
}

function startMarkReadAndArchiveJob(search, sender) {
  if (listRunningJobs().length > 0) {
    Logger.log("startMarkReadAndArchiveJob blocked: another job is running");
    return { error: "Another job is already running" };
  }
  var jobId = Utilities.getUuid();
  var job = {
    id: jobId,
    type: "markReadAndArchive",
    search: search,
    sender: sender,
    status: "running",
    processed: 0,
    total: 0,
    pageToken: null,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  PropertiesService.getScriptProperties().setProperty(jobId, JSON.stringify(job));
  ensureJobTrigger();
  Logger.log(`startMarkReadAndArchiveJob: started job ${jobId}`);
  return jobId;
}

function startFetchSendersJob(search) {
  if (listRunningJobs().length > 0) {
    Logger.log("startFetchSendersJob blocked: another job is running");
    return { error: "Another job is already running" };
  }
  var jobId = Utilities.getUuid();
  var job = {
    id: jobId,
    type: "fetchSenders",
    search: search,
    status: "running",
    processed: 0,
    total: 0,
    pageToken: null,
    senders: {},
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  PropertiesService.getScriptProperties().setProperty(jobId, JSON.stringify(job));
  ensureJobTrigger();
  Logger.log(`startFetchSendersJob: started job ${jobId}`);
  return jobId;
}

function getJobStatus(jobId) {
  const jobStr = PropertiesService.getScriptProperties().getProperty(jobId);
  if (!jobStr) {
    Logger.log(`getJobStatus: job ${jobId} not found`);
    return { status: "unknown", id: jobId };
  }
  const job = JSON.parse(jobStr);
  Logger.log(`getJobStatus: job ${jobId} status=${job.status} processed=${job.processed}/${job.total}`);
  return job;
}

function processBackgroundJobs() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Logger.log(`processBackgroundJobs: found ${Object.keys(all).length} jobs`);

  Object.keys(all).forEach((jobId) => {
    let job = JSON.parse(all[jobId]);
    if (!job) {
      Logger.log(`processBackgroundJobs: job ${jobId} invalid JSON`);
      return;
    }
    Logger.log(`processBackgroundJobs: job ${jobId} type=${job.type} status=${job.status}`);
    if (job.status === "cancelled") {
      Logger.log(`processBackgroundJobs: job ${jobId} is cancelled, skipping`);
      return;
    }
    if (job.status !== "running") return;

    if (job.type === "markReadAndArchive") {
      Logger.log(`processBackgroundJobs: processing archive job ${jobId}`);
      processArchiveBatchJob(job, props, jobId);
    } else if (job.type === "fetchSenders") {
      Logger.log(`processBackgroundJobs: processing fetch job ${jobId}`);
      processFetchSendersBatchJob(job, props, jobId);
    }
  });
}

function processArchiveBatchJob(job, props, jobId) {
  var q = job.search + " from:" + job.sender;
  try {
    var resp = Gmail.Users.Threads.list("me", {
      q: q,
      maxResults: page_size_default,
      pageToken: job.pageToken || null
    });
    if (job.total === 0) job.total = resp.resultSizeEstimate || 0;

    var threads = resp.threads || [];
    threads.forEach(function(t) {
      try {
        Gmail.Users.Threads.modify(
          { removeLabelIds: ["UNREAD", "INBOX"] },
          "me",
          t.id
        );
        job.processed++;
      } catch (e) {
        Logger.log(`BG modify error ${t.id}: ${e.message}`);
      }
    });

    job.pageToken = resp.nextPageToken || null;
    job.updatedAt = new Date().toISOString();
    if (!resp.nextPageToken) {
      job.status = "done";
      job.finishedAt = new Date().toISOString();
    }
    props.setProperty(jobId, JSON.stringify(job));
  } catch (e) {
    job.status = "error";
    job.error = e.message;
    job.updatedAt = new Date().toISOString();
    props.setProperty(jobId, JSON.stringify(job));
    Logger.log(`processArchiveBatchJob error for ${jobId}: ${e.message}`);
  }
}

function processFetchSendersBatchJob(job, props, jobId) {
  try {
    var resp = Gmail.Users.Threads.list("me", {
      q: job.search,
      maxResults: page_size_default,
      pageToken: job.pageToken || null
    });
    if (job.total === 0) job.total = resp.resultSizeEstimate || 0;

    var threads = resp.threads || [];
    threads.forEach(function(t) {
      var thread;
      try {
        thread = Gmail.Users.Threads.get("me", t.id, {
          format: "metadata",
          metadataHeaders: ["From"]
        });
      } catch (e) {
        Logger.log(`BG get thread error ${t.id}: ${e.message}`);
        return;
      }
      if (!thread || !thread.messages) return;

      thread.messages.forEach(function(m, i) {
        var headers = m.payload && m.payload.headers ? m.payload.headers : [];
        var from = headers.find(function(h){ return h.name === "From"; });
        if (!from) return;
        var match = from.value.match(email_regex);
        if (!match) return;
        var address = match[0];
        if (!job.senders[address]) job.senders[address] = { total:0, unread:0, threads:0 };
        job.senders[address].total++;
        if (m.labelIds && m.labelIds.indexOf("UNREAD") !== -1) job.senders[address].unread++;
        if (i === 0) job.senders[address].threads++;
      });
    });

    job.processed += threads.length;
    job.pageToken = resp.nextPageToken || null;
    job.updatedAt = new Date().toISOString();
    if (!resp.nextPageToken) {
      job.status = "done";
      job.finishedAt = new Date().toISOString();
    }
    props.setProperty(jobId, JSON.stringify(job));
  } catch (e) {
    job.status = "error";
    job.error = e.message;
    job.updatedAt = new Date().toISOString();
    props.setProperty(jobId, JSON.stringify(job));
    Logger.log(`processFetchSendersBatchJob error for ${jobId}: ${e.message}`);
  }
}