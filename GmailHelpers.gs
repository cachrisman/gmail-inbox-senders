/***********************
 * GMAIL HELPERS
 ***********************/
function getThreadCountExact(search) {
  let total = 0,
    token = null;
  do {
    const resp = Gmail.Users.Threads.list("me", {
      q: search,
      maxResults: PAGE_SIZE_DEFAULT,
      pageToken: token,
    });
    total += resp.threads ? resp.threads.length : 0;
    token = resp.nextPageToken || null;
  } while (token);
  return total;
}

function getThreadEstimate(search) {
  const resp = Gmail.Users.Threads.list("me", { q: search, maxResults: 1 });
  return resp.resultSizeEstimate || 0;
}

function getSendersPage(search, pageToken, pageSize) {
  const senders = {};
  const resp = Gmail.Users.Threads.list("me", {
    q: search,
    maxResults: pageSize || PAGE_SIZE_DEFAULT,
    pageToken: pageToken || null,
  });
  (resp.threads || []).forEach((t) => {
    const thread = safeGetThread(t.id, ["From", "Date"]);
    if (!thread || !thread.messages) return;
    thread.messages.forEach((m, i) => updateSenderStats(senders, m, i === 0));
  });
  return {
    senders: Object.entries(senders),
    nextPageToken: resp.nextPageToken || null,
  };
}

function getSubjects(sender, search) {
  const resp = Gmail.Users.Threads.list("me", {
    q: `${search} from:${sender}`,
    maxResults: 50,
  });
  return {
    sender,
    subjects: (resp.threads || [])
      .map((t) => {
        const thread = safeGetThread(t.id, ["Subject", "Date"]);
        if (!thread || !thread.messages) return null;
        const first = thread.messages[0];
        const last = thread.messages[thread.messages.length - 1];
        const subjH = first.payload.headers.find((h) => h.name === "Subject");
        const dateH = last.payload.headers.find((h) => h.name === "Date");
        return {
          subject: subjH ? subjH.value : "(no subject)",
          date: dateH ? new Date(dateH.value).toDateString() : "",
        };
      })
      .filter((x) => x),
  };
}

/***********************
 * SHARED HELPERS
 ***********************/
function safeGetThread(id, headers) {
  try {
    return Gmail.Users.Threads.get("me", id, {
      format: "metadata",
      metadataHeaders: headers,
    });
  } catch (e) {
    Logger.log(`Error fetching thread ${id}: ${e.message}`);
    return null;
  }
}

function updateSenderStats(map, message, isFirst) {
  const headers =
    message.payload && message.payload.headers ? message.payload.headers : [];
  const from = headers.find((h) => h.name === "From");
  if (!from) return;
  const match = from.value.match(EMAIL_REGEX);
  if (!match) return;
  const addr = match[0];
  if (!map[addr])
    map[addr] = { total: 0, unread: 0, threads: 0, lastDate: null };

  map[addr].total++;
  if (message.labelIds && message.labelIds.includes("UNREAD"))
    map[addr].unread++;
  if (isFirst) map[addr].threads++;

  if (message.internalDate) {
    const msgDate = new Date(parseInt(message.internalDate));
    if (!map[addr].lastDate || msgDate > new Date(map[addr].lastDate)) {
      map[addr].lastDate = msgDate.toISOString();
    }
  }
}
