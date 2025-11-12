/**
 * Gmail Analysis → Summary Report to Google Sheets
 * Notes:
 * - Uses GmailApp.
 * - Scans threads in pages to avoid timeouts.
 * - Configure QUERY and date bounds in CONFIG.
 */

var CONFIG = {
  // If running as standalone script, set SPREADSHEET_ID (recommended).
  // If left null, script will try getActiveSpreadsheet() or create one.
 SPREADSHEET_ID: null,
  SHEET_NAME: 'Gmail Analysis',
  CLEAR_SHEET_ON_RUN: true,

  QUERY: 'in:Promotions -in:spam -in:trash',

  // We now ignore AFTER_DAYS/BEFORE_DAYS and compute oldest + before:365d
  USE_OLDEST_AND_365D_WINDOW: true,

  PAGE_SIZE: 100,
  MAX_THREADS_PER_RUN: 800,

  PROP_NAMESPACE: 'GMAIL_ANALYSIS',
  TOKEN_KEY: 'CONT_TOKEN',
  OLDEST_DATE_PROP: 'OLDEST_MESSAGE_DATE_UTC' // stored as ISO string
};

/**
 * Entry point: Builds/updates the Gmail analysis report in the sheet.
 */
function buildReport() {
  var ss = getSpreadsheetHandle_();
  var sheet = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);

  if (CONFIG.CLEAR_SHEET_ON_RUN) {
    sheet.clear();
  }

  var bounds = buildDateQuery_();
  var query = CONFIG.QUERY + (bounds ? ' ' + bounds : '');
  var pageSize = CONFIG.PAGE_SIZE;
  var maxThreads = CONFIG.MAX_THREADS_PER_RUN;

  var aggregates = initAggregates_();

  var fetched = 0;
  var start = 0;

  while (true) {
    var threads = GmailApp.search(query, start, pageSize);
    if (!threads || threads.length === 0) break;

    for (var t = 0; t < threads.length; t++) {
      aggregateThread_(threads[t], aggregates);
      fetched++;
      if (fetched >= maxThreads) break;
    }

    if (fetched >= maxThreads) break;
    start += threads.length;

    Utilities.sleep(150);
  }

  writeSummaryToSheet_(sheet, query, aggregates, fetched);
}

/**
 * Ensure we have a Spreadsheet handle safely (bound, by ID, or created).
 */
function getSpreadsheetHandle_() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  // Standalone without ID: create once and remember
  var props = PropertiesService.getUserProperties();
  var cachedId = props.getProperty('GMAIL_ANALYSIS_SHEET_ID');
  if (cachedId) {
    try {
      return SpreadsheetApp.openById(cachedId);
    } catch (e) {
      // fall through to recreate
    }
  }
  var created = SpreadsheetApp.create('Gmail Analysis Report');
  props.setProperty('GMAIL_ANALYSIS_SHEET_ID', created.getId());
  return created;
}

/**
 * Returns a sheet by name or creates it if missing.
 */
function getOrCreateSheet_(ss, name) {
  if (!ss) throw new Error('Spreadsheet handle is null. Bind the script, set SPREADSHEET_ID, or allow auto-create.');
  var sh = ss.getSheetByName(name);
  return sh || ss.insertSheet(name);
}

/**
 * Build a Gmail query fragment for date bounds based on CONFIG.
 */
function buildDateQuery_() {
  var parts = [];
  var today = new Date();

  if (CONFIG.AFTER_DAYS != null) {
    var after = new Date(today.getTime() - CONFIG.AFTER_DAYS * 24 * 3600 * 1000);
    parts.push('after:' + formatYMD_(after));
  }
  if (CONFIG.BEFORE_DAYS != null) {
    var before = new Date(today.getTime() - CONFIG.BEFORE_DAYS * 24 * 3600 * 1000);
    parts.push('before:' + formatYMD_(before));
  }
  return parts.join(' ');
}

/**
 * Format date to YYYY/MM/DD for Gmail search operators.
 */
function formatYMD_(d) {
  var y = d.getFullYear();
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return y + '/' + m + '/' + day;
}

/**
 * Initialize aggregation containers.
 */
function initAggregates_() {
  return {
    totals: {
      threads: 0,
      messages: 0,
      unreadThreads: 0,
      unreadMessages: 0,
      withAttachments: 0,
      approxSizeBytes: 0
    },
    bySender: {},       // key: sender email
    byDomain: {},       // key: domain
    byLabel: {}         // key: label name
  };
}

/**
 * Aggregate a single thread into metrics maps.
 */
function aggregateThread_(thread, aggr) {
  aggr.totals.threads += 1;

  var msgs = thread.getMessages();
  var labels = thread.getLabels();
  var isThreadUnread = thread.isUnread();

  if (isThreadUnread) aggr.totals.unreadThreads += 1;

  var threadMsgCount = msgs.length;
  var threadUnreadMsgs = 0;
  var threadHasAttachment = false;
  var threadSize = 0;

  // Per-message size to avoid overcounting per sender/domain
  var perMessageSizes = [];

  for (var i = 0; i < msgs.length; i++) {
    var m = msgs[i];
    var from = parseEmail_(m.getFrom());
    var attachments = m.getAttachments({includeInlineImages: false, includeAttachments: true});
    var hasAtt = attachments && attachments.length > 0;

    if (m.isUnread()) threadUnreadMsgs++;

    var msgSize = 0;
    if (hasAtt) {
      threadHasAttachment = true;
      for (var a = 0; a < attachments.length; a++) {
        msgSize += attachments[a].getBytes().length;
      }
    }
    if (!hasAtt) {
      var body = m.getPlainBody() || m.getBody() || '';
      msgSize += body.length;
    }
    perMessageSizes.push({ email: from.email, domain: from.domain, size: msgSize });

    // By sender email (message-level)
    bump_(aggr.bySender, from.email, {
      messages: 1,
      unread: m.isUnread() ? 1 : 0,
      threads: 0,
      withAttachments: hasAtt ? 1 : 0
    });

    // By domain (message-level)
    bump_(aggr.byDomain, from.domain, {
      messages: 1,
      unread: m.isUnread() ? 1 : 0,
      threads: 0,
      withAttachments: hasAtt ? 1 : 0
    });
  }

  // Totals
  aggr.totals.messages += threadMsgCount;
  aggr.totals.unreadMessages += threadUnreadMsgs;
  if (threadHasAttachment) aggr.totals.withAttachments += 1;

  // Sum per-message sizes once into totals
  for (var p = 0; p < perMessageSizes.length; p++) {
    aggr.totals.approxSizeBytes += perMessageSizes[p].size;
  }

  // Allocate size per sender/domain based on message sizes
  for (var p2 = 0; p2 < perMessageSizes.length; p2++) {
    var rec = perMessageSizes[p2];
    bump_(aggr.bySender, rec.email, { approxSizeBytes: rec.size });
    bump_(aggr.byDomain, rec.domain, { approxSizeBytes: rec.size });
  }

  // Labels (thread-level)
  for (var l = 0; l < labels.length; l++) {
    var name = labels[l].getName();
    bump_(aggr.byLabel, name, {
      threads: 1,
      unreadThreads: isThreadUnread ? 1 : 0
    });
  }

  // Thread counts per unique sender/domain (once per thread)
  var participants = msgs.map(function(m){ return parseEmail_(m.getFrom()).email; });
  var uniqueSenders = uniq_(participants);
  for (var s = 0; s < uniqueSenders.length; s++) {
    ensureKey_(aggr.bySender, uniqueSenders[s]);
    aggr.bySender[uniqueSenders[s]].threads = (aggr.bySender[uniqueSenders[s]].threads || 0) + 1;
  }

  var uniqueDomains = uniq_(uniqueSenders.map(function(e){ return (e.split('@')[1] || '').toLowerCase(); }));
  for (var d = 0; d < uniqueDomains.length; d++) {
    ensureKey_(aggr.byDomain, uniqueDomains[d]);
    aggr.byDomain[uniqueDomains[d]].threads = (aggr.byDomain[uniqueDomains[d]].threads || 0) + 1;
  }
}

/**
 * Parse an email "Name <email@domain>" into parts.
 */
function parseEmail_(fromStr) {
  var emailMatch = /<?([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})>?/i.exec(fromStr || '');
  var email = emailMatch ? emailMatch[1].toLowerCase() : 'unknown@unknown';
  var domain = email.indexOf('@') > -1 ? email.split('@')[1].toLowerCase() : 'unknown';
  var name = (fromStr || '').replace(/<.*?>/g, '').trim();
  return { name: name, email: email, domain: domain };
}

/**
 * Ensure key exists in map.
 */
function ensureKey_(map, key) {
  if (!map[key]) map[key] = {};
}

/**
 * Bump metrics in a map.
 */
function bump_(map, key, deltas) {
  ensureKey_(map, key);
  var obj = map[key];
  for (var k in deltas) {
    if (!deltas.hasOwnProperty(k)) continue;
    var v = deltas[k];
    if (typeof v === 'number') {
      obj[k] = (obj[k] || 0) + v;
    } else {
      obj[k] = v;
    }
  }
}

/**
 * De-duplicate array preserving order.
 */
function uniq_(arr) {
  var seen = {};
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var val = arr[i];
    if (!seen[val]) {
      seen[val] = true;
      out.push(val);
    }
  }
  return out;
}

/**
 * Write results to the sheet: Overview, By Sender, By Domain, By Label.
 */
function writeSummaryToSheet_(sheet, query, aggr, fetchedThreads) {
  var row = 1;

  var overview = [
    ['Metric', 'Value'],
    ['Query', query],
    ['Threads scanned (this run)', fetchedThreads],
    ['Total threads', aggr.totals.threads],
    ['Total messages', aggr.totals.messages],
    ['Unread threads', aggr.totals.unreadThreads],
    ['Unread messages', aggr.totals.unreadMessages],
    ['Threads with attachments', aggr.totals.withAttachments],
    ['Approx size (MB)', Number((aggr.totals.approxSizeBytes / (1024*1024)).toFixed(2))]
  ];
  writeTable_(sheet, row, 1, overview);
  row += overview.length + 2;

  var senderRows = [['Sender Email', 'Threads', 'Messages', 'Unread Messages', 'With Attachments', 'Approx Size (MB)']];
  for (var email in aggr.bySender) {
    var s = aggr.bySender[email] || {};
    senderRows.push([
      email,
      s.threads || 0,
      s.messages || 0,
      s.unread || 0,
      s.withAttachments || 0,
      Number((((s.approxSizeBytes || 0) / (1024*1024))).toFixed(2))
    ]);
  }
  senderRows.sort(function(a,b){ return (b[2]||0) - (a[2]||0); });
  writeSection_(sheet, row, 'By Sender', senderRows);
  row += senderRows.length + 3;

  var domainRows = [['Domain', 'Threads', 'Messages', 'Unread Messages', 'With Attachments', 'Approx Size (MB)']];
  for (var domain in aggr.byDomain) {
    var d = aggr.byDomain[domain] || {};
    domainRows.push([
      domain,
      d.threads || 0,
      d.messages || 0,
      d.unread || 0,
      d.withAttachments || 0,
      Number((((d.approxSizeBytes || 0) / (1024*1024))).toFixed(2))
    ]);
  }
  domainRows.sort(function(a,b){ return (b[2]||0) - (a[2]||0); });
  writeSection_(sheet, row, 'By Domain', domainRows);
  row += domainRows.length + 3;

  var labelRows = [['Label', 'Threads', 'Unread Threads']];
  for (var label in aggr.byLabel) {
    var l = aggr.byLabel[label] || {};
    labelRows.push([label, l.threads || 0, l.unreadThreads || 0]);
  }
  labelRows.sort(function(a,b){ return (b[1]||0) - (a[1]||0); });
  writeSection_(sheet, row, 'By Label', labelRows);
  row += labelRows.length + 3;

  sheet.autoResizeColumns(1, Math.max(2, 6));
}

/**
 * Helper to write a titled section table.
 */
function writeSection_(sheet, startRow, title, rows) {
  sheet.getRange(startRow, 1, 1, 1).setValue(title).setFontWeight('bold').setFontSize(12);
  writeTable_(sheet, startRow + 1, 1, rows);
}

/**
 * Write a 2D array to a sheet.
 */
function writeTable_(sheet, row, col, data) {
  if (!data || !data.length) return;
  sheet.getRange(row, col, data.length, data[0].length).setValues(data);
}
