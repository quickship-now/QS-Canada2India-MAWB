// Quickship MAWB System
// Sections: Sahib Khan -> Meenakshi

var BOOKING_SHEET = "FreightBookings";
var USERS_SHEET = "Users";
var USER_CACHE_KEY = "quickship_mawb_users_v1";
var USER_CACHE_SECONDS = 600;

var BOOKING_HEADERS = [
  "Record ID", "Created At", "Last Updated",
  "Booking Date", "Consignee", "IATA Agent", "MAWB", "Weight", "Booked By", "Origin", "Destination",
  "Sahib Khan Status", "Sahib Khan Edit URL",
  "Pre Alert DateTime", "Pre Alert Sent To URL", "Manifest DateTime", "Pre Alert Ack DateTime", "Pick Up DateTime",
  "Meenakshi Status", "Meenakshi Edit URL"
];

var BCI = (function() {
  var m = {};
  for (var i = 0; i < BOOKING_HEADERS.length; i++) m[BOOKING_HEADERS[i]] = i;
  return m;
})();

var USER_HEADERS = ["Email", "Password", "Role", "Name"];
var DEFAULT_USERS = [
  ["sk@quickshipnow.com", "sahib123", "sahib", "Sahib Khan"],
  ["meenakshi@quickshipnow.com", "meenakshi123", "meenakshi", "Meenakshi"]
];

var COL_GROUPS = [
  { s:1, e:3, bg:"#1e3a5f" },   // Meta
  { s:4, e:13, bg:"#be185d" },  // Sahib Khan
  { s:14, e:20, bg:"#0f766e" }  // Meenakshi
];

var STATUS_DONE = "Submitted \u2713";
var STATUS_PENDING = "Pending";

var HEADER_ALIASES = {
  "Sanjay Status": "Sahib Khan Status",
  "Sanjay Edit URL": "Sahib Khan Edit URL"
};

var SECTION_FIELD_KEYS = {
  sahib: [
    "Booking Date", "Consignee", "IATA Agent", "MAWB", "Weight", "Booked By", "Origin", "Destination"
  ],
  meenakshi: [
    "Pre Alert DateTime", "Pre Alert Sent To URL", "Manifest DateTime", "Pre Alert Ack DateTime", "Pick Up DateTime"
  ]
};

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var page = params.page || "login";
  var section = (params.section || "sahib").toLowerCase();
  var recordId = params.recordId || "";
  var validSections = ["sahib", "meenakshi"];
  if (validSections.indexOf(section) === -1) section = "sahib";

  var recordData = {};
  if (page === "form" && recordId) {
    try {
      var found = getRecord(recordId);
      if (found && found.found) recordData = found.data;
    } catch (err) {}
  }

  var tmpl = HtmlService.createTemplateFromFile("index");
  tmpl.APP_INIT_JSON = JSON.stringify({
    page: page,
    section: section,
    recordId: recordId,
    data: recordData
  });

  return tmpl.evaluate()
    .setTitle("Quickship MAWB System")
    .addMetaTag("viewport", "width=device-width,initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Quickship")
      .addItem("Rearrange Headers & Sync Data", "rearrangeHeadersAndSyncData")
      .addItem("Refresh Section Statuses", "refreshAllStatuses")
      .addToUi();
  } catch (err) {}
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (!sh || sh.getName() !== BOOKING_SHEET) return;
    if (e.range.getLastRow() < 2) return;

    var startRow = Math.max(2, e.range.getRow());
    var endRow = Math.min(sh.getLastRow(), e.range.getLastRow());
    refreshStatusRows_(sh, startRow, endRow - startRow + 1);
  } catch (err) {}
}

function refreshAllStatuses() {
  var sh = getOrCreateBookingSheet();
  var lr = sh.getLastRow();
  if (lr < 2) return { success: true, updatedRows: 0 };
  return refreshStatusRows_(sh, 2, lr - 1);
}

function rearrangeHeadersAndSyncData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BOOKING_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BOOKING_SHEET);
    sh.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS]);
    sh.setFrozenRows(1);
    styleBookingHeaders(sh);
    return { success: true, rowsSynced: 0, message: "Created missing FreightBookings sheet." };
  }

  var currentLastCol = Math.max(sh.getLastColumn(), BOOKING_HEADERS.length);
  var currentHeaders = sh.getRange(1, 1, 1, currentLastCol).getValues()[0];
  return migrateBookingHeaders_(sh, currentHeaders, { rebuildLinks: true, refreshStatuses: true });
}

function login(email, password) {
  try {
    var rows = getCachedUserRows_();
    var em = normalizeLoginEmail_(email);
    var pass = String(password || "").trim();

    for (var i = 1; i < rows.length; i++) {
      var rowEmail = normalizeLoginEmail_(rows[i][0]);
      if (rowEmail === em && String(rows[i][1]).trim() === pass) {
        var role = normalizeRole_(rows[i][2]);
        if (!SECTION_META.hasOwnProperty(role)) {
          return { success: false, error: "This user no longer has an active tracker role." };
        }
        return {
          success: true,
          user: {
            email: String(rows[i][0]).trim(),
            role: role,
            name: String(rows[i][3]).trim()
          }
        };
      }
    }
    return { success: false, error: "Invalid email or password." };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getDashboard() {
  try {
    var sh = getOrCreateBookingSheet();
    var lr = sh.getLastRow();
    if (lr < 2) return { success: true, rows: [] };
    var rows = sh.getRange(2, 1, lr - 1, BOOKING_HEADERS.length).getDisplayValues();

    var out = [];
    for (var i = 0; i < rows.length; i++) {
      if (!String(rows[i][0] || "").trim()) continue;
      var obj = {};
      BOOKING_HEADERS.forEach(function(h, idx) {
        obj[h] = (rows[i][idx] !== undefined && rows[i][idx] !== null) ? String(rows[i][idx]) : "";
      });
      out.push(obj);
    }
    return { success: true, rows: out };
  } catch (err) {
    return { success: false, error: err.message, rows: [] };
  }
}

function getRecord(recordId) {
  try {
    if (!recordId) return { found: false, data: {} };
    var sh = getOrCreateBookingSheet();
    var row = findRow(sh, recordId);
    if (row < 0) return { found: false, data: {} };

    var vals = sh.getRange(row, 1, 1, BOOKING_HEADERS.length).getDisplayValues()[0];
    var obj = {};
    BOOKING_HEADERS.forEach(function(h, i) {
      obj[h] = (vals[i] !== null && vals[i] !== undefined) ? String(vals[i]) : "";
    });
    return { found: true, data: obj };
  } catch (err) {
    return { found: false, data: {}, error: err.message };
  }
}

var SECTION_META = {
  sahib:   { statusCol:"Sahib Khan Status", urlCol:"Sahib Khan Edit URL", label:"Edit Sahib Khan Section" },
  meenakshi: { statusCol:"Meenakshi Status", urlCol:"Meenakshi Edit URL", label:"Edit Meenakshi Section" }
};

function saveSection(payload) {
  try {
    var sh = getOrCreateBookingSheet();
    var section = normalizeRole_(payload.section || "sahib");
    if (!SECTION_META.hasOwnProperty(section)) {
      return { success: false, error: "Invalid or inactive section." };
    }
    var recordId = String(payload.recordId || "");
    var fields = payload.fields || {};
    var nCols = BOOKING_HEADERS.length;
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    var row = recordId ? findRow(sh, recordId) : -1;
    var isNew = row < 0;

    if (isNew && section !== "sahib") {
      return { success: false, error: "Sahib Khan must create the booking before Meenakshi can update it." };
    }

    if (isNew) {
      recordId = "FR-" + new Date().getTime();
      var blank = new Array(nCols).fill("");
      blank[bci("Record ID")] = recordId;
      blank[bci("Created At")] = now;
      blank[bci("Last Updated")] = now;
      sh.appendRow(blank);
      row = sh.getLastRow();
    }

    var rowRange = sh.getRange(row, 1, 1, nCols);
    var rowVals = rowRange.getValues()[0];

    rowVals[bci("Last Updated")] = now;

    (SECTION_FIELD_KEYS[section] || []).forEach(function(key) {
      if (!fields.hasOwnProperty(key)) return;
      var c = bci(key);
      if (c >= 0) rowVals[c] = fields[key];
    });

    var urls = buildSectionUrls_(recordId);

    var statuses = updateSectionStatuses_(rowVals);

    syncRowEditLinks_(rowVals, recordId, urls);

    rowRange.setValues([rowVals]);

    if (isNew) {
      Object.keys(SECTION_META).forEach(function(k) {
        sh.getRange(row, bci(SECTION_META[k].urlCol) + 1).setFontColor("#1155cc");
      });
    }

    return { success: true, recordId: recordId, urls: urls, statuses: statuses };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateSectionStatuses_(rowVals) {
  var statuses = {};
  Object.keys(SECTION_META).forEach(function(section) {
    var meta = SECTION_META[section];
    var complete = isSectionComplete_(rowVals, section);
    var status = complete ? STATUS_DONE : STATUS_PENDING;
    rowVals[bci(meta.statusCol)] = status;
    statuses[section] = status;
  });
  return statuses;
}

function buildSectionUrls_(recordId) {
  var base = "";
  try {
    base = ScriptApp.getService().getUrl();
  } catch (err) {}

  var urls = {};
  Object.keys(SECTION_META).forEach(function(k) {
    urls[k] = base ? (base + "?page=form&recordId=" + encodeURIComponent(recordId) + "&section=" + k) : "";
  });
  return urls;
}

function syncRowEditLinks_(rowVals, recordId, urls) {
  urls = urls || buildSectionUrls_(recordId);
  Object.keys(SECTION_META).forEach(function(k) {
    var m = SECTION_META[k];
    var col = bci(m.urlCol);
    if (col < 0) return;
    rowVals[col] = urls[k] ? '=HYPERLINK("' + urls[k] + '","' + m.label + '")' : rowVals[col];
  });
}

function refreshStatusRows_(sh, startRow, numRows) {
  if (!numRows || numRows < 1) return { success: true, updatedRows: 0 };
  var nCols = BOOKING_HEADERS.length;
  var rowRange = sh.getRange(startRow, 1, numRows, nCols);
  var rows = rowRange.getValues();
  var statusHeaders = Object.keys(SECTION_META).map(function(section) {
    return SECTION_META[section].statusCol;
  });
  var changedRows = 0;

  rows.forEach(function(rowVals) {
    if (!hasValue_(rowVals[bci("Record ID")])) return;
    var before = statusHeaders.map(function(h) { return rowVals[bci(h)]; }).join("|");
    updateSectionStatuses_(rowVals);
    var after = statusHeaders.map(function(h) { return rowVals[bci(h)]; }).join("|");
    if (before !== after) changedRows++;
  });

  statusHeaders.forEach(function(header) {
    var col = bci(header);
    if (col < 0) return;
    var values = rows.map(function(rowVals) { return [rowVals[col]]; });
    sh.getRange(startRow, col + 1, numRows, 1).setValues(values);
  });

  return { success: true, updatedRows: changedRows };
}

function isSectionComplete_(rowVals, section) {
  var keys = SECTION_FIELD_KEYS[section] || [];
  if (!keys.length) return false;
  for (var i = 0; i < keys.length; i++) {
    var col = bci(keys[i]);
    if (col < 0 || !hasValue_(rowVals[col])) return false;
  }
  return true;
}

function hasValue_(value) {
  return String(value === null || value === undefined ? "" : value).trim() !== "";
}

function bci(header) {
  return BCI.hasOwnProperty(header) ? BCI[header] : -1;
}

function findRow(sh, recordId) {
  var lr = sh.getLastRow();
  if (lr < 2) return -1;
  var target = String(recordId || "").trim();
  if (!target) return -1;
  var found = sh.getRange(2, 1, lr - 1, 1).createTextFinder(target).matchEntireCell(true).findNext();
  return found ? found.getRow() : -1;
}

function getOrCreateBookingSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BOOKING_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BOOKING_SHEET);
    sh.appendRow(BOOKING_HEADERS);
    sh.setFrozenRows(1);
    styleBookingHeaders(sh);
  } else {
    ensureBookingHeaders_(sh);
  }
  return sh;
}

function ensureBookingHeaders_(sh) {
  var nCols = BOOKING_HEADERS.length;
  if (sh.getMaxColumns() < nCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
  }
  var currentLastCol = Math.max(sh.getLastColumn(), nCols);
  var current = sh.getRange(1, 1, 1, currentLastCol).getValues()[0];
  for (var i = 0; i < nCols; i++) {
    if (String(current[i] || "") !== BOOKING_HEADERS[i]) {
      migrateBookingHeaders_(sh, current);
      return;
    }
  }
}

function migrateBookingHeaders_(sh, currentHeaders, options) {
  options = options || {};
  var lr = Math.max(sh.getLastRow(), 1);
  var oldColCount = Math.max(currentHeaders.length, sh.getLastColumn(), BOOKING_HEADERS.length);
  if (sh.getMaxColumns() < BOOKING_HEADERS.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), BOOKING_HEADERS.length - sh.getMaxColumns());
  }

  var oldRange = sh.getRange(1, 1, lr, oldColCount);
  var oldValues = oldRange.getValues();
  var oldFormulas = oldRange.getFormulas();
  var oldIndex = {};
  currentHeaders.forEach(function(h, idx) {
    h = String(h || "").trim();
    if (h && !oldIndex.hasOwnProperty(h)) oldIndex[h] = idx;
    if (h && HEADER_ALIASES[h] && !oldIndex.hasOwnProperty(HEADER_ALIASES[h])) {
      oldIndex[HEADER_ALIASES[h]] = idx;
    }
  });

  var newData = oldValues.map(function(row, rIdx) {
    if (rIdx === 0) return BOOKING_HEADERS.slice();
    return BOOKING_HEADERS.map(function(h) {
      if (!oldIndex.hasOwnProperty(h)) return "";
      var oldCol = oldIndex[h];
      return oldFormulas[rIdx][oldCol] || row[oldCol];
    });
  });

  var syncedRows = 0;
  for (var r = 1; r < newData.length; r++) {
    var recordId = newData[r][bci("Record ID")];
    if (!hasValue_(recordId)) continue;
    if (options.rebuildLinks) syncRowEditLinks_(newData[r], recordId);
    if (options.refreshStatuses !== false) updateSectionStatuses_(newData[r]);
    syncedRows++;
  }

  sh.getRange(1, 1, lr, BOOKING_HEADERS.length).setValues(newData);
  if (oldColCount > BOOKING_HEADERS.length) {
    sh.getRange(1, BOOKING_HEADERS.length + 1, lr, oldColCount - BOOKING_HEADERS.length).clearContent();
  }
  styleBookingHeaders(sh);
  return { success: true, rowsSynced: syncedRows, columns: BOOKING_HEADERS.length };
}

function styleBookingHeaders(sh) {
  sh.getRange(1, 1, 1, BOOKING_HEADERS.length)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setFontSize(9)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrap(false);

  COL_GROUPS.forEach(function(group) {
    sh.getRange(1, group.s, 1, group.e - group.s + 1).setBackground(group.bg);
  });

  sh.setRowHeight(1, 38);
  sh.setFrozenColumns(1);
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 130);

  for (var i = 4; i <= BOOKING_HEADERS.length; i++) {
    var h = BOOKING_HEADERS[i - 1];
    if (h.indexOf("Edit URL") >= 0) sh.setColumnWidth(i, 185);
    else if (h.indexOf("Status") >= 0) sh.setColumnWidth(i, 110);
    else sh.setColumnWidth(i, 125);
  }
}

function getOrCreateUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.appendRow(USER_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 4)
      .setFontWeight("bold")
      .setBackground("#1e3a5f")
      .setFontColor("#ffffff");
  }
  ensureDefaultUsers_(sh);
  [1,2,3,4].forEach(function(_, idx) {
    sh.setColumnWidth(idx + 1, [200, 140, 100, 150][idx]);
  });
  return sh;
}

function getCachedUserRows_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(USER_CACHE_KEY);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {}
  }

  var sh = getOrCreateUsersSheet();
  var lr = sh.getLastRow();
  var rows = lr ? sh.getRange(1, 1, lr, USER_HEADERS.length).getValues() : [];
  cache.put(USER_CACHE_KEY, JSON.stringify(rows), USER_CACHE_SECONDS);
  return rows;
}

function ensureDefaultUsers_(sh) {
  var lr = sh.getLastRow();
  var rows = lr ? sh.getRange(1, 1, lr, USER_HEADERS.length).getValues() : [];
  var known = {};
  for (var i = 1; i < rows.length; i++) {
    var em = normalizeLoginEmail_(rows[i][0]);
    if (em) known[em] = true;
  }
  var missing = [];
  DEFAULT_USERS.forEach(function(u) {
    var em = normalizeLoginEmail_(u[0]);
    if (!known[em]) {
      missing.push(u);
      known[em] = true;
    }
  });
  if (missing.length) {
    sh.getRange(sh.getLastRow() + 1, 1, missing.length, USER_HEADERS.length).setValues(missing);
    CacheService.getScriptCache().remove(USER_CACHE_KEY);
  }
}

function normalizeLoginEmail_(email) {
  var em = String(email || "").trim().toLowerCase();
  var aliases = {
    "sahib@quickshipnow.com": "sk@quickshipnow.com",
    "sahibkhan@quickshipnow.com": "sk@quickshipnow.com",
    "sahib.khan@quickshipnow.com": "sk@quickshipnow.com"
  };
  return aliases[em] || em;
}

function normalizeRole_(role) {
  return String(role || "").trim().toLowerCase();
}
