// ─── 69 Tracker — Google Apps Script Backend ─────────────────────────────────
// Paste ALL of this into your Apps Script editor, then re-deploy as a Web App
// (new version — bump the deployment so the new code goes live).
//
// STORAGE LAYOUT — one Google Sheet tab per entity:
//
//   Players             → id | name | clan | active | hero | level | might | G | M | S | E
//   LevelRequests       → id | playerId | playerName | clan | levelType | from | to | status | date | resolvedDate
//   RotationLog         → id | playerId | playerName | fromClan | toClan | date
//   FragmentDistributions → id | clan | eventId | date | scoreEntryDate | totalFragments | config | allocations
//   Config              → key | value   (pins, norms, ctSync, ctAliases, ctIgnored, lastBackup)
//
//   Score tabs — one row per player per entry (clan stamped at time of entry):
//   WeeklyChests        → weekStart | entryDate | syncedAt | clan | playerId | points
//   TinMan              → weekStart | entryDate | syncedAt | clan | playerId | points
//   Ragnarok            → weekStart | entryDate | syncedAt | clan | playerId | points
//   Omens               → weekStart | entryDate | syncedAt | clan | playerId | essence | damage | chests
//   Olympus             → weekStart | entryDate | syncedAt | clan | playerId | score | chests
//   EpicChests          → weekStart | entryDate | syncedAt | clan | playerId | score
//
// The frontend JSON protocol is UNCHANGED — the app sends and receives the same
// data shape as before. Only the backend storage format has changed.
//
// FIRST-TIME SETUP AFTER DEPLOYING:
//   1. Run migrateToTabs() once from the Apps Script editor to move your
//      existing cell data into the new tabs.
//   2. Re-deploy the Web App (New version).
// ─────────────────────────────────────────────────────────────────────────────

// ── Tab name constants ────────────────────────────────────────────────────────
var TAB = {
  PLAYERS:               "Players",
  LEVEL_REQUESTS:        "LevelRequests",
  ROTATION_LOG:          "RotationLog",
  FRAGMENT_DISTRIBUTIONS:"FragmentDistributions",
  CONFIG:                "Config",
  WEEKLY_CHESTS:         "WeeklyChests",
  TIN_MAN:               "TinMan",
  RAGNAROK:              "Ragnarok",
  OMENS:                 "Omens",
  OLYMPUS:               "Olympus",
  EPIC_CHESTS:           "EpicChests",
};

// Maps score event id (as used in frontend) → tab name and column definitions
// fields: ordered list of score-specific column keys for this event
var SCORE_EVENTS = {
  "weekly_chests": { tab: TAB.WEEKLY_CHESTS, fields: ["points"] },
  "tin_man":       { tab: TAB.TIN_MAN,       fields: ["points"] },
  "ragnarok":      { tab: TAB.RAGNAROK,       fields: ["points"] },
  "omens":         { tab: TAB.OMENS,          fields: ["essence","damage","chests"] },
  "olympus":       { tab: TAB.OLYMPUS,        fields: ["score","chests"] },
  "epic_chests":   { tab: TAB.EPIC_CHESTS,    fields: ["score"] },
};

// Clans that have score data
var CLANS = ["69R", "69S", "69D"];

// Default PINs — used on first run before any admin changes them.
var DEFAULT_PINS = {
  super: "9999",
  "69R": "6969",
  "69S": "6996",
  "69D": "9669",
  user:  "1111",
};

var EMPTY_DATA = {
  players: [], scores: {}, levelRequests: [], rotationLog: [],
  fragmentDistributions: [], lastBackup: null, pins: DEFAULT_PINS,
  ctSync: {}, ctAliases: {}, ctIgnored: {}, norms: {},
};

// ── Utility ───────────────────────────────────────────────────────────────────
function safeParse(raw, fallback) {
  if (raw === null || raw === undefined || raw === "") return (fallback !== undefined ? fallback : null);
  if (typeof raw !== "string") return raw; // already parsed
  try { return JSON.parse(raw); } catch(_) { return (fallback !== undefined ? fallback : null); }
}

function getOrCreateTab(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

// Read all data rows from a sheet (skips header row 1), returns array of arrays.
function readRows(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
}

// Append rows to a sheet. rows = array of arrays.
function appendRows(sheet, rows) {
  if (!rows || !rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

// Clear all data rows (keep header).
function clearDataRows(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  sheet.getRange(2, 1, last - 1, sheet.getMaxColumns()).clearContent();
}

function jsonResponse(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── Tab headers ───────────────────────────────────────────────────────────────
var HEADERS = {
  PLAYERS:               ["id","name","clan","active","hero","level","might","G","M","S","E"],
  LEVEL_REQUESTS:        ["id","playerId","playerName","clan","levelType","from","to","status","date","resolvedDate"],
  ROTATION_LOG:          ["id","playerId","playerName","fromClan","toClan","date"],
  FRAGMENT_DISTRIBUTIONS:["id","clan","eventId","date","scoreEntryDate","totalFragments","config","allocations"],
  CONFIG:                ["key","value"],
  WEEKLY_CHESTS:         ["weekStart","entryDate","syncedAt","clan","playerId","points"],
  TIN_MAN:               ["weekStart","entryDate","syncedAt","clan","playerId","points"],
  RAGNAROK:              ["weekStart","entryDate","syncedAt","clan","playerId","points"],
  OMENS:                 ["weekStart","entryDate","syncedAt","clan","playerId","essence","damage","chests"],
  OLYMPUS:               ["weekStart","entryDate","syncedAt","clan","playerId","score","chests"],
  EPIC_CHESTS:           ["weekStart","entryDate","syncedAt","clan","playerId","score"],
};

// ── Initialise all tabs (idempotent) ──────────────────────────────────────────
function initTabs(ss) {
  Object.keys(TAB).forEach(function(k) {
    getOrCreateTab(ss, TAB[k], HEADERS[k]);
  });
}

// ── Config helpers ────────────────────────────────────────────────────────────
function readConfig(ss) {
  var sheet = getOrCreateTab(ss, TAB.CONFIG, HEADERS.CONFIG);
  var rows  = readRows(sheet);
  var cfg   = {};
  rows.forEach(function(r) { if (r[0]) cfg[r[0]] = safeParse(r[1], r[1]); });
  return cfg;
}

function writeConfigKey(ss, key, value) {
  var sheet = getOrCreateTab(ss, TAB.CONFIG, HEADERS.CONFIG);
  var rows  = readRows(sheet);
  var found = false;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) {
      // Update in place — row index in sheet = i + 2 (1-indexed + header)
      sheet.getRange(i + 2, 2).setValue(JSON.stringify(value));
      found = true;
      break;
    }
  }
  if (!found) {
    appendRows(sheet, [[key, JSON.stringify(value)]]);
  }
}

function writeAllConfig(ss, cfg) {
  Object.keys(cfg).forEach(function(k) { writeConfigKey(ss, k, cfg[k]); });
}

// ── Players ───────────────────────────────────────────────────────────────────
var PLAYER_COLS = HEADERS.PLAYERS; // ["id","name","clan","active","hero","level","might","G","M","S","E"]

function readPlayers(ss) {
  var sheet = getOrCreateTab(ss, TAB.PLAYERS, PLAYER_COLS);
  var rows  = readRows(sheet);
  return rows.filter(function(r) { return r[0]; }).map(function(r) {
    return {
      id:     r[0],
      name:   r[1],
      clan:   r[2],
      active: r[3] === true || r[3] === "TRUE" || r[3] === 1,
      hero:   r[4] !== "" && r[4] !== null ? Number(r[4]) : null,
      level:  r[5] !== "" && r[5] !== null ? Number(r[5]) : null,
      might:  r[6] !== "" && r[6] !== null ? Number(r[6]) : null,
      levels: { G: r[7] || null, M: r[8] || null, S: r[9] || null, E: r[10] || null },
    };
  });
}

function writePlayers(ss, players) {
  var sheet = getOrCreateTab(ss, TAB.PLAYERS, PLAYER_COLS);
  clearDataRows(sheet);
  if (!players || !players.length) return;
  var rows = players.map(function(p) {
    var lv = p.levels || {};
    return [
      p.id   || "",
      p.name || "",
      p.clan || "",
      p.active ? "TRUE" : "FALSE",
      p.hero   !== null && p.hero   !== undefined ? p.hero   : "",
      p.level  !== null && p.level  !== undefined ? p.level  : "",
      p.might  !== null && p.might  !== undefined ? p.might  : "",
      lv.G || "",
      lv.M || "",
      lv.S || "",
      lv.E || "",
    ];
  });
  appendRows(sheet, rows);
}

// ── Level Requests ────────────────────────────────────────────────────────────
var LR_COLS = HEADERS.LEVEL_REQUESTS;

function readLevelRequests(ss) {
  var sheet = getOrCreateTab(ss, TAB.LEVEL_REQUESTS, LR_COLS);
  var rows  = readRows(sheet);
  return rows.filter(function(r) { return r[0]; }).map(function(r) {
    return {
      id:           r[0],
      playerId:     r[1],
      playerName:   r[2],
      clan:         r[3],
      levelType:    r[4],
      from:         r[5],
      to:           r[6],
      status:       r[7],
      date:         r[8],
      resolvedDate: r[9] || null,
    };
  });
}

function writeLevelRequests(ss, requests) {
  var sheet = getOrCreateTab(ss, TAB.LEVEL_REQUESTS, LR_COLS);
  clearDataRows(sheet);
  if (!requests || !requests.length) return;
  var rows = requests.map(function(r) {
    return [
      r.id          || "",
      r.playerId    || "",
      r.playerName  || "",
      r.clan        || "",
      r.levelType   || "",
      r.from        !== undefined ? r.from : "",
      r.to          !== undefined ? r.to   : "",
      r.status      || "",
      r.date        || "",
      r.resolvedDate || "",
    ];
  });
  appendRows(sheet, rows);
}

// ── Rotation Log ──────────────────────────────────────────────────────────────
var RL_COLS = HEADERS.ROTATION_LOG;

function readRotationLog(ss) {
  var sheet = getOrCreateTab(ss, TAB.ROTATION_LOG, RL_COLS);
  var rows  = readRows(sheet);
  return rows.filter(function(r) { return r[0]; }).map(function(r) {
    return {
      id:         r[0],
      playerId:   r[1],
      playerName: r[2],
      fromClan:   r[3],
      toClan:     r[4],
      date:       r[5],
    };
  });
}

function writeRotationLog(ss, log) {
  var sheet = getOrCreateTab(ss, TAB.ROTATION_LOG, RL_COLS);
  clearDataRows(sheet);
  if (!log || !log.length) return;
  var rows = log.map(function(r) {
    return [r.id||"", r.playerId||"", r.playerName||"", r.fromClan||"", r.toClan||"", r.date||""];
  });
  appendRows(sheet, rows);
}

// ── Fragment Distributions ────────────────────────────────────────────────────
var FD_COLS = HEADERS.FRAGMENT_DISTRIBUTIONS;

function readFragmentDistributions(ss) {
  var sheet = getOrCreateTab(ss, TAB.FRAGMENT_DISTRIBUTIONS, FD_COLS);
  var rows  = readRows(sheet);
  return rows.filter(function(r) { return r[0]; }).map(function(r) {
    return {
      id:             r[0],
      clan:           r[1],
      eventId:        r[2],
      date:           r[3],
      scoreEntryDate: r[4],
      totalFragments: r[5] !== "" ? Number(r[5]) : null,
      config:         safeParse(r[6], {}),
      allocations:    safeParse(r[7], []),
    };
  });
}

function writeFragmentDistributions(ss, dists) {
  var sheet = getOrCreateTab(ss, TAB.FRAGMENT_DISTRIBUTIONS, FD_COLS);
  clearDataRows(sheet);
  if (!dists || !dists.length) return;
  var rows = dists.map(function(d) {
    return [
      d.id             || "",
      d.clan           || "",
      d.eventId        || "",
      d.date           || "",
      d.scoreEntryDate || "",
      d.totalFragments !== undefined ? d.totalFragments : "",
      JSON.stringify(d.config      || {}),
      JSON.stringify(d.allocations || []),
    ];
  });
  appendRows(sheet, rows);
}

// ── Score tabs ────────────────────────────────────────────────────────────────
// Scores are stored as one row per player per entry.
// Each score entry in the frontend looks like:
//   { date, weekStart, syncedAt, scores: { playerId: { field1: val, field2: val } } }
// We flatten this into one row per player:
//   weekStart | entryDate | syncedAt | clan | playerId | field1 | field2 ...

function readScoreTab(ss, eventId) {
  var cfg   = SCORE_EVENTS[eventId];
  if (!cfg) return [];
  var sheet = getOrCreateTab(ss, cfg.tab, HEADERS[cfg.tab.toUpperCase().replace(" ","_")] || buildScoreHeaders(cfg.fields));
  var rows  = readRows(sheet);
  if (!rows.length) return [];

  // Group rows by weekStart+clan into entries
  var entryMap = {}; // key: clan+"||"+weekStart → { date, weekStart, syncedAt, clan, scores:{} }
  rows.forEach(function(r) {
    var weekStart = r[0] ? String(r[0]).slice(0,10) : "";
    var entryDate = r[1] ? String(r[1]) : "";
    var syncedAt  = r[2] ? String(r[2]) : "";
    var clan      = r[3] ? String(r[3]) : "";
    var playerId  = r[4] ? String(r[4]) : "";
    if (!weekStart || !clan || !playerId) return;

    var key = clan + "||" + weekStart;
    if (!entryMap[key]) {
      entryMap[key] = {
        date:      entryDate || (weekStart + "T00:00:00.000Z"),
        weekStart: weekStart,
        syncedAt:  syncedAt,
        clan:      clan,
        scores:    {},
      };
    }
    // Build per-player score object from field columns (index 5+)
    var scoreObj = {};
    cfg.fields.forEach(function(f, fi) {
      var val = r[5 + fi];
      scoreObj[f] = val !== "" && val !== null && val !== undefined ? Number(val) : 0;
    });
    entryMap[key].scores[playerId] = scoreObj;
  });

  // Return as frontend expects: { "69R_weekly_chests": [...entries] }
  // Group by clan, sort newest first
  var result = {}; // clanKey → entries array
  Object.keys(entryMap).forEach(function(k) {
    var entry = entryMap[k];
    var scoreKey = entry.clan + "_" + eventId;
    if (!result[scoreKey]) result[scoreKey] = [];
    result[scoreKey].push(entry);
  });
  // Sort each clan's entries newest first
  Object.keys(result).forEach(function(k) {
    result[k].sort(function(a, b) { return a.weekStart < b.weekStart ? 1 : -1; });
  });
  return result; // { "69R_weekly_chests": [...], "69S_weekly_chests": [...] }
}

function writeScoreTab(ss, eventId, scoresByKey) {
  // scoresByKey: { "69R_weekly_chests": [{date, weekStart, syncedAt, scores:{pid:{field:val}}}], ... }
  var cfg   = SCORE_EVENTS[eventId];
  if (!cfg) return;
  var tabName = cfg.tab;
  var headerKey = tabName.replace(/([A-Z])/g, function(m,l,i) { return (i>0?"_":"")+l; }).toUpperCase();
  var headers = HEADERS[tabName.toUpperCase()] || HEADERS[headerKey] || buildScoreHeaders(cfg.fields);
  var sheet   = getOrCreateTab(ss, tabName, headers);
  clearDataRows(sheet);

  var newRows = [];
  CLANS.forEach(function(clan) {
    var scoreKey = clan + "_" + eventId;
    var entries  = scoresByKey[scoreKey] || [];
    entries.forEach(function(entry) {
      var weekStart = entry.weekStart || (entry.date || "").slice(0,10);
      var entryDate = entry.date      || weekStart + "T00:00:00.000Z";
      var syncedAt  = entry.syncedAt  || "";
      var scores    = entry.scores    || {};
      Object.keys(scores).forEach(function(pid) {
        var s = scores[pid] || {};
        var row = [weekStart, entryDate, syncedAt, clan, pid];
        cfg.fields.forEach(function(f) {
          row.push(s[f] !== undefined && s[f] !== null ? s[f] : 0);
        });
        newRows.push(row);
      });
    });
  });

  if (newRows.length) appendRows(sheet, newRows);
}

// Fallback header builder (in case tab name doesn't map cleanly)
function buildScoreHeaders(fields) {
  return ["weekStart","entryDate","syncedAt","clan","playerId"].concat(fields);
}

// HEADERS lookup by tab name (handles camelCase tab names)
function getHeadersForTab(tabName) {
  var map = {
    "WeeklyChests": HEADERS.WEEKLY_CHESTS,
    "TinMan":       HEADERS.TIN_MAN,
    "Ragnarok":     HEADERS.RAGNAROK,
    "Omens":        HEADERS.OMENS,
    "Olympus":      HEADERS.OLYMPUS,
    "EpicChests":   HEADERS.EPIC_CHESTS,
  };
  return map[tabName] || null;
}

// ── doGet ─────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "syncEpicChests") {
      return jsonResponse(syncEpicChests());
    }

    if (action === "savePins") {
      var pinsStr = e.parameter.pins;
      if (!pinsStr) return jsonResponse({ error: "No pins data provided" });
      writeConfigKey(ss, "pins", JSON.parse(pinsStr));
      Logger.log("savePins: updated Config tab");
      return jsonResponse({ ok: true });
    }

    if (action === "saveNorms") {
      var normsStr = e.parameter.norms;
      if (!normsStr) return jsonResponse({ error: "No norms data provided" });
      writeConfigKey(ss, "norms", JSON.parse(normsStr));
      Logger.log("saveNorms: updated Config tab");
      return jsonResponse({ ok: true });
    }

    if (action === "saveAlias") {
      var clan   = e.parameter.clan;
      var ctName = e.parameter.ctName;
      var pid    = e.parameter.playerId;
      if (!clan || !ctName || !pid) return jsonResponse({ error: "Missing clan, ctName, or playerId" });
      var cfg     = readConfig(ss);
      var aliases = cfg.ctAliases || {};
      if (!aliases[clan]) aliases[clan] = {};
      aliases[clan][ctName] = pid;
      // Remove from ignored if present
      var ignored = cfg.ctIgnored || {};
      if (ignored[clan]) ignored[clan] = (ignored[clan] || []).filter(function(n) { return n !== ctName; });
      writeConfigKey(ss, "ctAliases", aliases);
      writeConfigKey(ss, "ctIgnored", ignored);
      Logger.log("saveAlias: " + clan + " [" + ctName + "] → " + pid);
      return jsonResponse({ ok: true });
    }

    if (action === "removeAlias") {
      var clan   = e.parameter.clan;
      var ctName = e.parameter.ctName;
      if (!clan || !ctName) return jsonResponse({ error: "Missing clan or ctName" });
      var cfg     = readConfig(ss);
      var aliases = cfg.ctAliases || {};
      if (aliases[clan]) delete aliases[clan][ctName];
      writeConfigKey(ss, "ctAliases", aliases);
      Logger.log("removeAlias: " + clan + " [" + ctName + "]");
      return jsonResponse({ ok: true });
    }

    if (action === "saveIgnored") {
      var clan   = e.parameter.clan;
      var ctName = e.parameter.ctName;
      var remove = e.parameter.remove === "1";
      if (!clan || !ctName) return jsonResponse({ error: "Missing clan or ctName" });
      var cfg     = readConfig(ss);
      var ignored = cfg.ctIgnored || {};
      var aliases = cfg.ctAliases || {};
      if (!ignored[clan]) ignored[clan] = [];
      if (remove) {
        ignored[clan] = ignored[clan].filter(function(n) { return n !== ctName; });
        Logger.log("saveIgnored: un-ignored " + clan + " [" + ctName + "]");
      } else {
        if (ignored[clan].indexOf(ctName) === -1) ignored[clan].push(ctName);
        if (aliases[clan]) delete aliases[clan][ctName];
        Logger.log("saveIgnored: ignored " + clan + " [" + ctName + "]");
      }
      writeConfigKey(ss, "ctIgnored", ignored);
      writeConfigKey(ss, "ctAliases", aliases);
      return jsonResponse({ ok: true });
    }

    return jsonResponse(readData(ss));
  }
  catch (err) { return jsonResponse({ error: err.message }); }
}

// ── doPost ────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var incoming = JSON.parse(e.postData.contents);
    if (incoming.action === "syncEpicChests") {
      return jsonResponse(syncEpicChests());
    }
    if (incoming.action) {
      return jsonResponse({ error: "Unknown action: " + incoming.action });
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    writeData(ss, incoming);
    return jsonResponse({ ok: true });
  } catch (err) { return jsonResponse({ error: err.message }); }
}

// ── readData ──────────────────────────────────────────────────────────────────
function readData(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  initTabs(ss);

  var players              = readPlayers(ss);
  var levelRequests        = readLevelRequests(ss);
  var rotationLog          = readRotationLog(ss);
  var fragmentDistributions = readFragmentDistributions(ss);
  var cfg                  = readConfig(ss);

  // Merge all score events into one scores object
  var scores = {};
  Object.keys(SCORE_EVENTS).forEach(function(eventId) {
    var eventScores = readScoreTab(ss, eventId); // { "69R_weekly_chests": [...], ... }
    Object.keys(eventScores).forEach(function(k) { scores[k] = eventScores[k]; });
  });

  return {
    players:               players,
    scores:                scores,
    levelRequests:         levelRequests,
    rotationLog:           rotationLog,
    fragmentDistributions: fragmentDistributions,
    lastBackup:            cfg.lastBackup  || null,
    pins:                  cfg.pins        || DEFAULT_PINS,
    ctSync:                cfg.ctSync      || {},
    ctAliases:             cfg.ctAliases   || {},
    ctIgnored:             cfg.ctIgnored   || {},
    norms:                 cfg.norms       || {},
  };
}

// ── writeData ─────────────────────────────────────────────────────────────────
function writeData(ss, data) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!data || (data.action && !data.players && !data.scores && !data.levelRequests && !data.pins && !data.ctSync)) {
    throw new Error("writeData: invalid payload — missing required data fields. Keys: " + JSON.stringify(Object.keys(data || {})));
  }
  initTabs(ss);

  // PIN-only save
  if (data.pins && !data.players && !data.scores) {
    writeConfigKey(ss, "pins", data.pins);
    return;
  }

  if (data.players !== undefined)              writePlayers(ss, data.players);
  if (data.levelRequests !== undefined)        writeLevelRequests(ss, data.levelRequests);
  if (data.rotationLog !== undefined)          writeRotationLog(ss, data.rotationLog);
  if (data.fragmentDistributions !== undefined) writeFragmentDistributions(ss, data.fragmentDistributions);

  // Config keys — only update keys present in the payload
  if (data.pins        !== undefined) writeConfigKey(ss, "pins",       data.pins);
  if (data.lastBackup  !== undefined) writeConfigKey(ss, "lastBackup", data.lastBackup);
  if (data.norms       !== undefined) writeConfigKey(ss, "norms",      data.norms);

  // CT metadata — only update when CT-related keys are present in the payload
  if (data.ctSync !== undefined || data.ctAliases !== undefined || data.ctIgnored !== undefined) {
    var existingCfg = readConfig(ss);
    writeConfigKey(ss, "ctSync",    data.ctSync    !== undefined ? data.ctSync    : (existingCfg.ctSync    || {}));
    writeConfigKey(ss, "ctAliases", data.ctAliases !== undefined ? data.ctAliases : (existingCfg.ctAliases || {}));
    writeConfigKey(ss, "ctIgnored", data.ctIgnored !== undefined ? data.ctIgnored : (existingCfg.ctIgnored || {}));
  }

  // Scores — write each event's tab from the incoming scores object
  if (data.scores !== undefined) {
    Object.keys(SCORE_EVENTS).forEach(function(eventId) {
      // Collect all score keys for this event across all clans
      var relevant = {};
      CLANS.forEach(function(clan) {
        var k = clan + "_" + eventId;
        if (data.scores[k] !== undefined) relevant[k] = data.scores[k];
      });
      // Only rewrite the tab if at least one clan's data is present in the payload
      if (Object.keys(relevant).length > 0) {
        // Merge with existing data for clans not in the payload
        var existing = readScoreTab(ss, eventId);
        var merged = {};
        CLANS.forEach(function(clan) {
          var k = clan + "_" + eventId;
          merged[k] = relevant[k] !== undefined ? relevant[k] : (existing[k] || []);
        });
        writeScoreTab(ss, eventId, merged);
      }
    });
  }
}

// ── Migration from old cell-based layout ──────────────────────────────────────
// Run this ONCE from the Apps Script editor after deploying this new backend.
// Reads ALL old data in one batch call, writes each tab in one setValues call.
// Safe to run multiple times — it reads live data each time.
function migrateToTabs() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var oldSheet = ss.getSheetByName("AppData");

  if (!oldSheet) {
    Logger.log("migrateToTabs: No 'AppData' sheet found — nothing to migrate.");
    return;
  }

  Logger.log("migrateToTabs: Reading AppData in one batch...");

  // ── Read ALL of column A in one API call ──────────────────────────────────
  var lastRow  = Math.max(oldSheet.getLastRow(), 30);
  var rawCol   = oldSheet.getRange(1, 1, lastRow, 1).getValues(); // one API call
  var parse    = function(v) { try { return v ? JSON.parse(v) : {}; } catch(_) { return {}; } };
  var cell     = function(row) { return rawCol[row - 1] ? rawCol[row - 1][0] : ""; };

  var dataA1  = parse(cell(1));
  var dataA2  = parse(cell(2));
  var dataA21 = parse(cell(21));
  var dataA23 = parse(cell(23));
  var dataA24 = parse(cell(24));
  var normsData = cell(25) ? parse(cell(25)) : {};

  var players = (dataA1.players || []).concat(dataA23.players || []);

  var OLD_SCORE_ROW_MAP = {
    "69R_weekly_chests": 3,  "69R_tin_man": 4,  "69R_ragnarok": 5,
    "69R_omens": 7,          "69R_olympus": 8,
    "69S_weekly_chests": 9,  "69S_tin_man": 10, "69S_ragnarok": 11,
    "69S_omens": 13,         "69S_olympus": 14,
    "69D_weekly_chests": 15, "69D_tin_man": 16, "69D_ragnarok": 17,
    "69D_omens": 19,         "69D_olympus": 20,
    "69R_epic_chests": 22,   "69S_epic_chests": 26,
  };

  var oldScores = {};
  Object.keys(OLD_SCORE_ROW_MAP).forEach(function(key) {
    var raw = cell(OLD_SCORE_ROW_MAP[key]);
    if (raw) {
      var d = parse(raw);
      if (d.scores && d.scores[key]) oldScores[key] = d.scores[key];
    }
  });

  Logger.log("migrateToTabs: " + players.length + " players | " +
    (dataA2.levelRequests||[]).length + " LRs | " +
    Object.keys(oldScores).length + " score keys");

  // ── Helper: create tab + write all rows in ONE setValues call ─────────────
  function writeTabBatch(tabName, headers, rows) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
    } else {
      sheet.clearContents(); // wipe everything including any partial data from a previous failed run
    }
    // Always write headers to row 1
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows && rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    SpreadsheetApp.flush(); // commit after each tab — preserves progress on partial timeout
    Logger.log("  ✓ " + tabName + " — " + (rows ? rows.length : 0) + " rows");
  }

  // ── Players ───────────────────────────────────────────────────────────────
  var playerRows = players.map(function(p) {
    var lv = p.levels || {};
    return [p.id||"", p.name||"", p.clan||"", p.active?"TRUE":"FALSE",
      p.hero  !==null&&p.hero  !==undefined ? p.hero   : "",
      p.level !==null&&p.level !==undefined ? p.level  : "",
      p.might !==null&&p.might !==undefined ? p.might  : "",
      lv.G||"", lv.M||"", lv.S||"", lv.E||""];
  });
  writeTabBatch("Players", HEADERS.PLAYERS, playerRows);

  // ── Level Requests ────────────────────────────────────────────────────────
  var lrRows = (dataA2.levelRequests || []).map(function(r) {
    return [r.id||"", r.playerId||"", r.playerName||"", r.clan||"",
      r.levelType||"", r.from!==undefined?r.from:"", r.to!==undefined?r.to:"",
      r.status||"", r.date||"", r.resolvedDate||""];
  });
  writeTabBatch("LevelRequests", HEADERS.LEVEL_REQUESTS, lrRows);

  // ── Rotation Log ──────────────────────────────────────────────────────────
  var rlRows = (dataA2.rotationLog || []).map(function(r) {
    return [r.id||"", r.playerId||"", r.playerName||"", r.fromClan||"", r.toClan||"", r.date||""];
  });
  writeTabBatch("RotationLog", HEADERS.ROTATION_LOG, rlRows);

  // ── Fragment Distributions ────────────────────────────────────────────────
  var fdRows = (dataA2.fragmentDistributions || []).map(function(d) {
    return [d.id||"", d.clan||"", d.eventId||"", d.date||"", d.scoreEntryDate||"",
      d.totalFragments!==undefined?d.totalFragments:"",
      JSON.stringify(d.config||{}), JSON.stringify(d.allocations||[])];
  });
  writeTabBatch("FragmentDistributions", HEADERS.FRAGMENT_DISTRIBUTIONS, fdRows);

  // ── Config ────────────────────────────────────────────────────────────────
  var cfgRows = [
    ["pins",       JSON.stringify(dataA21.pins     || DEFAULT_PINS)],
    ["norms",      JSON.stringify(normsData)],
    ["ctSync",     JSON.stringify(dataA24.ctSync   || {})],
    ["ctAliases",  JSON.stringify(dataA24.ctAliases|| {})],
    ["ctIgnored",  JSON.stringify(dataA24.ctIgnored|| {})],
    ["lastBackup", JSON.stringify(dataA1.lastBackup|| null)],
  ];
  writeTabBatch("Config", HEADERS.CONFIG, cfgRows);

  // ── Score tabs ────────────────────────────────────────────────────────────
  Object.keys(SCORE_EVENTS).forEach(function(eventId) {
    var cfg     = SCORE_EVENTS[eventId];
    var headers = getHeadersForTab(cfg.tab) || buildScoreHeaders(cfg.fields);
    var newRows = [];

    CLANS.forEach(function(clan) {
      var entries = oldScores[clan + "_" + eventId] || [];
      entries.forEach(function(entry) {
        var weekStart = entry.weekStart || (entry.date||"").slice(0,10);
        var entryDate = entry.date      || weekStart + "T00:00:00.000Z";
        var syncedAt  = entry.syncedAt  || "";
        Object.keys(entry.scores || {}).forEach(function(pid) {
          var s   = entry.scores[pid] || {};
          var row = [weekStart, entryDate, syncedAt, clan, pid];
          cfg.fields.forEach(function(f) { row.push(s[f] !== undefined ? s[f] : 0); });
          newRows.push(row);
        });
      });
    });

    writeTabBatch(cfg.tab, headers, newRows);
  });

  Logger.log("migrateToTabs: COMPLETE — all data written to tabs.");
  Logger.log("Verify the tabs look correct, then re-deploy the Web App as a new version.");
}

// ── Diagnostics ───────────────────────────────────────────────────────────────
function diagnose() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("=== Tab Diagnostics ===");
  Object.keys(TAB).forEach(function(k) {
    var sheet = ss.getSheetByName(TAB[k]);
    if (!sheet) { Logger.log(TAB[k] + ": NOT FOUND"); return; }
    var rows = sheet.getLastRow() - 1;
    Logger.log(TAB[k] + ": " + (rows < 0 ? 0 : rows) + " data rows");
  });
  Logger.log("=== Config keys ===");
  var cfg = readConfig(ss);
  Object.keys(cfg).forEach(function(k) {
    var v = cfg[k];
    var summary = typeof v === "object" ? JSON.stringify(v).slice(0,80) : String(v).slice(0,80);
    Logger.log(k + ": " + summary);
  });
}

// ─── CHESTTRACKER API SYNC ────────────────────────────────────────────────────
// To configure: go to Apps Script → Project Settings → Script Properties and add:
//   CT_EMAIL    → your ChestTracker login email
//   CT_PASSWORD → your ChestTracker login password
//
// To set up the hourly trigger: run setupEpicChestsTrigger() once from
// the Apps Script editor.
// ─────────────────────────────────────────────────────────────────────────────

var CT_API_BASE = "https://api.chesttracker.com/v1";

function getCTToken(memberId) {
  var props    = PropertiesService.getScriptProperties();
  var email    = props.getProperty("CT_EMAIL");
  var password = props.getProperty("CT_PASSWORD");
  if (!email || !password) {
    throw new Error("ChestTracker credentials not configured. Add CT_EMAIL and CT_PASSWORD in Script Properties.");
  }
  var body = { email: email, password: password };
  if (memberId) body.memberId = memberId;
  var resp = UrlFetchApp.fetch(CT_API_BASE + "/authenticate", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error("ChestTracker auth failed (" + resp.getResponseCode() + "): " + resp.getContentText());
  }
  var result = JSON.parse(resp.getContentText());
  var token = result.authToken || result.token || result.accessToken || result.access_token || result.jwt;
  if (!token) {
    throw new Error("No token in ChestTracker auth response. Keys: " + JSON.stringify(Object.keys(result)));
  }
  return token;
}

function syncEpicChests() {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();

    // Week boundary: Sunday 18:00 UTC → following Sunday 17:59 UTC
    var dayOfWeek = now.getUTCDay();
    var daysBack  = dayOfWeek === 0 && now.getUTCHours() < 18 ? 7 : dayOfWeek;
    var startOfWeek = new Date(now);
    startOfWeek.setUTCDate(now.getUTCDate() - daysBack);
    startOfWeek.setUTCHours(18, 0, 0, 0);
    var weekStart = startOfWeek.toISOString().slice(0, 10);

    var timeParams = "?levels=1"
                   + "&start=" + encodeURIComponent(startOfWeek.toISOString())
                   + "&end="   + encodeURIComponent(now.toISOString());

    // Base auth for listing endpoints
    var baseToken   = getCTToken();
    var baseHeaders = { "Authorization": "Bearer " + baseToken };

    // Get clan tag → id map
    var clansRaw  = JSON.parse(UrlFetchApp.fetch(CT_API_BASE + "/clans", { headers: baseHeaders, muteHttpExceptions: true }).getContentText());
    var clansList = Array.isArray(clansRaw[0]) ? clansRaw[0] : clansRaw;
    var tagToClanId = {};
    clansList.forEach(function(c) { if (c.tag && c.id) tagToClanId[c.tag] = c.id; });

    // Get account's per-clan member IDs
    var membersRaw   = JSON.parse(UrlFetchApp.fetch(CT_API_BASE + "/members", { headers: baseHeaders, muteHttpExceptions: true }).getContentText());
    var acctMembers  = Array.isArray(membersRaw[0]) ? membersRaw[0] : membersRaw;
    var clanIdToMemberId = {};
    acctMembers.forEach(function(m) { if (m.clanId && m.id) clanIdToMemberId[m.clanId] = m.id; });
    var tagToMemberId = {};
    Object.keys(tagToClanId).forEach(function(tag) {
      var clanId = tagToClanId[tag];
      if (clanIdToMemberId[clanId]) tagToMemberId[tag] = clanIdToMemberId[clanId];
    });
    Logger.log("syncEpicChests: clans=" + JSON.stringify(tagToClanId) + " memberIds=" + JSON.stringify(tagToMemberId));

    // Read current data
    var appData = readData(ss);
    if (!appData.scores)  appData.scores  = {};
    if (!appData.ctSync)  appData.ctSync  = {};

    var clansToSync = ["69R", "69S"];
    var results = {};

    clansToSync.forEach(function(clan) {
      var memberId = tagToMemberId[clan];
      if (!memberId) {
        Logger.log("syncEpicChests: no CT memberId for " + clan + " — skipping");
        results[clan] = { matched: 0, unmatched: 0 };
        return;
      }

      var clanToken   = getCTToken(memberId);
      var clanHeaders = { "Authorization": "Bearer " + clanToken };

      var url  = CT_API_BASE + "/chests/breakdown" + timeParams;
      var resp = UrlFetchApp.fetch(url, { headers: clanHeaders, muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) {
        Logger.log("syncEpicChests: breakdown failed for " + clan + " (" + resp.getResponseCode() + ")");
        results[clan] = { matched: 0, unmatched: 0 };
        return;
      }

      var parsed  = JSON.parse(resp.getContentText());
      var members = Array.isArray(parsed[0]) ? parsed[0] : parsed;
      Logger.log("syncEpicChests " + clan + ": CT returned " + members.length + " members");

      var players = (appData.players || []).filter(function(p) { return p.clan === clan && p.active; });
      var aliases = (appData.ctAliases || {})[clan] || {};
      var ignored = (appData.ctIgnored || {})[clan] || [];

      var nameLookup = {};
      players.forEach(function(p) { nameLookup[p.name.toLowerCase().trim()] = p.id; });

      var matched   = {};
      var unmatched = [];

      members.forEach(function(member) {
        var ctName = (member.name || "").trim();
        if (!ctName) return;
        if (ignored.indexOf(ctName) !== -1) return;
        var epicSquad = member["epic squad"];
        var epics = (epicSquad && typeof epicSquad.chests === "number") ? epicSquad.chests : 0;
        var playerId = aliases[ctName] || nameLookup[ctName.toLowerCase().trim()];
        if (playerId) {
          matched[playerId] = epics;
        } else {
          unmatched.push({ name: ctName, epics: epics });
        }
      });

      // Build scores for this week's entry
      var scores = {};
      Object.keys(matched).forEach(function(pid) { scores[pid] = { score: matched[pid] }; });

      var entry = {
        date:      weekStart + "T00:00:00.000Z",
        weekStart: weekStart,
        syncedAt:  now.toISOString(),
        scores:    scores,
      };

      var key = clan + "_epic_chests";
      if (!appData.scores[key]) appData.scores[key] = [];
      var entries = appData.scores[key];
      if (entries.length > 0 && entries[0].weekStart === weekStart) {
        // Merge: keep existing scores for players not returned by CT this sync
        // (avoids wiping unmatched players who had a score from a previous run)
        entry.scores = Object.assign({}, entries[0].scores, entry.scores);
        entries[0] = entry;
      } else {
        entries.unshift(entry);
        if (entries.length > 3) entries = entries.slice(0, 3);
      }
      appData.scores[key] = entries;

      appData.ctSync[clan] = {
        lastSync:  now.toISOString(),
        matched:   Object.keys(matched).length,
        unmatched: unmatched,
      };

      results[clan] = { matched: Object.keys(matched).length, unmatched: unmatched.length };
      Logger.log("syncEpicChests " + clan + ": week=" + weekStart + " matched=" + Object.keys(matched).length + " unmatched=" + unmatched.length);
    });

    writeData(ss, appData);
    return { success: true, "69R": results["69R"], "69S": results["69S"] };

  } catch (err) {
    Logger.log("syncEpicChests ERROR: " + err.message);
    return { success: false, error: err.message };
  }
}

// Run this from the Apps Script editor to see exactly what ChestTracker returns.
// Look for "dman" (or any name) in the log to see their raw API values.
function debugCtRawData() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();

  var dayOfWeek   = now.getUTCDay();
  var daysBack    = dayOfWeek === 0 && now.getUTCHours() < 18 ? 7 : dayOfWeek;
  var startOfWeek = new Date(now);
  startOfWeek.setUTCDate(now.getUTCDate() - daysBack);
  startOfWeek.setUTCHours(18, 0, 0, 0);

  var timeParams = "?levels=1"
                 + "&start=" + encodeURIComponent(startOfWeek.toISOString())
                 + "&end="   + encodeURIComponent(now.toISOString());

  Logger.log("Week window: " + startOfWeek.toISOString() + " → " + now.toISOString());

  var baseToken   = getCTToken();
  var baseHeaders = { "Authorization": "Bearer " + baseToken };

  var clansRaw    = JSON.parse(UrlFetchApp.fetch(CT_API_BASE + "/clans",   { headers: baseHeaders, muteHttpExceptions: true }).getContentText());
  var clansList   = Array.isArray(clansRaw[0])  ? clansRaw[0]  : clansRaw;
  var membersRaw  = JSON.parse(UrlFetchApp.fetch(CT_API_BASE + "/members", { headers: baseHeaders, muteHttpExceptions: true }).getContentText());
  var acctMembers = Array.isArray(membersRaw[0]) ? membersRaw[0] : membersRaw;

  var tagToClanId = {};
  clansList.forEach(function(c) { if (c.tag && c.id) tagToClanId[c.tag] = c.id; });
  var clanIdToMemberId = {};
  acctMembers.forEach(function(m) { if (m.clanId && m.id) clanIdToMemberId[m.clanId] = m.id; });

  ["69R", "69S"].forEach(function(clan) {
    var clanId   = tagToClanId[clan];
    var memberId = clanId && clanIdToMemberId[clanId];
    if (!memberId) { Logger.log(clan + ": no member ID"); return; }

    var clanToken = getCTToken(memberId);
    var resp = UrlFetchApp.fetch(CT_API_BASE + "/chests/breakdown" + timeParams,
                                 { headers: { "Authorization": "Bearer " + clanToken }, muteHttpExceptions: true });
    var members = JSON.parse(resp.getContentText());
    if (Array.isArray(members[0])) members = members[0];

    Logger.log("=== " + clan + " (" + members.length + " members) ===");
    members.forEach(function(m) {
      var epicSquad = m["epic squad"] || {};
      Logger.log(m.name + " | epic squad chests: " + epicSquad.chests + " | raw epic squad: " + JSON.stringify(epicSquad));
    });
  });
}

function setupEpicChestsTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "syncEpicChests") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("syncEpicChests").timeBased().everyHours(1).create();
  Logger.log("Epic Chests trigger set: every hour.");
}
