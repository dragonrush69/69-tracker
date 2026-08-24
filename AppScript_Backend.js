// ─── 69 Tracker — Google Apps Script Backend ─────────────────────────────────
// Paste ALL of this into your Apps Script editor, then re-deploy as a Web App
// (new version — bump the deployment so the new code goes live).
//
// Storage layout — each score key gets its own cell (column A, rows 1-21):
//   A1  → { players, lastBackup }
//   A2  → { levelRequests, rotationLog, fragmentDistributions }
//   A3  → { scores: { 69R_weekly_chests } }
//   A4  → { scores: { 69R_tin_man } }
//   A5  → { scores: { 69R_ragnarok } }
//   A6  → { scores: { 69R_armageddon } }  ← RETIRED (data kept, not read/written)
//   A7  → { scores: { 69R_omens } }
//   A8  → { scores: { 69R_olympus } }
//   A9  → { scores: { 69S_weekly_chests } }
//   A10 → { scores: { 69S_tin_man } }
//   A11 → { scores: { 69S_ragnarok } }
//   A12 → { scores: { 69S_armageddon } }  ← RETIRED (data kept, not read/written)
//   A13 → { scores: { 69S_omens } }
//   A14 → { scores: { 69S_olympus } }
//   A15 → { scores: { 69D_weekly_chests } }
//   A16 → { scores: { 69D_tin_man } }
//   A17 → { scores: { 69D_ragnarok } }
//   A18 → { scores: { 69D_armageddon } }  ← RETIRED (data kept, not read/written)
//   A19 → { scores: { 69D_omens } }
//   A20 → { scores: { 69D_olympus } }
//   A21 → { pins: { super, 69R, 69S, 69D, user } }
//   A22 → { scores: { 69R_epic_chests } }
//   A24 → { ctSync: { 69R: { lastSync, matched, unmatched } }, ctAliases: { 69R: { "CT Name": "playerId" } } }
//   A25 → { "69R": { "epic_chests": { "T9": 500, "G9": 400 }, ... }, ... }  ← performance norms
//
// Each cell stays well under Google's 50,000 char limit.
// Entries are pruned per-event — limits are set below based on 130-player clan sizes.
// ─────────────────────────────────────────────────────────────────────────────

const SHEET_NAME = "AppData";

// Per-event history limits — derived from cell-size analysis at 130 players per clan.
// single-field events: 12 entries ≈ 44K chars (safe)
// two/three-field events: 5 entries ≈ 37K chars (safe)
// Raising these above these values risks silently exceeding the 50K cell limit.
const MAX_HISTORY_BY_EVENT = {
  "weekly_chests": 12,
  "tin_man":       12,
  "ragnarok":      12,
  "omens":          5,
  "olympus":        5,
  "epic_chests":    3,   // auto-synced from ChestTracker; keep last 3 weekly syncs
};

// Fixed mapping: score key → row number in column A
const SCORE_ROW_MAP = {
  "69R_weekly_chests":  3,
  "69R_tin_man":        4,
  "69R_ragnarok":       5,
  // "69R_armageddon":     6,  // retired — data preserved in sheet row 6, not read or written
  "69R_omens":          7,
  "69R_olympus":        8,
  "69S_weekly_chests":  9,
  "69S_tin_man":        10,
  "69S_ragnarok":       11,
  // "69S_armageddon":     12,  // retired — data preserved in sheet row 12, not read or written
  "69S_omens":          13,
  "69S_olympus":        14,
  "69D_weekly_chests":  15,
  "69D_tin_man":        16,
  "69D_ragnarok":       17,
  // "69D_armageddon":     18,  // retired — data preserved in sheet row 18, not read or written
  "69D_omens":          19,
  "69D_olympus":        20,
  "69R_epic_chests":    22,  // auto-synced from ChestTracker API (69R only)
};
const SCORE_KEYS = Object.keys(SCORE_ROW_MAP);

// Default PINs — used when A21 is empty (first-run or restored from old backup).
// These are the source of truth for initial PIN values.
// Once an admin changes a PIN via the app, the new value is stored in A21 and takes over.
const DEFAULT_PINS = {
  super: "9999",
  "69R": "6969",
  "69S": "6996",
  "69D": "9669",
  user:  "1111",
};

const EMPTY_DATA = {
  players: [], scores: {}, levelRequests: [], rotationLog: [],
  fragmentDistributions: [], lastBackup: null, pins: DEFAULT_PINS,
  ctSync: {}, ctAliases: {},
};

// ── Prune a single event's entries to the per-event limit ────────────────────
// eventId: e.g. "omens", "weekly_chests" — used to look up the right limit.
// Handles both the current ARRAY format and the legacy date-keyed OBJECT format.
function pruneEventDates(eventData, eventId) {
  var limit = MAX_HISTORY_BY_EVENT[eventId] || 12;

  if (Array.isArray(eventData)) {
    // Current format: array of { date, scores } objects
    if (eventData.length <= limit) return eventData;
    // Sort newest first (ISO date strings sort correctly as strings), keep most recent
    return eventData.slice().sort(function(a, b) {
      return a.date < b.date ? 1 : a.date > b.date ? -1 : 0;
    }).slice(0, limit);
  }

  // Legacy format: date-keyed object { "2026-01-01": { pid: { points: ... } }, ... }
  if (!eventData || typeof eventData !== "object") return eventData;
  var dates = Object.keys(eventData).sort(); // ascending — oldest first
  if (dates.length <= limit) return eventData;
  var pruned = {};
  dates.slice(dates.length - limit).forEach(function(d) {
    pruned[d] = eventData[d];
  });
  return pruned;
}

// ── Read all data ─────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;

    if (action === "syncEpicChests") {
      return jsonResponse(syncEpicChests());
    }

    // Save PINs via GET so the app can verify the write succeeded (no-cors POST is unverifiable).
    if (action === "savePins") {
      var pinsStr = e.parameter.pins;
      if (!pinsStr) return jsonResponse({ error: "No pins data provided" });
      var pinsData = JSON.parse(pinsStr);
      var sheet = getSheet();
      sheet.getRange("A21").setValue(JSON.stringify({ pins: pinsData }));
      Logger.log("savePins: wrote to A21 — " + JSON.stringify(pinsData));
      return jsonResponse({ ok: true });
    }

    // Save performance norms via GET (same verifiable pattern as savePins).
    if (action === "saveNorms") {
      var normsStr = e.parameter.norms;
      if (!normsStr) return jsonResponse({ error: "No norms data provided" });
      var normsData = JSON.parse(normsStr);
      var sheet = getSheet();
      sheet.getRange("A25").setValue(JSON.stringify(normsData));
      Logger.log("saveNorms: wrote to A25");
      return jsonResponse({ ok: true });
    }

    return jsonResponse(readData());
  }
  catch (err) { return jsonResponse({ error: err.message }); }
}

// ── Write all data ────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const incoming = JSON.parse(e.postData.contents);
    if (incoming.action === "syncEpicChests") {
      return jsonResponse(syncEpicChests());
    }
    // Safety: if the payload has an unrecognised 'action' field and no data fields,
    // reject it rather than passing it to writeData (which would wipe the sheet).
    if (incoming.action) {
      return jsonResponse({ error: "Unknown action: " + incoming.action });
    }
    writeData(incoming);
    return jsonResponse({ ok: true });
  } catch (err) { return jsonResponse({ error: err.message }); }
}

// ── Sheet helper ──────────────────────────────────────────────────────────────
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange("A1").setValue(JSON.stringify({ players: [], lastBackup: null }));
    sheet.getRange("A2").setValue(JSON.stringify({ levelRequests: [], rotationLog: [], fragmentDistributions: [] }));
    SCORE_KEYS.forEach(function(key) {
      var row = SCORE_ROW_MAP[key];
      sheet.getRange("A" + row).setValue(JSON.stringify({ scores: {} }));
    });
    sheet.getRange("A21").setValue(JSON.stringify({ pins: DEFAULT_PINS }));
    sheet.getRange("A24").setValue(JSON.stringify({ ctSync: {}, ctAliases: {} }));
    sheet.getRange("A25").setValue(JSON.stringify({}));
  }
  return sheet;
}

// ── Read ──────────────────────────────────────────────────────────────────────
function readData() {
  var sheet = getSheet();
  var rawA1 = sheet.getRange("A1").getValue();
  var rawA2 = sheet.getRange("A2").getValue();

  if (!rawA1) return EMPTY_DATA;

  var parse = function(raw) {
    try { return raw ? JSON.parse(raw) : {}; } catch(_) { return {}; }
  };

  var dataA1 = parse(rawA1);

  // ── Detect legacy format (old row-1 layout: A1/B1/C1...) ──────────────────
  var rawB1 = sheet.getRange("B1").getValue();
  if (rawB1 && rawB1.length > 20) {
    return migrateLegacy(sheet);
  }

  var dataA2 = parse(rawA2);

  // ── Read players — may be split across A1 + A23 if >50K chars ────────────
  var players = dataA1.players || [];
  var rawA23  = sheet.getRange("A23").getValue();
  if (rawA23) {
    var dataA23 = parse(rawA23);
    if (dataA23.players) players = players.concat(dataA23.players);
  }

  // Read all score rows
  var scores = {};
  SCORE_KEYS.forEach(function(key) {
    var row  = SCORE_ROW_MAP[key];
    var raw  = sheet.getRange("A" + row).getValue();
    var d    = parse(raw);
    if (d.scores && d.scores[key]) {
      scores[key] = d.scores[key];
    }
  });

  var rawA21 = sheet.getRange("A21").getValue();
  var dataA21 = parse(rawA21);
  var rawA24 = sheet.getRange("A24").getValue();
  var dataA24 = parse(rawA24);
  var rawA25 = sheet.getRange("A25").getValue();
  var normsData = rawA25 ? parse(rawA25) : {};

  return {
    players:               players,
    lastBackup:            dataA1.lastBackup             || null,
    scores:                scores,
    levelRequests:         dataA2.levelRequests          || [],
    rotationLog:           dataA2.rotationLog            || [],
    fragmentDistributions: dataA2.fragmentDistributions  || [],
    pins:                  dataA21.pins                  || DEFAULT_PINS,
    ctSync:                dataA24.ctSync                || {},
    ctAliases:             dataA24.ctAliases             || {},
    norms:                 normsData,
  };
}

// ── Write ─────────────────────────────────────────────────────────────────────
function writeData(data) {
  // Safety guard: if the object looks like an action request (has .action but no
  // recognised data fields), refuse to write rather than silently wipe the sheet.
  if (!data || (data.action && !data.players && !data.scores && !data.levelRequests && !data.pins && !data.ctSync)) {
    throw new Error("writeData: invalid payload — missing required data fields. Keys: " + JSON.stringify(Object.keys(data || {})));
  }

  var sheet = getSheet();

  // ── PIN-only save: just update A21, leave everything else untouched ──────────
  if (data.pins && !data.players && !data.scores) {
    sheet.getRange("A21").setValue(JSON.stringify({ pins: data.pins }));
    return;
  }

  var scores = data.scores || {};

  // ── Split players across A1 + A23 if needed to stay under 50K char limit ──
  var allPlayers   = data.players || [];
  var mid          = Math.ceil(allPlayers.length / 2);
  var playersA1    = allPlayers.slice(0, mid);
  var playersA23   = allPlayers.slice(mid);
  var a1Json       = JSON.stringify({ players: playersA1, lastBackup: data.lastBackup || null });
  // If first half alone fits, use it; otherwise split evenly
  if (a1Json.length > 49000) {
    // Re-split more aggressively
    mid       = Math.floor(allPlayers.length / 3);
    playersA1 = allPlayers.slice(0, mid);
    playersA23 = allPlayers.slice(mid);
    a1Json    = JSON.stringify({ players: playersA1, lastBackup: data.lastBackup || null });
  }
  sheet.getRange("A1").setValue(a1Json);
  sheet.getRange("A23").setValue(playersA23.length > 0 ? JSON.stringify({ players: playersA23 }) : "");
  sheet.getRange("A2").setValue(JSON.stringify({
    levelRequests:         data.levelRequests         || [],
    rotationLog:           data.rotationLog           || [],
    fragmentDistributions: data.fragmentDistributions || [],
  }));

  // Write PINs to A21 (only if provided — null means no change)
  if (data.pins) {
    sheet.getRange("A21").setValue(JSON.stringify({ pins: data.pins }));
  }

  // Write ChestTracker sync metadata to A24
  var existingA24Raw = sheet.getRange("A24").getValue();
  var existingA24 = (function(raw) { try { return raw ? JSON.parse(raw) : {}; } catch(_) { return {}; } })(existingA24Raw);
  sheet.getRange("A24").setValue(JSON.stringify({
    ctSync:    data.ctSync    !== undefined ? data.ctSync    : (existingA24.ctSync    || {}),
    ctAliases: data.ctAliases !== undefined ? data.ctAliases : (existingA24.ctAliases || {}),
  }));

  // Write each score key to its own cell, pruning old entries
  SCORE_KEYS.forEach(function(key) {
    var row     = SCORE_ROW_MAP[key];
    var eventId = key.split("_").slice(1).join("_"); // "69R_weekly_chests" → "weekly_chests"
    var eventData = pruneEventDates(scores[key] || [], eventId);
    sheet.getRange("A" + row).setValue(JSON.stringify({ scores: { [key]: eventData } }));
  });
}

// ── Migrate from legacy row-1 layout (B1/C1/D1/E1/F1/G1/H1) ─────────────────
function migrateLegacy(sheet) {
  Logger.log("Detected legacy layout — migrating to column-A format...");

  var parse = function(raw) {
    try { return raw ? JSON.parse(raw) : {}; } catch(_) { return {}; }
  };

  // Gather all scores from old cells (B1, D1, E1, F1, G1, H1)
  var scores = {};
  ["B1","D1","E1","F1","G1","H1"].forEach(function(ref) {
    var raw = sheet.getRange(ref).getValue();
    var d   = parse(raw);
    Object.keys(d.scores || {}).forEach(function(k) { scores[k] = d.scores[k]; });
  });

  var rawA1 = parse(sheet.getRange("A1").getValue());
  var rawC1 = parse(sheet.getRange("C1").getValue());

  // Write into new layout
  sheet.getRange("A1").setValue(JSON.stringify({
    players: rawA1.players || [], lastBackup: rawA1.lastBackup || null
  }));
  sheet.getRange("A2").setValue(JSON.stringify({
    levelRequests:         rawC1.levelRequests         || [],
    rotationLog:           rawC1.rotationLog           || [],
    fragmentDistributions: rawC1.fragmentDistributions || [],
  }));
  SCORE_KEYS.forEach(function(key) {
    var row     = SCORE_ROW_MAP[key];
    var eventId = key.split("_").slice(1).join("_");
    var eventData = pruneEventDates(scores[key] || [], eventId);
    sheet.getRange("A" + row).setValue(JSON.stringify({ scores: { [key]: eventData } }));
  });

  // Clear old cells
  ["B1","C1","D1","E1","F1","G1","H1"].forEach(function(ref) {
    sheet.getRange(ref).clearContent();
  });

  Logger.log("Migration complete.");
  return readData();
}

// ── Run once manually after deploying to migrate existing data ─────────────────
// Select "redistributeScores" from the function dropdown and click Run.
function redistributeScores() {
  var sheet = getSheet();
  var rawB1 = sheet.getRange("B1").getValue();

  if (!rawB1 || rawB1.length < 20) {
    Logger.log("No legacy data found in B1 — nothing to migrate.");
    return;
  }

  migrateLegacy(sheet);

  Logger.log("=== Results after migration ===");
  SCORE_KEYS.forEach(function(key) {
    var row = SCORE_ROW_MAP[key];
    var raw = sheet.getRange("A" + row).getValue();
    Logger.log("A" + row + " (" + key + "): " + (raw ? raw.length : 0) + " chars");
  });
}

// ── Diagnostic ────────────────────────────────────────────────────────────────
function diagnose() {
  var sheet = getSheet();

  var r1 = sheet.getRange("A1").getValue();
  var r2 = sheet.getRange("A2").getValue();
  Logger.log("=== A1 — players (" + (r1?r1.length:0) + " chars) ===");
  try { var d = JSON.parse(r1); Logger.log("players: " + (d.players||[]).length); } catch(_){}
  Logger.log("=== A2 — requests/log (" + (r2?r2.length:0) + " chars) ===");
  try { var d = JSON.parse(r2); Logger.log("levelRequests: " + (d.levelRequests||[]).length); } catch(_){}

  SCORE_KEYS.forEach(function(key) {
    var row = SCORE_ROW_MAP[key];
    var raw = sheet.getRange("A" + row).getValue();
    var len = raw ? raw.length : 0;
    var pct = Math.round(len / 500);
    Logger.log("A" + row + " [" + key + "]: " + len + " chars (" + pct + "% of limit)" + (len > 40000 ? " ⚠️ NEARLY FULL" : ""));
    try {
      var d = JSON.parse(raw);
      var entries = d.scores && d.scores[key] ? Object.keys(d.scores[key]).length : 0;
      Logger.log("  → " + entries + " date entries");
    } catch(_){}
  });
}

function jsonResponse(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ─── CHESTTRACKER API SYNC ────────────────────────────────────────────────────
// To configure: go to Apps Script → Project Settings → Script Properties and add:
//   CT_EMAIL    → your ChestTracker login email
//   CT_PASSWORD → your ChestTracker login password
//
// To set up the Sunday 10pm trigger: run setupEpicChestsTrigger() once from
// the Apps Script editor (select it from the function dropdown and click Run).
// ─────────────────────────────────────────────────────────────────────────────

var CT_API_BASE = "https://api.chesttracker.com/v1";

function getCTToken() {
  var props = PropertiesService.getScriptProperties();
  var email    = props.getProperty("CT_EMAIL");
  var password = props.getProperty("CT_PASSWORD");
  if (!email || !password) {
    throw new Error("ChestTracker credentials not configured. Add CT_EMAIL and CT_PASSWORD in Script Properties.");
  }

  var resp = UrlFetchApp.fetch(CT_API_BASE + "/authenticate", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ email: email, password: password }),
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

// Main sync function — called by hourly time trigger or manually.
// Queries ChestTracker from the start of the current week (Sunday midnight UTC)
// to now, so each run updates the same weekly entry rather than creating a new one.
// On the first run of a new week the old entry is archived and a fresh one starts.
function syncEpicChests() {
  try {
    var token = getCTToken();

    var now = new Date();

    // Week boundary: Sunday 18:00 UTC → following Sunday 17:59 UTC
    // If it's Sunday but before 18:00, we're still in last week.
    var dayOfWeek = now.getUTCDay(); // 0 = Sunday
    var daysBack  = dayOfWeek === 0 && now.getUTCHours() < 18 ? 7 : dayOfWeek;
    var startOfWeek = new Date(now);
    startOfWeek.setUTCDate(now.getUTCDate() - daysBack);
    startOfWeek.setUTCHours(18, 0, 0, 0);
    var weekStart = startOfWeek.toISOString().slice(0, 10); // "2026-08-23" (the Sunday)

    var url = CT_API_BASE + "/chests/breakdown?levels=1"
            + "&start=" + encodeURIComponent(startOfWeek.toISOString())
            + "&end="   + encodeURIComponent(now.toISOString());

    var resp = UrlFetchApp.fetch(url, {
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error("ChestTracker breakdown failed (" + resp.getResponseCode() + ")");
    }

    // Response is [ [memberArray], schemaMeta ] or just [memberArray]
    var parsed  = JSON.parse(resp.getContentText());
    var members = Array.isArray(parsed[0]) ? parsed[0] : parsed;

    // Load current app data (includes ctAliases for name mapping)
    var appData  = readData();
    var players  = (appData.players || []).filter(function(p) { return p.clan === "69R" && p.active; });
    var aliases  = (appData.ctAliases || {})["69R"] || {};  // { "CT Name": "playerId" }

    // Build a lookup: normalised name → playerId, incorporating saved aliases
    var nameLookup = {};
    players.forEach(function(p) {
      nameLookup[p.name.toLowerCase().trim()] = p.id;
    });

    var matched   = {};  // playerId → epics count
    var unmatched = [];  // { name, epics }

    members.forEach(function(member) {
      var ctName = (member.name || "").trim();
      if (!ctName) return;

      // Use only the "epic squad" chest type
      var epicSquad = member["epic squad"];
      var epics = (epicSquad && typeof epicSquad.chests === "number") ? epicSquad.chests : 0;

      // Match: saved alias → name lookup → unmatched
      var playerId = aliases[ctName] || nameLookup[ctName.toLowerCase()];
      if (playerId) {
        matched[playerId] = epics;
      } else {
        unmatched.push({ name: ctName, epics: epics });
      }
    });

    // Build score entry — date is set to week start so player profiles display
    // the week's Sunday date rather than the exact sync timestamp.
    var scores = {};
    Object.keys(matched).forEach(function(pid) {
      scores[pid] = { score: matched[pid] };
    });

    var entry = {
      date:      weekStart + "T00:00:00.000Z",  // week start date for display
      weekStart: weekStart,
      syncedAt:  now.toISOString(),
      scores:    scores,
    };

    // Update or create this week's entry
    if (!appData.scores) appData.scores = {};
    var key = "69R_epic_chests";
    if (!appData.scores[key]) appData.scores[key] = [];

    var entries = appData.scores[key];
    if (entries.length > 0 && entries[0].weekStart === weekStart) {
      // Same week — overwrite with refreshed scores
      entries[0] = entry;
    } else {
      // New week — archive the old entry and start a fresh one
      entries.unshift(entry);
      if (entries.length > 3) entries = entries.slice(0, 3);
    }
    appData.scores[key] = entries;

    // Update ctSync metadata
    if (!appData.ctSync) appData.ctSync = {};
    appData.ctSync["69R"] = {
      lastSync:  now.toISOString(),
      matched:   Object.keys(matched).length,
      unmatched: unmatched,
    };

    writeData(appData);
    Logger.log("syncEpicChests: week=" + weekStart + " matched=" + Object.keys(matched).length + " unmatched=" + unmatched.length);
    return { success: true, matched: Object.keys(matched).length, unmatched: unmatched.length };

  } catch (err) {
    Logger.log("syncEpicChests ERROR: " + err.message);
    return { success: false, error: err.message };
  }
}

// Run this ONCE from the Apps Script editor to register the hourly trigger.
// If you need to change the frequency, run it again — it deletes the old trigger first.
// To use every 30 minutes instead, change everyHours(1) to everyMinutes(30).
function setupEpicChestsTrigger() {
  // Remove any existing epic chest triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "syncEpicChests") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create hourly trigger
  ScriptApp.newTrigger("syncEpicChests")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("Epic Chests trigger set: every hour.");
}
