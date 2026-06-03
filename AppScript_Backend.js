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
//
// Each cell stays well under Google's 50,000 char limit.
// Old entries are automatically pruned — MAX_HISTORY_ENTRIES kept per event.
// ─────────────────────────────────────────────────────────────────────────────

const SHEET_NAME       = "AppData";
const MAX_HISTORY_ENTRIES = 20; // How many date entries to keep per event

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
};
const SCORE_KEYS = Object.keys(SCORE_ROW_MAP);

const EMPTY_DATA = {
  players: [], scores: {}, levelRequests: [], rotationLog: [],
  fragmentDistributions: [], lastBackup: null, pins: null,
};

// ── Prune a single event's date entries to MAX_HISTORY_ENTRIES ─────────────────
function pruneEventDates(eventData) {
  if (!eventData || typeof eventData !== "object") return eventData;
  var dates = Object.keys(eventData).sort(); // ascending — oldest first
  if (dates.length <= MAX_HISTORY_ENTRIES) return eventData;
  var pruned = {};
  dates.slice(dates.length - MAX_HISTORY_ENTRIES).forEach(function(d) {
    pruned[d] = eventData[d];
  });
  return pruned;
}

// ── Read all data ─────────────────────────────────────────────────────────────
function doGet(e) {
  try { return jsonResponse(readData()); }
  catch (err) { return jsonResponse({ error: err.message }); }
}

// ── Write all data ────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const incoming = JSON.parse(e.postData.contents);
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

  return {
    players:               players,
    lastBackup:            dataA1.lastBackup             || null,
    scores:                scores,
    levelRequests:         dataA2.levelRequests          || [],
    rotationLog:           dataA2.rotationLog            || [],
    fragmentDistributions: dataA2.fragmentDistributions  || [],
    pins:                  dataA21.pins                  || null,
  };
}

// ── Write ─────────────────────────────────────────────────────────────────────
function writeData(data) {
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

  // Write each score key to its own cell, pruning old entries
  SCORE_KEYS.forEach(function(key) {
    var row       = SCORE_ROW_MAP[key];
    var eventData = pruneEventDates(scores[key] || {});
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
    var row       = SCORE_ROW_MAP[key];
    var eventData = pruneEventDates(scores[key] || {});
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
