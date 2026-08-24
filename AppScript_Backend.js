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
//   A24 → { ctSync: { 69R: {...}, 69S: {...} }, ctAliases: { 69R: {...}, 69S: {...} } }
//   A26 → { scores: { 69S_epic_chests } }
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
  "69R_epic_chests":    22,  // auto-synced from ChestTracker API
  "69S_epic_chests":    26,  // auto-synced from ChestTracker API (same account)
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
  ctSync: {}, ctAliases: {}, ctIgnored: {},
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

    // Save a CT name → player ID alias for a given clan.
    // Also removes the name from ctIgnored if it was there.
    if (action === "saveAlias") {
      var clan    = e.parameter.clan;
      var ctName  = e.parameter.ctName;
      var pid     = e.parameter.playerId;
      if (!clan || !ctName || !pid) return jsonResponse({ error: "Missing clan, ctName, or playerId" });
      var sheet   = getSheet();
      var raw24   = sheet.getRange("A24").getValue();
      var a24     = raw24 ? JSON.parse(raw24) : {};
      if (!a24.ctAliases)      a24.ctAliases = {};
      if (!a24.ctAliases[clan]) a24.ctAliases[clan] = {};
      a24.ctAliases[clan][ctName] = pid;
      // Remove from ignored if present
      if (a24.ctIgnored && a24.ctIgnored[clan]) {
        a24.ctIgnored[clan] = (a24.ctIgnored[clan] || []).filter(function(n) { return n !== ctName; });
      }
      sheet.getRange("A24").setValue(JSON.stringify(a24));
      Logger.log("saveAlias: " + clan + " [" + ctName + "] → " + pid);
      return jsonResponse({ ok: true });
    }

    // Remove a saved CT alias for a clan (so it goes back to unmatched on next sync).
    if (action === "removeAlias") {
      var clan   = e.parameter.clan;
      var ctName = e.parameter.ctName;
      if (!clan || !ctName) return jsonResponse({ error: "Missing clan or ctName" });
      var sheet  = getSheet();
      var raw24  = sheet.getRange("A24").getValue();
      var a24    = raw24 ? JSON.parse(raw24) : {};
      if (a24.ctAliases && a24.ctAliases[clan]) {
        delete a24.ctAliases[clan][ctName];
      }
      sheet.getRange("A24").setValue(JSON.stringify(a24));
      Logger.log("removeAlias: " + clan + " [" + ctName + "]");
      return jsonResponse({ ok: true });
    }

    // Add or remove a CT name from the ignore list for a given clan.
    // Pass remove=1 to un-ignore a name.
    if (action === "saveIgnored") {
      var clan   = e.parameter.clan;
      var ctName = e.parameter.ctName;
      var remove = e.parameter.remove === "1";
      if (!clan || !ctName) return jsonResponse({ error: "Missing clan or ctName" });
      var sheet  = getSheet();
      var raw24  = sheet.getRange("A24").getValue();
      var a24    = raw24 ? JSON.parse(raw24) : {};
      if (!a24.ctIgnored)       a24.ctIgnored = {};
      if (!a24.ctIgnored[clan]) a24.ctIgnored[clan] = [];
      if (remove) {
        a24.ctIgnored[clan] = a24.ctIgnored[clan].filter(function(n) { return n !== ctName; });
        Logger.log("saveIgnored: un-ignored " + clan + " [" + ctName + "]");
      } else {
        if (a24.ctIgnored[clan].indexOf(ctName) === -1) a24.ctIgnored[clan].push(ctName);
        // Remove any alias for this name when ignoring
        if (a24.ctAliases && a24.ctAliases[clan]) {
          delete a24.ctAliases[clan][ctName];
        }
        Logger.log("saveIgnored: ignored " + clan + " [" + ctName + "]");
      }
      sheet.getRange("A24").setValue(JSON.stringify(a24));
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
    sheet.getRange("A24").setValue(JSON.stringify({ ctSync: {}, ctAliases: {}, ctIgnored: {} }));
    sheet.getRange("A25").setValue(JSON.stringify({}));
    sheet.getRange("A26").setValue(JSON.stringify({ scores: {} }));
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
    ctIgnored:             dataA24.ctIgnored             || {},
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
    ctIgnored: data.ctIgnored !== undefined ? data.ctIgnored : (existingA24.ctIgnored || {}),
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

// Authenticate and return a token.
// Pass memberId to get a clan-scoped token (required for /chests/breakdown).
// Without memberId, the token is only valid for listing endpoints (/clans, /members).
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

// Main sync function — called by hourly time trigger or manually.
// Queries ChestTracker from the start of the current week (Sunday midnight UTC)
// to now, so each run updates the same weekly entry rather than creating a new one.
// On the first run of a new week the old entry is archived and a fresh one starts.
// Syncs both 69R and 69S from the same CT account in a single API call.
function syncEpicChests() {
  try {
    var now = new Date();

    // Week boundary: Sunday 18:00 UTC → following Sunday 17:59 UTC
    var dayOfWeek = now.getUTCDay(); // 0 = Sunday
    var daysBack  = dayOfWeek === 0 && now.getUTCHours() < 18 ? 7 : dayOfWeek;
    var startOfWeek = new Date(now);
    startOfWeek.setUTCDate(now.getUTCDate() - daysBack);
    startOfWeek.setUTCHours(18, 0, 0, 0);
    var weekStart = startOfWeek.toISOString().slice(0, 10);

    var timeParams = "?levels=1"
                   + "&start=" + encodeURIComponent(startOfWeek.toISOString())
                   + "&end="   + encodeURIComponent(now.toISOString());

    // Step 1: Base auth (no memberId) — only valid for listing endpoints.
    var baseToken   = getCTToken();
    var baseHeaders = { "Authorization": "Bearer " + baseToken };

    // Step 2: Get clan tag→id map from /clans.
    var clansRaw  = JSON.parse(UrlFetchApp.fetch(CT_API_BASE + "/clans", { headers: baseHeaders, muteHttpExceptions: true }).getContentText());
    var clansList = Array.isArray(clansRaw[0]) ? clansRaw[0] : clansRaw;
    var tagToClanId = {};
    clansList.forEach(function(c) { if (c.tag && c.id) tagToClanId[c.tag] = c.id; });

    // Step 3: Get the account's per-clan member IDs from /members.
    // /members returns the CT account's own membership records (one per clan).
    var membersRaw   = JSON.parse(UrlFetchApp.fetch(CT_API_BASE + "/members", { headers: baseHeaders, muteHttpExceptions: true }).getContentText());
    var acctMembers  = Array.isArray(membersRaw[0]) ? membersRaw[0] : membersRaw;
    // Build clanId → memberId (the account's membership ID in that clan)
    var clanIdToMemberId = {};
    acctMembers.forEach(function(m) { if (m.clanId && m.id) clanIdToMemberId[m.clanId] = m.id; });
    // Derive tag → memberId
    var tagToMemberId = {};
    Object.keys(tagToClanId).forEach(function(tag) {
      var clanId = tagToClanId[tag];
      if (clanIdToMemberId[clanId]) tagToMemberId[tag] = clanIdToMemberId[clanId];
    });
    Logger.log("syncEpicChests: clans=" + JSON.stringify(tagToClanId) + " memberIds=" + JSON.stringify(tagToMemberId));

    // Load current app data
    var appData = readData();
    if (!appData.scores) appData.scores = {};
    if (!appData.ctSync) appData.ctSync = {};

    var clansToSync = ["69R", "69S"];
    var results = {};

    clansToSync.forEach(function(clan) {
      var memberId = tagToMemberId[clan];
      if (!memberId) {
        Logger.log("syncEpicChests: no CT memberId for " + clan + " — skipping");
        results[clan] = { matched: 0, unmatched: 0 };
        return;
      }

      // Step 4: Per-clan auth — memberId scopes the token to this clan.
      var clanToken   = getCTToken(memberId);
      var clanHeaders = { "Authorization": "Bearer " + clanToken };

      // Step 5: Fetch chest breakdown for this clan.
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

      var players  = (appData.players || []).filter(function(p) { return p.clan === clan && p.active; });
      var aliases  = (appData.ctAliases || {})[clan] || {};   // { "CT Name": "playerId" }
      var ignored  = (appData.ctIgnored || {})[clan] || [];   // ["CT Name", ...]

      // Build name → playerId lookup
      var nameLookup = {};
      players.forEach(function(p) {
        nameLookup[p.name.toLowerCase().trim()] = p.id;
      });

      var matched   = {};  // playerId → epics count
      var unmatched = [];  // { name, epics } — CT names not found in this clan

      members.forEach(function(member) {
        var ctName = (member.name || "").trim();
        if (!ctName) return;
        if (ignored.indexOf(ctName) !== -1) return;

        // Use only the "epic squad" chest type
        var epicSquad = member["epic squad"];
        var epics = (epicSquad && typeof epicSquad.chests === "number") ? epicSquad.chests : 0;

        // Match: saved alias → name lookup → unmatched
        var playerId = aliases[ctName] || nameLookup[ctName.toLowerCase().trim()];
        if (playerId) {
          matched[playerId] = epics;
        } else {
          unmatched.push({ name: ctName, epics: epics });
        }
      });

      // Build score entry — date is week start so player profiles show the Sunday date
      var scores = {};
      Object.keys(matched).forEach(function(pid) {
        scores[pid] = { score: matched[pid] };
      });

      var entry = {
        date:      weekStart + "T00:00:00.000Z",
        weekStart: weekStart,
        syncedAt:  now.toISOString(),
        scores:    scores,
      };

      // Update or create this week's entry
      var key = clan + "_epic_chests";
      if (!appData.scores[key]) appData.scores[key] = [];

      var entries = appData.scores[key];
      if (entries.length > 0 && entries[0].weekStart === weekStart) {
        entries[0] = entry;  // Same week — overwrite with refreshed scores
      } else {
        entries.unshift(entry);  // New week — archive old and start fresh
        if (entries.length > 3) entries = entries.slice(0, 3);
      }
      appData.scores[key] = entries;

      // Update ctSync metadata for this clan
      appData.ctSync[clan] = {
        lastSync:  now.toISOString(),
        matched:   Object.keys(matched).length,
        unmatched: unmatched,
      };

      results[clan] = { matched: Object.keys(matched).length, unmatched: unmatched.length };
      Logger.log("syncEpicChests " + clan + ": week=" + weekStart + " matched=" + Object.keys(matched).length + " unmatched=" + unmatched.length);
    });

    writeData(appData);
    return { success: true, "69R": results["69R"], "69S": results["69S"] };

  } catch (err) {
    Logger.log("syncEpicChests ERROR: " + err.message);
    return { success: false, error: err.message };
  }
}

// ── DIAGNOSTIC — run this ONCE from the Apps Script editor ───────────────────
// Probes the CT API to find the right endpoint/parameter for selecting a clan.
function debugCtResponse() {
  try {
    var props    = PropertiesService.getScriptProperties();
    var email    = props.getProperty("CT_EMAIL");
    var password = props.getProperty("CT_PASSWORD");

    // Step 1: Log full auth response — might contain group/clan IDs
    Logger.log("=== AUTH RESPONSE ===");
    var authResp = UrlFetchApp.fetch(CT_API_BASE + "/authenticate", {
      method: "post",
      headers: { "Content-Type": "application/json" },
      payload: JSON.stringify({ email: email, password: password }),
      muteHttpExceptions: true,
    });
    Logger.log("Auth status: " + authResp.getResponseCode());
    Logger.log("Auth body: " + authResp.getContentText());

    var authData = JSON.parse(authResp.getContentText());
    var token    = authData.authToken || authData.token || authData.access_token || authData.accessToken || "";
    if (!token) { Logger.log("No token found — check auth body above"); return; }
    Logger.log("Token obtained: " + token.slice(0, 30) + "...");

    var headers = { "Authorization": "Bearer " + token };

    // Step 2: Try /groups or /clans or /me to find group IDs
    Logger.log("=== PROBING ENDPOINTS ===");
    var endpoints = ["/groups", "/clans", "/me", "/user", "/user/groups", "/account"];
    endpoints.forEach(function(ep) {
      try {
        var r = UrlFetchApp.fetch(CT_API_BASE + ep, { headers: headers, muteHttpExceptions: true });
        Logger.log(ep + " → " + r.getResponseCode() + ": " + r.getContentText().slice(0, 300));
      } catch(e) { Logger.log(ep + " → ERROR: " + e.message); }
    });

    // Step 3: Get clan IDs from /clans
    Logger.log("=== CLAN IDs ===");
    var clansRaw  = JSON.parse(UrlFetchApp.fetch(CT_API_BASE + "/clans", { headers: headers, muteHttpExceptions: true }).getContentText());
    var clansList = Array.isArray(clansRaw[0]) ? clansRaw[0] : clansRaw;
    var tagToId = {};
    clansList.forEach(function(c) { if (c.tag && c.id) tagToId[c.tag] = c.id; });
    Logger.log("Clans: " + JSON.stringify(tagToId));
    var id69S = tagToId["69S"] || "82ee1599-67c3-41c2-868f-e751541264e1";
    var id69R = tagToId["69R"] || "97abb703-1122-43b4-a77c-2f5b21464c6d";

    // Step 4: Try breakdown endpoint path variations
    Logger.log("=== TRYING ENDPOINT PATH VARIATIONS ===");
    var now   = new Date();
    var start = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    var timeQ = "levels=1&start=" + encodeURIComponent(start.toISOString()) + "&end=" + encodeURIComponent(now.toISOString());
    var pathTests = [
      "/clans/" + id69S + "/chests/breakdown?" + timeQ,
      "/clans/" + id69S + "/breakdown?" + timeQ,
      "/chests/breakdown?clanId=" + id69S + "&" + timeQ,
      "/chests/breakdown?groupId=" + id69S + "&" + timeQ,
      "/chests/breakdown?clan=" + id69S + "&" + timeQ,
    ];
    pathTests.forEach(function(path) {
      try {
        var r = UrlFetchApp.fetch(CT_API_BASE + path, { headers: headers, muteHttpExceptions: true });
        var code = r.getResponseCode();
        if (code === 200) {
          var parsed = JSON.parse(r.getContentText());
          var members = Array.isArray(parsed[0]) ? parsed[0] : parsed;
          Logger.log(path.slice(0,60) + " → " + code + " members: " + members.length + " first: " + (members[0] && members[0].name));
        } else {
          Logger.log(path.slice(0,60) + " → " + code);
        }
      } catch(e) { Logger.log(path.slice(0,60) + " → ERROR: " + e.message); }
    });

    // Step 5: Get actual member IDs from /members (CT account memberships, not game players)
    Logger.log("=== /members — GET ACCOUNT MEMBER IDs ===");
    var membersResp = UrlFetchApp.fetch(CT_API_BASE + "/members", { headers: headers, muteHttpExceptions: true });
    var membersRaw  = JSON.parse(membersResp.getContentText());
    var acctMembers = Array.isArray(membersRaw[0]) ? membersRaw[0] : membersRaw;
    Logger.log("Account members: " + JSON.stringify(acctMembers.map(function(m) { return { id: m.id, name: m.name, clanId: m.clanId }; })));

    var memberId69R = null, memberId69S = null;
    acctMembers.forEach(function(m) {
      if (m.clanId === id69R) memberId69R = m.id;
      if (m.clanId === id69S) memberId69S = m.id;
    });
    Logger.log("69R member ID: " + memberId69R);
    Logger.log("69S member ID: " + memberId69S);

    // Step 6: Auth with specific MEMBER IDs to get clan-scoped tokens
    Logger.log("=== AUTH WITH MEMBER IDs ===");
    function decodeJwtPayload(jwt) {
      try {
        var parts = jwt.split(".");
        var payload = parts[1].replace(/-/g, "+").replace(/_/g, "/");
        while (payload.length % 4 !== 0) payload += "=";
        return Utilities.newBlob(Utilities.base64Decode(payload)).getDataAsString();
      } catch(e) { return "decode error: " + e.message; }
    }

    function tryAuth(label, body) {
      var r = UrlFetchApp.fetch(CT_API_BASE + "/authenticate", {
        method: "post",
        headers: { "Content-Type": "application/json" },
        payload: JSON.stringify(body),
        muteHttpExceptions: true,
      });
      var d = JSON.parse(r.getContentText());
      var tk = d.authToken || d.token || "";
      if (!tk) { Logger.log(label + ": no token — " + r.getContentText().slice(0, 100)); return null; }
      var decoded = JSON.parse(decodeJwtPayload(tk));
      Logger.log(label + ": authorizedMember=" + JSON.stringify(decoded.authorizedMember) + " favoriteId=" + decoded.favoriteId);
      // Test breakdown with this token
      var br = UrlFetchApp.fetch(CT_API_BASE + "/chests/breakdown?" + timeQ, {
        headers: { "Authorization": "Bearer " + tk },
        muteHttpExceptions: true,
      });
      if (br.getResponseCode() === 200) {
        var bp = JSON.parse(br.getContentText());
        var bm = Array.isArray(bp[0]) ? bp[0] : bp;
        Logger.log(label + " breakdown → " + br.getResponseCode() + " members=" + (Array.isArray(bm) ? bm.length : "?") + " first=" + (bm[0] && bm[0].name));
      } else {
        Logger.log(label + " breakdown → " + br.getResponseCode() + ": " + br.getContentText().slice(0,100));
      }
      return tk;
    }

    // Base auth (no extras) — confirm current state
    tryAuth("base auth", { email: email, password: password });
    // Auth with 69R member ID
    if (memberId69R) tryAuth("memberId=69R(" + memberId69R.slice(0,8) + "...)", { email: email, password: password, memberId: memberId69R });
    // Auth with 69S member ID
    if (memberId69S) tryAuth("memberId=69S(" + memberId69S.slice(0,8) + "...)", { email: email, password: password, memberId: memberId69S });

  } catch(err) {
    Logger.log("debugCtResponse ERROR: " + err.message);
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
