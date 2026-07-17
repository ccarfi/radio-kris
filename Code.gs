/**
 * Radio KRIS — Apps Script backend
 * Bound to the Google Sheet. Deploy as Web App:
 *   Execute as: Me   |   Who has access: Anyone
 *
 * Setup:
 *   1. Project Settings → Script Properties → add  YT_API_KEY = <your YouTube Data API v3 key>
 *   2. (Optional) change AUTH_HASH/AUTH_SALT below if you rotate the password.
 *      Current password hash corresponds to the salted word agreed in the spec.
 *   3. Deploy → New deployment → Web app. Copy the /exec URL into index.html CONFIG.
 *
 * The frontend sends the SALTED HASH (never the plaintext word) as `auth` on every call.
 */

// ---- Config -----------------------------------------------------------------
var AUTH_SALT = 'ef6cc100b3416ca73a69323d1e3b3ef1';
var AUTH_HASH = '7d522062184616b7716bc3c8fca7a593e8692032ce6206377d0326ac7fc34555'; // SHA-256(salt + password)

var META_SHEET = 'Metadata';
var PRESENCE_SHEET = 'Presence';
var GONG_COOLDOWN_MS = 5000;       // server-enforced Gong cooldown
var PRESENCE_TTL_MS = 30000;       // a listener is "tuning in" if seen within this window
var LOCK_WAIT_MS = 10000;

var STATION_HEADERS = ['id', 'videoId', 'title', 'artist', 'durationSec', 'thumbnailUrl', 'addedBy', 'addedAt'];
var META_HEADERS = ['station', 'currentTrackId', 'trackStartedAt', 'lastGongBy', 'lastGongAt', 'lastActionAt'];
var PRESENCE_HEADERS = ['station', 'handle', 'lastSeenAt'];

// ---- Cohorts (password -> default landing playlist) -------------------------
// Any valid cohort hash unlocks the whole app; the hash only picks the landing
// station. All hashes use the single global AUTH_SALT.
var COHORT_SHEET = 'Cohorts';
var COHORT_HEADERS = ['label', 'passwordHash', 'defaultStation'];
var SHEEPLE_HASH = '64f7004c837f7232900520bad14446f00dfe2874e90fc4e4929504da9536e555'; // SHA-256(salt + "baa-ram-ewe")
// Seeded into the Cohorts tab if it's empty, so the app can never brick itself.
var COHORT_SEED = [
  ['Extricity', AUTH_HASH, 'Extricity Classics'],
  ['We The Sheeple', SHEEPLE_HASH, 'We The Sheeple']
];

// ---- Entry points -----------------------------------------------------------
function doGet(e) { return handle(e && e.parameter ? e.parameter : {}); }

function doPost(e) {
  var data = {};
  try {
    if (e && e.postData && e.postData.contents) data = JSON.parse(e.postData.contents);
  } catch (err) { data = (e && e.parameter) || {}; }
  return handle(data);
}

function handle(data) {
  try {
    if (!data || !isValidAuth(data.auth)) return json({ ok: false, error: 'unauthorized' });
    var action = data.action;
    var out;
    switch (action) {
      case 'login':         out = doLogin(data.auth); break;
      case 'getState':      out = getState(data.station); break;
      case 'getStations':   out = { stations: getStations() }; break;
      case 'createPlaylist':out = createPlaylist(data.station); break;
      case 'search':        out = { results: searchYouTube(data.query) }; break;
      case 'addTrack':      out = addTrack(data); break;
      case 'removeTrack':   out = removeTrack(data.station, data.id, data.by); break;
      case 'playTrack':     out = playTrack(data.station, data.id, data.by); break;
      case 'gong':          out = gong(data.station, data.by); break;
      case 'advance':       out = advance(data.station, data.expectedTrackId, data.force); break;
      case 'heartbeat':     out = heartbeat(data.station, data.handle); break;
      default:              return json({ ok: false, error: 'unknown action: ' + action });
    }
    out = out || {};
    out.ok = (out.ok !== false);
    out.serverNow = Date.now();
    return json(out);
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err), serverNow: Date.now() });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Sheet helpers ----------------------------------------------------------
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function ensureSheet(name, headers) {
  var s = ss().getSheetByName(name);
  if (!s) {
    s = ss().insertSheet(name);
    s.appendRow(headers);
  } else if (s.getLastRow() === 0) {
    s.appendRow(headers);
  }
  return s;
}

/** Station tabs = every sheet except the system tabs. */
function getStations() {
  return ss().getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (n) { return n !== META_SHEET && n !== PRESENCE_SHEET && n !== COHORT_SHEET; });
}

/** Create a new empty playlist tab (no-op if it already exists). */
function createPlaylist(station) {
  if (!station) return { ok: false, error: 'missing station name' };
  if (station === META_SHEET || station === PRESENCE_SHEET || station === COHORT_SHEET) {
    return { ok: false, error: 'reserved name' };
  }
  var lock = LockService.getScriptLock(); lock.tryLock(LOCK_WAIT_MS);
  try {
    if (ss().getSheetByName(station)) return { ok: true, created: false };
    var s = ss().insertSheet(station);
    s.appendRow(STATION_HEADERS);
    return { ok: true, created: true };
  } finally { lock.releaseLock(); }
}

// ---- Cohorts / auth ---------------------------------------------------------
/** Read cohort rows; seed defaults if the tab is empty so the app never bricks. */
function getCohorts() {
  var s = ensureSheet(COHORT_SHEET, COHORT_HEADERS);
  if (s.getLastRow() < 2) {
    s.getRange(2, 1, COHORT_SEED.length, COHORT_HEADERS.length).setValues(COHORT_SEED);
  }
  var last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, COHORT_HEADERS.length).getValues()
    .filter(function (r) { return r[1]; })
    .map(function (r) {
      return { label: String(r[0]), passwordHash: String(r[1]).trim(), defaultStation: String(r[2]) };
    });
}

function cohortForAuth(hash) {
  if (!hash) return null;
  var cohorts = getCohorts();
  for (var i = 0; i < cohorts.length; i++) {
    if (cohorts[i].passwordHash === hash) return cohorts[i];
  }
  return null;
}

function isValidAuth(hash) { return !!cohortForAuth(hash); }

/** Return the cohort's default landing playlist for a (validated) password hash. */
function doLogin(hash) {
  var c = cohortForAuth(hash);
  if (!c) return { ok: false, error: 'unauthorized' };
  return { ok: true, defaultStation: c.defaultStation, label: c.label };
}

function getStationSheet(station) {
  var s = ss().getSheetByName(station);
  if (!s) throw new Error('no such station: ' + station);
  // make sure header exists
  if (s.getLastRow() === 0) s.appendRow(STATION_HEADERS);
  return s;
}

/** Returns [{...track, row}], in queue (row) order. */
function readQueue(station) {
  var s = getStationSheet(station);
  var last = s.getLastRow();
  if (last < 2) return [];
  var values = s.getRange(2, 1, last - 1, STATION_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue; // skip blank rows
    out.push({
      id: String(r[0]), videoId: String(r[1]), title: String(r[2]), artist: String(r[3]),
      durationSec: Number(r[4]) || 0, thumbnailUrl: String(r[5]), addedBy: String(r[6]),
      addedAt: Number(r[7]) || 0, row: i + 2
    });
  }
  return out;
}

function getMeta(station) {
  var s = ensureSheet(META_SHEET, META_HEADERS);
  var last = s.getLastRow();
  if (last >= 2) {
    var values = s.getRange(2, 1, last - 1, META_HEADERS.length).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) === station) {
        return {
          row: i + 2,
          station: station,
          currentTrackId: String(values[i][1] || ''),
          trackStartedAt: Number(values[i][2]) || 0,
          lastGongBy: String(values[i][3] || ''),
          lastGongAt: Number(values[i][4]) || 0,
          lastActionAt: Number(values[i][5]) || 0
        };
      }
    }
  }
  // create a fresh row
  s.appendRow([station, '', 0, '', 0, 0]);
  return { row: s.getLastRow(), station: station, currentTrackId: '', trackStartedAt: 0, lastGongBy: '', lastGongAt: 0, lastActionAt: 0 };
}

function writeMeta(meta) {
  var s = ensureSheet(META_SHEET, META_HEADERS);
  meta.lastActionAt = Date.now();
  s.getRange(meta.row, 1, 1, META_HEADERS.length).setValues([[
    meta.station, meta.currentTrackId, meta.trackStartedAt, meta.lastGongBy, meta.lastGongAt, meta.lastActionAt
  ]]);
}

// ---- Clock / advancement ----------------------------------------------------
function indexOfId(queue, id) {
  for (var i = 0; i < queue.length; i++) if (queue[i].id === id) return i;
  return -1;
}

/**
 * Deterministically roll the station's clock forward to "now".
 * Loops back to track 1 at the end. Mutates+persists `meta` only if it changed.
 * Returns the up-to-date meta.
 */
function rollForward(meta, queue) {
  if (!queue.length) {
    if (meta.currentTrackId !== '' || meta.trackStartedAt !== 0) {
      meta.currentTrackId = ''; meta.trackStartedAt = 0; writeMeta(meta);
    }
    return meta;
  }
  var now = Date.now();
  var idx = indexOfId(queue, meta.currentTrackId);
  var changed = false;

  if (idx === -1 || !meta.trackStartedAt) {
    // never started, or current track was removed → (re)start from first track
    idx = 0; meta.currentTrackId = queue[0].id; meta.trackStartedAt = now; changed = true;
  }

  // advance through any tracks whose play-window has fully elapsed
  var guard = 0;
  while (guard++ < 100000) {
    var dur = (queue[idx].durationSec || 0) * 1000;
    if (dur <= 0) break; // unknown duration → don't auto-advance, let client error-skip
    if (now <= meta.trackStartedAt + dur) break;
    meta.trackStartedAt += dur;
    idx = (idx + 1) % queue.length;
    meta.currentTrackId = queue[idx].id;
    changed = true;
  }

  if (changed) writeMeta(meta);
  return meta;
}

// ---- Actions ----------------------------------------------------------------
function getState(station) {
  var queue = readQueue(station);
  var meta = getMeta(station);
  rollForward(meta, queue);
  return {
    station: station,
    currentTrackId: meta.currentTrackId,
    trackStartedAt: meta.trackStartedAt,
    lastGongBy: meta.lastGongBy,
    lastGongAt: meta.lastGongAt,
    lastActionAt: meta.lastActionAt,
    queue: queue.map(function (t) {
      return { id: t.id, videoId: t.videoId, title: t.title, artist: t.artist,
               durationSec: t.durationSec, thumbnailUrl: t.thumbnailUrl, addedBy: t.addedBy, addedAt: t.addedAt };
    })
  };
}

function addTrack(d) {
  var lock = LockService.getScriptLock(); lock.tryLock(LOCK_WAIT_MS);
  try {
    var s = getStationSheet(d.station);
    var id = Utilities.getUuid();
    s.appendRow([id, d.videoId, d.title, d.artist, Number(d.durationSec) || 0, d.thumbnailUrl, d.by || d.addedBy || '', Date.now()]);
    // if the station was empty/silent, start broadcasting this track
    var queue = readQueue(d.station);
    var meta = getMeta(d.station);
    if (!meta.currentTrackId || indexOfId(queue, meta.currentTrackId) === -1) {
      meta.currentTrackId = id; meta.trackStartedAt = Date.now(); writeMeta(meta);
    }
    return { id: id };
  } finally { lock.releaseLock(); }
}

function removeTrack(station, id, by) {
  var lock = LockService.getScriptLock(); lock.tryLock(LOCK_WAIT_MS);
  try {
    var queue = readQueue(station);
    var meta = getMeta(station);
    rollForward(meta, queue);
    var idx = indexOfId(queue, id);
    if (idx === -1) return { ok: true, removed: false };

    // if removing the currently-playing track, advance first (gong-and-delete)
    if (meta.currentTrackId === id) {
      if (queue.length > 1) {
        var nextIdx = (idx + 1) % queue.length;
        meta.currentTrackId = queue[nextIdx].id;
      } else {
        meta.currentTrackId = '';
      }
      meta.trackStartedAt = Date.now();
      writeMeta(meta);
    }
    getStationSheet(station).deleteRow(queue[idx].row);
    return { ok: true, removed: true };
  } finally { lock.releaseLock(); }
}

function playTrack(station, id, by) {
  var lock = LockService.getScriptLock(); lock.tryLock(LOCK_WAIT_MS);
  try {
    var queue = readQueue(station);
    if (indexOfId(queue, id) === -1) return { ok: false, error: 'no such track' };
    var meta = getMeta(station);
    meta.currentTrackId = id;
    meta.trackStartedAt = Date.now();
    writeMeta(meta);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function gong(station, by) {
  var lock = LockService.getScriptLock(); lock.tryLock(LOCK_WAIT_MS);
  try {
    var meta = getMeta(station);
    var now = Date.now();
    if (now - meta.lastGongAt < GONG_COOLDOWN_MS) {
      return { ok: false, cooldown: true, retryInMs: GONG_COOLDOWN_MS - (now - meta.lastGongAt) };
    }
    var queue = readQueue(station);
    rollForward(meta, queue);
    if (!queue.length) return { ok: false, error: 'empty station' };
    var idx = indexOfId(queue, meta.currentTrackId);
    var nextIdx = idx === -1 ? 0 : (idx + 1) % queue.length;
    meta.currentTrackId = queue[nextIdx].id;
    meta.trackStartedAt = now;
    meta.lastGongBy = by || '';
    meta.lastGongAt = now;
    writeMeta(meta);
    return { ok: true, gongedBy: meta.lastGongBy };
  } finally { lock.releaseLock(); }
}

/** Roll forward; if `force` (e.g. unplayable video), skip current track immediately. */
function advance(station, expectedTrackId, force) {
  var lock = LockService.getScriptLock(); lock.tryLock(LOCK_WAIT_MS);
  try {
    var queue = readQueue(station);
    var meta = getMeta(station);
    rollForward(meta, queue);
    // compare-and-set: only force-skip if we're still on the track the caller saw
    if (force && queue.length && meta.currentTrackId === expectedTrackId) {
      var idx = indexOfId(queue, meta.currentTrackId);
      var nextIdx = idx === -1 ? 0 : (idx + 1) % queue.length;
      meta.currentTrackId = queue[nextIdx].id;
      meta.trackStartedAt = Date.now();
      writeMeta(meta);
    }
    return { currentTrackId: meta.currentTrackId, trackStartedAt: meta.trackStartedAt };
  } finally { lock.releaseLock(); }
}

// ---- Presence ---------------------------------------------------------------
function heartbeat(station, handle) {
  var s = ensureSheet(PRESENCE_SHEET, PRESENCE_HEADERS);
  var now = Date.now();
  var last = s.getLastRow();
  var values = last >= 2 ? s.getRange(2, 1, last - 1, PRESENCE_HEADERS.length).getValues() : [];
  var found = -1;
  var listeners = {};
  var staleRows = [];
  for (var i = 0; i < values.length; i++) {
    var st = String(values[i][0]), hd = String(values[i][1]), seen = Number(values[i][2]) || 0;
    if (st === station && hd === handle) found = i + 2;
    if (now - seen > PRESENCE_TTL_MS * 4) { staleRows.push(i + 2); continue; }
    if (st === station && (now - seen < PRESENCE_TTL_MS || (st === station && hd === handle))) {
      listeners[hd] = true;
    }
  }
  if (found > 0) {
    s.getRange(found, 1, 1, 3).setValues([[station, handle, now]]);
  } else {
    s.appendRow([station, handle, now]);
  }
  listeners[handle] = true;
  // prune very old rows (delete bottom-up)
  staleRows.sort(function (a, b) { return b - a; }).forEach(function (r) { try { s.deleteRow(r); } catch (e) {} });
  var names = Object.keys(listeners);
  return { listeners: names, count: names.length };
}

// ---- YouTube search ---------------------------------------------------------
function ytKey() {
  var k = PropertiesService.getScriptProperties().getProperty('YT_API_KEY');
  if (!k) throw new Error('YT_API_KEY not set in Script Properties');
  return k;
}

function searchYouTube(query) {
  if (!query) return [];
  var key = ytKey();
  var url = 'https://www.googleapis.com/youtube/v3/search'
    + '?part=snippet&type=video&videoEmbeddable=true&maxResults=12'
    + '&q=' + encodeURIComponent(query) + '&key=' + key;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var body = JSON.parse(resp.getContentText());
  if (body.error) throw new Error('YouTube: ' + (body.error.message || 'search failed'));
  var items = body.items || [];
  var ids = items.map(function (it) { return it.id.videoId; }).filter(Boolean);
  var durations = fetchDurations(ids, key);
  return items.map(function (it) {
    var vid = it.id.videoId;
    var sn = it.snippet || {};
    var thumb = sn.thumbnails && (sn.thumbnails.medium || sn.thumbnails.default) || {};
    return {
      videoId: vid,
      title: decodeEntities(sn.title || ''),
      artist: decodeEntities(sn.channelTitle || ''),
      durationSec: durations[vid] || 0,
      thumbnailUrl: thumb.url || ''
    };
  }).filter(function (r) { return r.videoId; });
}

function fetchDurations(ids, key) {
  var out = {};
  if (!ids.length) return out;
  var url = 'https://www.googleapis.com/youtube/v3/videos?part=contentDetails&id='
    + ids.join(',') + '&key=' + key;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var body = JSON.parse(resp.getContentText());
  (body.items || []).forEach(function (it) {
    out[it.id] = parseISODuration(it.contentDetails && it.contentDetails.duration);
  });
  return out;
}

function parseISODuration(iso) {
  if (!iso) return 0;
  var m = iso.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!m) return 0;
  return (Number(m[1] || 0) * 3600) + (Number(m[2] || 0) * 60) + Number(m[3] || 0);
}

function decodeEntities(s) {
  return String(s)
    .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>');
}
