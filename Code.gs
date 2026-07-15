/**
 * PFE Journal — Google Apps Script sync backend.
 *
 * Setup:
 *  1. Create a new Google Sheet (any name, e.g. "PFE Journal").
 *  2. Extensions > Apps Script. Delete the default code, paste this file, Save.
 *  3. Deploy > New deployment > type "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone   (required for the phone app to reach it)
 *  4. Copy the /exec URL. Paste it into the app's Settings > endpoint field.
 *  5. Any Code.gs change later needs Deploy > Manage deployments > Edit > New version.
 *
 * Data model: one "Journal" sheet, one row per record.
 *  id = kind + ":" + date   (kind = "day" for daily entries, "feed" for weekly feed snapshots)
 *  Row-level upsert, newer updatedAt wins. Deletes via the deleted flag.
 */

const SHEET_JOURNAL = 'Journal';
const SHEET_META = 'Meta';
const HEADERS = ['id', 'kind', 'date', 'body', 'updatedAt', 'deleted'];

function doGet() {
  return json_(readAll_());
}

function doPost(e) {
  var payload = {};
  try { payload = JSON.parse(e.postData.contents || '{}'); } catch (err) { payload = {}; }
  var rows = payload.rows || [];
  upsert_(rows);
  return json_(readAll_());
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function readAll_() {
  var sh = sheet_(SHEET_JOURNAL, HEADERS);
  var vals = sh.getDataRange().getValues();
  var rows = [];
  if (vals.length > 1) {
    var hdr = vals[0];
    for (var i = 1; i < vals.length; i++) {
      var r = {};
      for (var c = 0; c < hdr.length; c++) r[hdr[c]] = vals[i][c];
      if (r.id) {
        rows.push({
          id: String(r.id),
          kind: String(r.kind),
          date: normDate_(r.date),
          body: String(r.body),
          updatedAt: normStamp_(r.updatedAt),
          deleted: (r.deleted === true || r.deleted === 'true')
        });
      }
    }
  }
  return { ok: true, lastModified: getMeta_('lastModified') || '', rows: rows };
}

// Sheets auto-coerces plain "YYYY-MM-DD" strings into Date objects. Force the
// date column to plain text and format any Date back to YYYY-MM-DD on read.
function normDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}
function normStamp_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return v.toISOString();
  return String(v);
}

function upsert_(incoming) {
  var sh = sheet_(SHEET_JOURNAL, HEADERS);
  sh.getRange(1, 3, Math.max(sh.getMaxRows(), 2), 1).setNumberFormat('@'); // date column = text
  var vals = sh.getDataRange().getValues();
  var idx = {};
  for (var i = 1; i < vals.length; i++) {
    var id = vals[i][0];
    if (id) idx[String(id)] = i; // 0-based into vals; sheet row = i+1
  }
  var changed = false;
  for (var j = 0; j < incoming.length; j++) {
    var r = incoming[j];
    if (!r || !r.id) continue;
    var rowArr = [
      String(r.id), r.kind || '', r.date || '',
      r.body == null ? '' : String(r.body),
      r.updatedAt || '', r.deleted ? true : false
    ];
    if (idx.hasOwnProperty(r.id)) {
      var existingUpdated = String(vals[idx[r.id]][4]);
      if (String(r.updatedAt) > existingUpdated) {
        sh.getRange(idx[r.id] + 1, 1, 1, HEADERS.length).setValues([rowArr]);
        changed = true;
      }
    } else {
      sh.appendRow(rowArr);
      changed = true;
    }
  }
  if (changed) setMeta_('lastModified', new Date().toISOString());
}

function getMeta_(k) {
  var sh = sheet_(SHEET_META, ['key', 'value']);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) if (vals[i][0] === k) return vals[i][1];
  return '';
}

function setMeta_(k, v) {
  var sh = sheet_(SHEET_META, ['key', 'value']);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === k) { sh.getRange(i + 1, 2).setValue(v); return; }
  }
  sh.appendRow([k, v]);
}
