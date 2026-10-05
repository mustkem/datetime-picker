/**
 * DateTime Picker - a Google Sheets editor add-on.
 * Pick a date and a time together, in any time zone, and insert it into the
 * selected cells, as ISO 8601 text (default) or as a formatted date value.
 *
 * Open source under the MIT License.
 * https://github.com/mustkem/datetime-picker
 */

var DEFAULTS = {
  zone: 'GMT',
  output: 'iso',      // iso: ISO 8601 text, e.g. 2026-11-12T16:00:00.000Z | date: real date value
  toGmt: false,       // convert the picked time to GMT before inserting
  dateOrder: 'dmy',   // date output only: dmy | mdy | ymd
  clock: '12',        // date output only: 12 | 24
  showZone: true      // date output only: append the zone label to the cell format
};

var DATE_PATTERNS = { dmy: 'dd/mm/yyyy', mdy: 'mm/dd/yyyy', ymd: 'yyyy-mm-dd' };
var ISO_RE = /^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2})(?::\d{2}(?:\.\d+)?)?(Z|[+-]\d{2}:?\d{2})$/;

/* ---------- Menu ---------- */

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Pick date and time', 'openPicker')
    .addItem('Insert current date and time', 'insertNow')
    .addSeparator()
    .addItem('Settings', 'openSettings')
    .addItem('Help', 'openHelp')
    .addToUi();
}

/* ---------- Dialogs ---------- */

function openPicker() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Select one or more cells first.');
    return;
  }
  var settings = getSettings();
  var current = readCell_(range.getCell(1, 1));

  var t = HtmlService.createTemplateFromFile('Picker');
  t.data = JSON.stringify({
    target: range.getSheet().getName() + '!' + range.getA1Notation(),
    cellCount: range.getNumRows() * range.getNumColumns(),
    current: current.value,
    zone: settings.zone, // always open on the default zone (GMT+0 unless changed in Settings)
    toGmt: settings.toGmt,
    output: settings.output
  });
  SpreadsheetApp.getUi().showModalDialog(
    t.evaluate().setWidth(380).setHeight(350),
    'Pick date and time'
  );
}

function openSettings() {
  var t = HtmlService.createTemplateFromFile('Settings');
  t.data = JSON.stringify(getSettings());
  SpreadsheetApp.getUi().showModalDialog(
    t.evaluate().setWidth(380).setHeight(460),
    'DateTime Picker settings'
  );
}

function openHelp() {
  var html = HtmlService.createTemplateFromFile('Help').evaluate().setWidth(400).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'DateTime Picker help');
}

/* ---------- Settings (stored per spreadsheet) ---------- */

function getSettings() {
  var raw = PropertiesService.getDocumentProperties().getProperty('settings');
  var saved = raw ? JSON.parse(raw) : {};
  var out = {};
  for (var k in DEFAULTS) out[k] = saved.hasOwnProperty(k) ? saved[k] : DEFAULTS[k];
  return out;
}

function saveSettings(settings) {
  var clean = {
    zone: isValidZone_(settings.zone) ? settings.zone : DEFAULTS.zone,
    output: settings.output === 'date' ? 'date' : 'iso',
    toGmt: !!settings.toGmt,
    dateOrder: DATE_PATTERNS[settings.dateOrder] ? settings.dateOrder : DEFAULTS.dateOrder,
    clock: settings.clock === '24' ? '24' : '12',
    showZone: !!settings.showZone
  };
  PropertiesService.getDocumentProperties().setProperty('settings', JSON.stringify(clean));
  return clean;
}

/* ---------- Inserting ---------- */

/**
 * Called from the picker.
 * @param {string} value  "yyyy-MM-ddTHH:mm" wall-clock time in `zone`
 * @param {string} zone   "GMT" or an IANA time zone id
 * @param {string} target "Sheet!A1:B2"
 * @param {boolean} toGmt convert to GMT before inserting
 * @return {string} the text that was inserted (ISO) or a description
 */
function insertDateTime(value, zone, target, toGmt) {
  var m = String(value).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/);
  if (!m) throw new Error('Pick a valid date and time.');
  if (!isValidZone_(zone)) throw new Error('Unknown time zone: ' + zone);

  // The zone chosen here applies to this insert only; the default stays as set
  // in Settings (GMT+0 unless the user changes it there).
  var settings = getSettings();
  settings.toGmt = !!toGmt;

  var instant = zone === 'GMT'
    ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]))
    : Utilities.parseDate(value, zone, "yyyy-MM-dd'T'HH:mm");
  var outZone = settings.toGmt ? 'GMT' : zone;
  return write_(target, instant, outZone, settings);
}

function insertNow() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Select one or more cells first.');
    return;
  }
  var settings = getSettings();
  var zone = settings.toGmt ? 'GMT' : settings.zone;
  write_(range.getSheet().getName() + '!' + range.getA1Notation(), new Date(), zone, settings);
}

/** Writes `instant`, expressed in `zone`, into every cell of `target`. */
function write_(target, instant, zone, settings) {
  var range = SpreadsheetApp.getActive().getRange(target);
  var value, format;

  if (settings.output === 'date') {
    // Store the wall-clock time as a serial number, so the cell shows exactly
    // this time no matter what time zone the spreadsheet is set to.
    var w = Utilities.formatDate(instant, zone, 'yyyy,MM,dd,HH,mm').split(',').map(Number);
    value = Date.UTC(w[0], w[1] - 1, w[2], w[3], w[4]) / 86400000 + 25569;
    format = buildFormat_(settings, zone);
  } else {
    value = toIso_(instant, zone);
    format = '@'; // plain text, so Sheets keeps the ISO string as typed
  }

  var values = [];
  for (var r = 0; r < range.getNumRows(); r++) {
    var row = [];
    for (var c = 0; c < range.getNumColumns(); c++) row.push(value);
    values.push(row);
  }
  range.setNumberFormat(format).setValues(values);
  return String(value);
}

/** ISO 8601 with milliseconds: "2026-11-12T16:00:00.000Z" or "...T21:30:00.000+05:30". */
function toIso_(instant, zone) {
  if (zone === 'GMT') return instant.toISOString();
  var local = Utilities.formatDate(instant, zone, "yyyy-MM-dd'T'HH:mm:ss.SSS");
  var z = Utilities.formatDate(instant, zone, 'Z'); // e.g. +0530
  return local + (z === '+0000' ? 'Z' : z.slice(0, 3) + ':' + z.slice(3));
}

function buildFormat_(settings, label) {
  var time = settings.clock === '24' ? 'hh:mm' : 'hh:mm AM/PM';
  var fmt = DATE_PATTERNS[settings.dateOrder] + ' ' + time;
  if (settings.showZone) fmt += ' "' + String(label).replace(/"/g, '') + '"';
  return fmt;
}

/* ---------- Reading the current cell ---------- */

// Returns { value: "yyyy-MM-ddTHH:mm" or "", zone: zone id or "" } for the picker.
function readCell_(cell) {
  var v = cell.getValue();

  if (typeof v === 'string') {
    var m = v.trim().match(ISO_RE);
    if (!m) return { value: '', zone: '' };
    if (m[2] === 'Z') return { value: m[1], zone: 'GMT' };
    // An offset like +05:30 does not name a zone; show the time as GMT instead.
    var instant = new Date(v.trim());
    if (isNaN(instant)) return { value: '', zone: '' };
    return { value: Utilities.formatDate(instant, 'GMT', "yyyy-MM-dd'T'HH:mm"), zone: 'GMT' };
  }

  if (v instanceof Date) {
    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    return {
      value: Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm"),
      zone: zoneFromFormat_(cell.getNumberFormat())
    };
  }
  return { value: '', zone: '' };
}

// Reads the zone label written into a cell's number format, if any.
function zoneFromFormat_(format) {
  var m = String(format || '').match(/"([^"]+)"\s*$/);
  return m && isValidZone_(m[1]) ? m[1] : '';
}

function isValidZone_(zone) {
  if (zone === 'GMT') return true;
  return /^[A-Za-z_]+(\/[A-Za-z0-9_+\-]+){1,2}$/.test(String(zone));
}
