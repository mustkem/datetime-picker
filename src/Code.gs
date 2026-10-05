/**
 * DateTime Picker - a Google Sheets editor add-on.
 * Pick a date and a time together, in any time zone, and insert it into the
 * selected cells as a real date value.
 *
 * Open source under the MIT License.
 * https://github.com/mustkem/datetime-picker
 */

var DEFAULTS = {
  zone: 'GMT',
  dateOrder: 'dmy',   // dmy | mdy | ymd
  clock: '12',        // 12 | 24
  showZone: true,     // append the zone label to the cell format
  toGmt: false        // convert the picked time to GMT before inserting
};

var DATE_PATTERNS = { dmy: 'dd/mm/yyyy', mdy: 'mm/dd/yyyy', ymd: 'yyyy-mm-dd' };

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
  var first = range.getCell(1, 1);

  var t = HtmlService.createTemplateFromFile('Picker');
  t.data = JSON.stringify({
    target: range.getSheet().getName() + '!' + range.getA1Notation(),
    cellCount: range.getNumRows() * range.getNumColumns(),
    current: currentValue_(first),
    zone: zoneFromFormat_(first.getNumberFormat()) || settings.zone,
    toGmt: settings.toGmt
  });
  SpreadsheetApp.getUi().showModalDialog(
    t.evaluate().setWidth(380).setHeight(330),
    'Pick date and time'
  );
}

function openSettings() {
  var t = HtmlService.createTemplateFromFile('Settings');
  t.data = JSON.stringify(getSettings());
  SpreadsheetApp.getUi().showModalDialog(
    t.evaluate().setWidth(380).setHeight(400),
    'DateTime Picker settings'
  );
}

function openHelp() {
  var html = HtmlService.createTemplateFromFile('Help').evaluate().setWidth(400).setHeight(330);
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
    dateOrder: DATE_PATTERNS[settings.dateOrder] ? settings.dateOrder : DEFAULTS.dateOrder,
    clock: settings.clock === '24' ? '24' : '12',
    showZone: !!settings.showZone,
    toGmt: !!settings.toGmt
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
 */
function insertDateTime(value, zone, target, toGmt) {
  var m = String(value).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/);
  if (!m) throw new Error('Pick a valid date and time.');
  if (!isValidZone_(zone)) throw new Error('Unknown time zone: ' + zone);

  // Remember the last zone and GMT choice for this spreadsheet.
  var settings = getSettings();
  settings.zone = zone;
  settings.toGmt = !!toGmt;
  settings = saveSettings(settings);

  var wall = [+m[1], +m[2], +m[3], +m[4], +m[5]];
  var label = zone;
  if (toGmt && zone !== 'GMT') {
    var instant = Utilities.parseDate(value, zone, "yyyy-MM-dd'T'HH:mm");
    wall = Utilities.formatDate(instant, 'GMT', 'yyyy,MM,dd,HH,mm').split(',').map(Number);
    label = 'GMT';
  }

  writeWallClock_(target, wall, label, settings);
  return true;
}

function insertNow() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Select one or more cells first.');
    return;
  }
  var settings = getSettings();
  var zone = settings.toGmt ? 'GMT' : settings.zone;
  var wall = Utilities.formatDate(new Date(), zone, 'yyyy,MM,dd,HH,mm').split(',').map(Number);
  writeWallClock_(range.getSheet().getName() + '!' + range.getA1Notation(), wall, zone, settings);
}

/**
 * Writes the wall-clock time as a spreadsheet serial number, so the cell shows
 * exactly what was picked no matter what time zone the spreadsheet is set to.
 */
function writeWallClock_(target, wall, label, settings) {
  var ms = Date.UTC(wall[0], wall[1] - 1, wall[2], wall[3], wall[4]);
  var serial = ms / 86400000 + 25569; // days since 1899-12-30

  var range = SpreadsheetApp.getActive().getRange(target);
  var values = [];
  for (var r = 0; r < range.getNumRows(); r++) {
    var row = [];
    for (var c = 0; c < range.getNumColumns(); c++) row.push(serial);
    values.push(row);
  }
  range.setNumberFormat(buildFormat_(settings, label)).setValues(values);
}

function buildFormat_(settings, label) {
  var time = settings.clock === '24' ? 'hh:mm' : 'hh:mm AM/PM';
  var fmt = DATE_PATTERNS[settings.dateOrder] + ' ' + time;
  if (settings.showZone) fmt += ' "' + String(label).replace(/"/g, '') + '"';
  return fmt;
}

/* ---------- Helpers ---------- */

// Current cell value as "yyyy-MM-ddTHH:mm" (as displayed), or "" if not a date.
function currentValue_(cell) {
  var v = cell.getValue();
  if (!(v instanceof Date)) return '';
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm");
}

// Reads the zone label we wrote into a cell's number format, if any.
function zoneFromFormat_(format) {
  var m = String(format || '').match(/"([^"]+)"\s*$/);
  return m && isValidZone_(m[1]) ? m[1] : '';
}

function isValidZone_(zone) {
  if (zone === 'GMT') return true;
  if (!/^[A-Za-z_]+(\/[A-Za-z0-9_+\-]+){1,2}$/.test(String(zone))) return false;
  try {
    Utilities.formatDate(new Date(), zone, 'Z');
    return true;
  } catch (e) {
    return false;
  }
}
