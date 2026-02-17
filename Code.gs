/**
 * Configuration for calendar download controls.
 */
const CONFIG = {
  SHEET_NAME: '',
  OUTPUT_START_ROW: 2,
  OUTPUT_START_COLUMN: 1, // A
  OUTPUT_COLUMN_COUNT: 5, // A:E => id, event, date, time, duration
  DOWNLOAD_CHECKBOX_CELL: 'I1',
  PERIOD_START_CELL: 'I3',
  PERIOD_END_CELL: 'I4'
};

/**
 * Adds a menu for one-time trigger installation and manual download.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Sync')
    .addItem('Install download trigger', 'setupDownloadTrigger_')
    .addItem('Download now', 'downloadNow')
    .addToUi();
}

/**
 * Ensures the custom menu is added after installation.
 *
 * @param {GoogleAppsScript.Events.SheetsOnOpen=} e
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Simple trigger guard. CalendarApp requires installable trigger auth.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (e && e.authMode === ScriptApp.AuthMode.LIMITED) return;
  handleSheetEdit_(e);
}

/**
 * Installable edit trigger entry point.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEditInstallable(e) {
  handleSheetEdit_(e);
}

/**
 * Creates one installable ON_EDIT trigger for download checkbox handling.
 */
function setupDownloadTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(
    (trigger) =>
      trigger.getHandlerFunction() === 'onEditInstallable' &&
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT
  );

  if (!exists) {
    ScriptApp.newTrigger('onEditInstallable')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
    SpreadsheetApp.getActive().toast('Download trigger installed.', 'Calendar Sync', 5);
    return;
  }

  SpreadsheetApp.getActive().toast('Download trigger already installed.', 'Calendar Sync', 5);
}

/**
 * Runs a download when the configured checkbox cell is checked.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleSheetEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (CONFIG.SHEET_NAME && sheet.getName() !== CONFIG.SHEET_NAME) return;
  if (e.range.getA1Notation() !== CONFIG.DOWNLOAD_CHECKBOX_CELL) return;

  const isChecked =
    typeof e.range.isChecked === 'function'
      ? e.range.isChecked()
      : e.value === 'TRUE' || e.value === true;
  if (!isChecked) return;

  downloadCalendarEntries_(sheet);
  e.range.setValue(false);
}

/**
 * Manual helper for download.
 */
function downloadNow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (CONFIG.SHEET_NAME && sheet.getName() !== CONFIG.SHEET_NAME) {
    throw new Error(`Active sheet must be "${CONFIG.SHEET_NAME}".`);
  }
  downloadCalendarEntries_(sheet);
}

/**
 * Downloads events in the inclusive period range and writes them to A:E.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function downloadCalendarEntries_(sheet) {
  const start = parseSheetDate_(sheet.getRange(CONFIG.PERIOD_START_CELL).getValue());
  const end = parseSheetDate_(sheet.getRange(CONFIG.PERIOD_END_CELL).getValue());

  if (!start || !end) {
    throw new Error('period_start and period_end must both be valid dates.');
  }

  const startAtMidnight = startOfDay_(start);
  const endExclusive = addDays_(startOfDay_(end), 1);

  const events = CalendarApp.getDefaultCalendar().getEvents(startAtMidnight, endExclusive);

  const lastRow = Math.max(sheet.getLastRow(), CONFIG.OUTPUT_START_ROW);
  const rowsToClear = Math.max(0, lastRow - CONFIG.OUTPUT_START_ROW + 1);
  if (rowsToClear > 0) {
    sheet
      .getRange(
        CONFIG.OUTPUT_START_ROW,
        CONFIG.OUTPUT_START_COLUMN,
        rowsToClear,
        CONFIG.OUTPUT_COLUMN_COUNT
      )
      .clearContent();
  }

  if (events.length === 0) return;

  const values = events.map((event) => {
    const startTime = event.getStartTime();
    const durationHours =
      (event.getEndTime().getTime() - startTime.getTime()) / (1000 * 60 * 60);

    return [event.getId(), event.getTitle(), startTime, startTime, durationHours];
  });

  sheet
    .getRange(
      CONFIG.OUTPUT_START_ROW,
      CONFIG.OUTPUT_START_COLUMN,
      values.length,
      CONFIG.OUTPUT_COLUMN_COUNT
    )
    .setValues(values);

  sheet.getRange(CONFIG.OUTPUT_START_ROW, 4, values.length, 1).setNumberFormat('HH:mm');
  sheet.getRange(CONFIG.OUTPUT_START_ROW, 5, values.length, 1).setNumberFormat('0.##');
}

/**
 * Parses sheet dates from Date cells or common day-first manual text input.
 *
 * @param {Date|string|number} value
 * @returns {Date|null}
 */
function parseSheetDate_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value !== 'string') return null;

  const text = value.trim();
  if (!text) return null;

  const dayFirstMatch = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (dayFirstMatch) {
    const day = Number(dayFirstMatch[1]);
    const month = Number(dayFirstMatch[2]);
    const year = Number(dayFirstMatch[3]);
    const parsed = new Date(year, month - 1, day);

    if (
      parsed.getFullYear() === year &&
      parsed.getMonth() === month - 1 &&
      parsed.getDate() === day
    ) {
      return parsed;
    }
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * @param {Date} date
 * @returns {Date}
 */
function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * @param {Date} date
 * @param {number} days
 * @returns {Date}
 */
function addDays_(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}
