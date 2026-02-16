/**
 * Configuration for the sheet and calendar download controls.
 */
const CONFIG = {
  SHEET_NAME: '',
  OUTPUT_HEADER_ROW: 1,
  OUTPUT_START_ROW: 2,
  OUTPUT_START_COLUMN: 1, // A
  OUTPUT_COLUMN_COUNT: 5, // A:E => id, event, date, time, duration
  DOWNLOAD_CHECKBOX_CELL: 'J1',
  PERIOD_START_CELL: 'J3',
  PERIOD_END_CELL: 'J4'
};

/**
 * Adds a custom menu so trigger setup and manual download are available in UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Sync')
    .addItem('Install download trigger', 'setupDownloadTrigger_')
    .addItem('Download now', 'downloadNow')
    .addToUi();
}

/**
 * Simple trigger.
 * Note: simple onEdit triggers run in LIMITED auth mode and cannot call CalendarApp.
 * Keep this as a no-op guard so accidental simple-trigger execution is harmless.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Edit event object.
 */
function onEdit(e) {
  if (e && e.authMode === ScriptApp.AuthMode.LIMITED) {
    return;
  }
  handleDownloadCheckboxEdit_(e);
}

/**
 * Installable edit trigger entry point.
 * Create this trigger once by running setupDownloadTrigger_.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Edit event object.
 */
function onEditInstallable(e) {
  handleDownloadCheckboxEdit_(e);
}

/**
 * Creates the required installable edit trigger (one-time setup).
 * Run this manually once from Apps Script editor.
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
 * Handles checkbox edits and runs download when checkbox is checked.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Edit event object.
 */
function handleDownloadCheckboxEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (CONFIG.SHEET_NAME && sheet.getName() !== CONFIG.SHEET_NAME) return;
  if (e.range.getA1Notation() !== CONFIG.DOWNLOAD_CHECKBOX_CELL) return;

  const isChecked = typeof e.range.isChecked === 'function' ? e.range.isChecked() : e.value === 'TRUE' || e.value === true;
  if (!isChecked) return;

  downloadCalendarEntries_(sheet);
  e.range.setValue(false);
}

/**
 * Manual helper for running the same download without editing checkbox.
 */
function downloadNow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (CONFIG.SHEET_NAME && sheet.getName() !== CONFIG.SHEET_NAME) {
    throw new Error(`Active sheet must be "${CONFIG.SHEET_NAME}".`);
  }
  downloadCalendarEntries_(sheet);
}

/**
 * Downloads events in the inclusive period range and writes them to columns A:E.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Target sheet.
 */
function downloadCalendarEntries_(sheet) {
  const startRaw = sheet.getRange(CONFIG.PERIOD_START_CELL).getValue();
  const endRaw = sheet.getRange(CONFIG.PERIOD_END_CELL).getValue();

  if (!(startRaw instanceof Date) || !(endRaw instanceof Date)) {
    throw new Error('period_start and period_end must both be valid dates.');
  }

  const start = startOfDay_(startRaw);
  const endExclusive = addDays_(startOfDay_(endRaw), 1); // inclusive end date

  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(start, endExclusive);

  // Clear previous output rows from A:E while keeping headers.
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

    return [
      event.getId(),
      event.getTitle(),
      startTime,
      startTime,
      durationHours
    ];
  });

  const outputRange = sheet.getRange(
    CONFIG.OUTPUT_START_ROW,
    CONFIG.OUTPUT_START_COLUMN,
    values.length,
    CONFIG.OUTPUT_COLUMN_COUNT
  );
  outputRange.setValues(values);

  // Apply display formats for date/time/duration columns.
  sheet
    .getRange(CONFIG.OUTPUT_START_ROW, 3, values.length, 1)
    .setNumberFormat('dd/MM/yyyy');
  sheet
    .getRange(CONFIG.OUTPUT_START_ROW, 4, values.length, 1)
    .setNumberFormat('HH:mm');
  sheet
    .getRange(CONFIG.OUTPUT_START_ROW, 5, values.length, 1)
    .setNumberFormat('0.##');
}

/**
 * Returns a copy of the date at midnight.
 *
 * @param {Date} date
 * @returns {Date}
 */
function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Returns a new date offset by a given number of days.
 *
 * @param {Date} date
 * @param {number} days
 * @returns {Date}
 */
function addDays_(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}
