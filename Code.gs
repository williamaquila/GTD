/**
 * Configuration for the sheet and calendar download controls.
 */
const CONFIG = {
  SHEET_NAME: '',
  OUTPUT_HEADER_ROW: 1,
  OUTPUT_START_ROW: 2,
  OUTPUT_START_COLUMN: 1, // A
  OUTPUT_COLUMN_COUNT: 5, // A:E => id, event, date, time, duration
  ID_COLUMN: 1,
  EVENT_COLUMN: 2,
  DATE_COLUMN: 3,
  TIME_COLUMN: 4,
  DURATION_COLUMN: 5,
  UPLOAD_COLUMN: 6,
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
 * Ensures the custom menu is also added when the project is installed.
 *
 * @param {GoogleAppsScript.Events.SheetsOnOpen=} e Open event object.
 */
function onInstall(e) {
  onOpen(e);
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
  handleSheetEdit_(e);
}

/**
 * Installable edit trigger entry point.
 * Create this trigger once by running setupDownloadTrigger_.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Edit event object.
 */
function onEditInstallable(e) {
  handleSheetEdit_(e);
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
 * Routes sheet edits to download/upload actions.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Edit event object.
 */
function handleSheetEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (CONFIG.SHEET_NAME && sheet.getName() !== CONFIG.SHEET_NAME) return;

  if (e.range.getA1Notation() === CONFIG.DOWNLOAD_CHECKBOX_CELL) {
    handleDownloadCheckboxEdit_(e);
    return;
  }

  if (e.range.getColumn() === CONFIG.UPLOAD_COLUMN && e.range.getRow() >= CONFIG.OUTPUT_START_ROW) {
    handleUploadCheckboxEdit_(e);
  }
}

/**
 * Handles download checkbox edits and runs download when checkbox is checked.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Edit event object.
 */
function handleDownloadCheckboxEdit_(e) {
  const isChecked =
    typeof e.range.isChecked === 'function'
      ? e.range.isChecked()
      : e.value === 'TRUE' || e.value === true;
  if (!isChecked) return;

  downloadCalendarEntries_(e.range.getSheet());
  e.range.setValue(false);
}

/**
 * Handles upload checkbox edits and pushes the row content back to Calendar.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Edit event object.
 */
function handleUploadCheckboxEdit_(e) {
  const isChecked =
    typeof e.range.isChecked === 'function'
      ? e.range.isChecked()
      : e.value === 'TRUE' || e.value === true;
  if (!isChecked) return;

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const event = upsertCalendarEventFromRow_(sheet, row);

  // Ensure ID in sheet matches final calendar event ID.
  sheet.getRange(row, CONFIG.ID_COLUMN).setValue(event.getId());

  // Reset upload checkbox.
  e.range.setValue(false);
}

/**
 * Reads a row and updates/creates the calendar event represented by that row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Target sheet.
 * @param {number} row Row number.
 * @returns {GoogleAppsScript.Calendar.CalendarEvent}
 */
function upsertCalendarEventFromRow_(sheet, row) {
  const idValue = String(sheet.getRange(row, CONFIG.ID_COLUMN).getValue() || '').trim();
  const title = String(sheet.getRange(row, CONFIG.EVENT_COLUMN).getValue() || '').trim();
  const dateValue = sheet.getRange(row, CONFIG.DATE_COLUMN).getValue();
  const timeValue = sheet.getRange(row, CONFIG.TIME_COLUMN).getValue();
  const durationHoursRaw = sheet.getRange(row, CONFIG.DURATION_COLUMN).getValue();

  if (!title) {
    throw new Error(`Row ${row}: event title is required.`);
  }
  if (!(dateValue instanceof Date) || Number.isNaN(dateValue.getTime())) {
    throw new Error(`Row ${row}: date must be a valid date.`);
  }
  if (!(timeValue instanceof Date) || Number.isNaN(timeValue.getTime())) {
    throw new Error(`Row ${row}: time must be a valid time.`);
  }

  const durationHours = Number(durationHoursRaw);
  if (!Number.isFinite(durationHours) || durationHours <= 0) {
    throw new Error(`Row ${row}: duration must be a positive number of hours.`);
  }

  const start = combineDateAndTime_(dateValue, timeValue);
  const end = new Date(start.getTime() + durationHours * 60 * 60 * 1000);

  const calendar = CalendarApp.getDefaultCalendar();
  let calendarEvent = idValue ? calendar.getEventById(idValue) : null;

  if (calendarEvent) {
    calendarEvent.setTitle(title);
    calendarEvent.setTime(start, end);
    return calendarEvent;
  }

  return calendar.createEvent(title, start, end);
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
  const start = parseSheetDate_(startRaw);
  const end = parseSheetDate_(endRaw);

  if (!start || !end) {
    throw new Error('period_start and period_end must both be valid dates.');
  }

  const startAtMidnight = startOfDay_(start);
  const endExclusive = addDays_(startOfDay_(end), 1); // inclusive end date

  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(startAtMidnight, endExclusive);

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

  // Keep sheet-defined formatting for date column; format time/duration only.
  sheet
    .getRange(CONFIG.OUTPUT_START_ROW, 4, values.length, 1)
    .setNumberFormat('HH:mm');
  sheet
    .getRange(CONFIG.OUTPUT_START_ROW, 5, values.length, 1)
    .setNumberFormat('0.##');
}

/**
 * Parses sheet values that may be Date objects or text input (for example: dd/mm/yyyy).
 *
 * @param {Date|string|number} value
 * @returns {Date|null}
 */
function parseSheetDate_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value !== 'string') {
    return null;
  }

  const text = value.trim();
  if (!text) {
    return null;
  }

  // Prefer day-first for slash/dot-separated values entered manually in many sheet locales.
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

  // Fallback for ISO (yyyy-mm-dd) and other parseable values.
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed;
  }

  return null;
}

/**
 * Combines a date value and time value into a single Date.
 *
 * @param {Date} dateValue
 * @param {Date} timeValue
 * @returns {Date}
 */
function combineDateAndTime_(dateValue, timeValue) {
  return new Date(
    dateValue.getFullYear(),
    dateValue.getMonth(),
    dateValue.getDate(),
    timeValue.getHours(),
    timeValue.getMinutes(),
    0,
    0
  );
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
