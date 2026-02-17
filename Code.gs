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
  STATUS_COLUMN: 7,
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

  const uploadColumnStart = e.range.getColumn();
  const uploadColumnEnd = uploadColumnStart + e.range.getNumColumns() - 1;
  const touchesUploadColumn =
    uploadColumnStart <= CONFIG.UPLOAD_COLUMN && uploadColumnEnd >= CONFIG.UPLOAD_COLUMN;

  if (touchesUploadColumn && e.range.getLastRow() >= CONFIG.OUTPUT_START_ROW) {
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
  const sheet = e.range.getSheet();
  const editedRange = e.range;
  const startRow = Math.max(editedRange.getRow(), CONFIG.OUTPUT_START_ROW);
  const endRow = editedRange.getLastRow();

  if (startRow > endRow) return;

  const checkboxRange = sheet.getRange(
    startRow,
    CONFIG.UPLOAD_COLUMN,
    endRow - startRow + 1,
    1
  );
  const checkboxValues = checkboxRange.getValues();

  let hasCheckedRows = false;

  checkboxValues.forEach((rowValues, index) => {
    const isChecked = rowValues[0] === true;
    if (!isChecked) return;

    hasCheckedRows = true;
    const row = startRow + index;

    try {
      const syncResult = upsertCalendarEventFromRow_(sheet, row);

      if (syncResult.deleted) {
        sheet.getRange(row, CONFIG.ID_COLUMN).clearContent();
        setUploadStatus_(sheet, row, 'Deleted');
        return;
      }

      // Ensure ID in sheet matches final calendar event ID.
      sheet.getRange(row, CONFIG.ID_COLUMN).setValue(syncResult.event.getId());
      setUploadStatus_(sheet, row, 'Uploaded');
    } catch (error) {
      const message = error && error.message ? error.message : String(error);
      setUploadStatus_(sheet, row, `Error: ${message}`);
    }
  });

  if (hasCheckedRows) {
    // Reset upload checkboxes for edited rows in upload column.
    checkboxRange.setValue(false);
  }
}

/**
 * Writes upload status to the cell right of the upload checkbox.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Target sheet.
 * @param {number} row Row number.
 * @param {string} status Status text.
 */
function setUploadStatus_(sheet, row, status) {
  sheet.getRange(row, CONFIG.STATUS_COLUMN).setValue(status);
}

/**
 * Reads a row and updates/creates/deletes the calendar event represented by that row.
 * Empty event titles delete the existing event referenced by ID.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Target sheet.
 * @param {number} row Row number.
 * @returns {{event?: GoogleAppsScript.Calendar.CalendarEvent, deleted: boolean}}
 */
function upsertCalendarEventFromRow_(sheet, row) {
  const idValue = String(sheet.getRange(row, CONFIG.ID_COLUMN).getValue() || '').trim();
  const title = String(sheet.getRange(row, CONFIG.EVENT_COLUMN).getValue() || '').trim();
  const dateCell = sheet.getRange(row, CONFIG.DATE_COLUMN);
  const timeCell = sheet.getRange(row, CONFIG.TIME_COLUMN);
  const dateValue = dateCell.getValue();
  const timeValue = timeCell.getValue();
  const timeDisplayValue = timeCell.getDisplayValue();
  const durationHoursRaw = sheet.getRange(row, CONFIG.DURATION_COLUMN).getValue();

  const calendar = CalendarApp.getDefaultCalendar();
  let calendarEvent = idValue ? calendar.getEventById(idValue) : null;

  if (!title) {
    if (calendarEvent) {
      calendarEvent.deleteEvent();
    }
    return { deleted: true };
  }
  if (!(dateValue instanceof Date) || Number.isNaN(dateValue.getTime())) {
    throw new Error(`Row ${row}: date must be a valid date.`);
  }

  const timeParts = parseTimeValue_(timeValue, timeDisplayValue);

  const durationHours = Number(durationHoursRaw);
  if (!Number.isFinite(durationHours) || durationHours <= 0) {
    throw new Error(`Row ${row}: duration must be a positive number of hours.`);
  }

  const start = combineDateAndTime_(dateValue, timeParts);
  const end = new Date(start.getTime() + durationHours * 60 * 60 * 1000);

  if (calendarEvent) {
    calendarEvent.setTitle(title);
    calendarEvent.setTime(start, end);
    return { event: calendarEvent, deleted: false };
  }

  return { event: calendar.createEvent(title, start, end), deleted: false };
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
      stripTime_(startTime),
      toTimeFraction_(startTime),
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
 * Returns a date-only copy (midnight) of a DateTime.
 *
 * @param {Date} dateTime
 * @returns {Date}
 */
function stripTime_(dateTime) {
  return new Date(dateTime.getFullYear(), dateTime.getMonth(), dateTime.getDate());
}

/**
 * Converts a Date into a Sheets time-only numeric fraction of a day.
 *
 * @param {Date} dateTime
 * @returns {number}
 */
function toTimeFraction_(dateTime) {
  return (
    dateTime.getHours() * 60 * 60 +
    dateTime.getMinutes() * 60 +
    dateTime.getSeconds()
  ) / (24 * 60 * 60);
}

/**
 * Parses time values from sheet value/display into hour/minute components.
 * Supports Date objects, numeric day fractions, and display strings like HH:mm.
 *
 * @param {*} timeValue
 * @param {string} timeDisplayValue
 * @returns {{hours: number, minutes: number}}
 */
function parseTimeValue_(timeValue, timeDisplayValue) {
  if (timeValue instanceof Date && !Number.isNaN(timeValue.getTime())) {
    return {
      hours: timeValue.getHours(),
      minutes: timeValue.getMinutes()
    };
  }

  if (typeof timeValue === 'number' && Number.isFinite(timeValue)) {
    const totalMinutes = Math.round((timeValue % 1) * 24 * 60);
    return {
      hours: Math.floor(totalMinutes / 60) % 24,
      minutes: totalMinutes % 60
    };
  }

  const text = String(timeDisplayValue || timeValue || '').trim();
  const match = text.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (match) {
    const hours = Number(match[1]);
    const minutes = Number(match[2]);
    if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
      return { hours, minutes };
    }
  }

  throw new Error('time must be a valid time (for example 13:00).');
}

/**
 * Combines a date value and time parts into a single Date.
 *
 * @param {Date} dateValue
 * @param {{hours: number, minutes: number}} timeParts
 * @returns {Date}
 */
function combineDateAndTime_(dateValue, timeParts) {
  return new Date(
    dateValue.getFullYear(),
    dateValue.getMonth(),
    dateValue.getDate(),
    timeParts.hours,
    timeParts.minutes,
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
