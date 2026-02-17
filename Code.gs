/**
 * Configuration for calendar sync controls.
 */
const CONFIG = {
  SHEET_NAME: '',
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
  DOWNLOAD_CHECKBOX_RANGE_NAME: 'download',
  PERIOD_START_RANGE_NAME: 'period_start',
  PERIOD_END_RANGE_NAME: 'period_end'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Sync')
    .addItem('Install download trigger', 'setupDownloadTrigger_')
    .addItem('Download now', 'downloadNow')
    .addToUi();
}

/**
 * @param {GoogleAppsScript.Events.SheetsOnOpen=} e
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (e && e.authMode === ScriptApp.AuthMode.LIMITED) return;
  handleSheetEdit_(e);
}

/**
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEditInstallable(e) {
  handleSheetEdit_(e);
}

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
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleSheetEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (CONFIG.SHEET_NAME && sheet.getName() !== CONFIG.SHEET_NAME) return;

  const downloadCheckboxRange = getNamedRangeOrThrow_(CONFIG.DOWNLOAD_CHECKBOX_RANGE_NAME);
  const editedRange = e.range;

  const isDownloadCheckboxEdit =
    editedRange.getA1Notation() === downloadCheckboxRange.getA1Notation() &&
    editedRange.getSheet().getSheetId() === downloadCheckboxRange.getSheet().getSheetId();

  if (isDownloadCheckboxEdit) {
    const isChecked =
      typeof editedRange.isChecked === 'function'
        ? editedRange.isChecked()
        : e.value === 'TRUE' || e.value === true;

    if (isChecked) {
      downloadCalendarEntries_(sheet);
      editedRange.setValue(false);
    }
    return;
  }

  const isUploadCell =
    editedRange.getColumn() === CONFIG.UPLOAD_COLUMN && editedRange.getRow() >= CONFIG.OUTPUT_START_ROW;
  if (!isUploadCell) return;

  const isChecked =
    typeof editedRange.isChecked === 'function'
      ? editedRange.isChecked()
      : e.value === 'TRUE' || e.value === true;
  if (!isChecked) return;

  handleUploadCheckboxEdit_(sheet, editedRange.getRow(), editedRange);
}

function downloadNow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (CONFIG.SHEET_NAME && sheet.getName() !== CONFIG.SHEET_NAME) {
    throw new Error(`Active sheet must be "${CONFIG.SHEET_NAME}".`);
  }
  downloadCalendarEntries_(sheet);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function downloadCalendarEntries_(sheet) {
  const periodStartRange = getNamedRangeOrThrow_(CONFIG.PERIOD_START_RANGE_NAME);
  const periodEndRange = getNamedRangeOrThrow_(CONFIG.PERIOD_END_RANGE_NAME);

  const start = parseSheetDate_(periodStartRange.getValue());
  const end = parseSheetDate_(periodEndRange.getValue());

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

    return [
      event.getId(),
      event.getTitle(),
      startOfDay_(startTime),
      createTimeOnly_(startTime),
      durationHours
    ];
  });

  sheet
    .getRange(
      CONFIG.OUTPUT_START_ROW,
      CONFIG.OUTPUT_START_COLUMN,
      values.length,
      CONFIG.OUTPUT_COLUMN_COUNT
    )
    .setValues(values);

  sheet.getRange(CONFIG.OUTPUT_START_ROW, CONFIG.DATE_COLUMN, values.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(CONFIG.OUTPUT_START_ROW, CONFIG.TIME_COLUMN, values.length, 1).setNumberFormat('HH:mm');
  sheet.getRange(CONFIG.OUTPUT_START_ROW, CONFIG.DURATION_COLUMN, values.length, 1).setNumberFormat('0.##');
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {GoogleAppsScript.Spreadsheet.Range} checkboxRange
 */
function handleUploadCheckboxEdit_(sheet, row, checkboxRange) {
  const statusCell = sheet.getRange(row, CONFIG.STATUS_COLUMN);

  try {
    const resultMessage = upsertOrDeleteCalendarEventFromRow_(sheet, row);
    statusCell.setValue(resultMessage);
  } catch (error) {
    statusCell.setValue(`Error: ${error.message}`);
  } finally {
    checkboxRange.setValue(false);
  }
}

/**
 * If title is empty, delete existing event (if present). Otherwise create/update it.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @returns {string}
 */
function upsertOrDeleteCalendarEventFromRow_(sheet, row) {
  const calendar = CalendarApp.getDefaultCalendar();
  const idValue = String(sheet.getRange(row, CONFIG.ID_COLUMN).getValue() || '').trim();
  const title = String(sheet.getRange(row, CONFIG.EVENT_COLUMN).getValue() || '').trim();

  if (!title) {
    if (!idValue) {
      return 'Skipped: empty title and no event ID to delete.';
    }

    const existing = calendar.getEventById(idValue);
    if (!existing) {
      sheet.getRange(row, CONFIG.ID_COLUMN).clearContent();
      return 'Skipped: event ID not found; nothing deleted.';
    }

    existing.deleteEvent();
    sheet.getRange(row, CONFIG.ID_COLUMN, 1, CONFIG.OUTPUT_COLUMN_COUNT).clearContent();
    return 'Deleted event (empty title).';
  }

  const dateValue = parseSheetDate_(sheet.getRange(row, CONFIG.DATE_COLUMN).getValue());
  if (!dateValue) {
    throw new Error(`Row ${row}: date must be a valid date.`);
  }

  const timeValue = parseSheetTime_(sheet.getRange(row, CONFIG.TIME_COLUMN).getValue());
  if (!timeValue) {
    throw new Error(`Row ${row}: time must be a valid time.`);
  }

  const durationHours = Number(sheet.getRange(row, CONFIG.DURATION_COLUMN).getValue());
  if (!Number.isFinite(durationHours) || durationHours <= 0) {
    throw new Error(`Row ${row}: duration must be a positive number of hours.`);
  }

  const start = combineDateAndTime_(dateValue, timeValue);
  const end = new Date(start.getTime() + durationHours * 60 * 60 * 1000);

  let event = idValue ? calendar.getEventById(idValue) : null;
  if (event) {
    event.setTitle(title);
    event.setTime(start, end);
    sheet.getRange(row, CONFIG.ID_COLUMN).setValue(event.getId());
    return 'Updated existing event.';
  }

  event = calendar.createEvent(title, start, end);
  sheet.getRange(row, CONFIG.ID_COLUMN).setValue(event.getId());
  return 'Created new event.';
}

/**
 * @param {string} rangeName
 * @returns {GoogleAppsScript.Spreadsheet.Range}
 */
function getNamedRangeOrThrow_(rangeName) {
  const range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
  if (!range) {
    throw new Error(`Named range "${rangeName}" is missing.`);
  }
  return range;
}

/**
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
 * @param {Date|string|number} value
 * @returns {Date|null}
 */
function parseSheetTime_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value !== 'string') return null;

  const text = value.trim();
  if (!text) return null;

  const timeMatch = text.match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) {
    const hours = Number(timeMatch[1]);
    const minutes = Number(timeMatch[2]);
    if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
      return new Date(1899, 11, 30, hours, minutes, 0, 0);
    }
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

/**
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
 * @param {Date} date
 * @returns {Date}
 */
function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * @param {Date} time
 * @returns {Date}
 */
function createTimeOnly_(time) {
  return new Date(1899, 11, 30, time.getHours(), time.getMinutes(), time.getSeconds(), 0);
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
