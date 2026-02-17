/**
 * Configuration for calendar sync controls.
 */
const CONFIG = {
  SHEET_NAME: '',
  HEADER_ROW: 1,
  OUTPUT_START_ROW: 2,
  DOWNLOAD_CHECKBOX_RANGE_NAME: 'download',
  PERIOD_START_RANGE_NAME: 'period_start',
  PERIOD_END_RANGE_NAME: 'period_end',
  HEADER_NAMES: {
    id: 'id',
    event: 'event',
    date: 'date',
    time: 'time',
    duration: 'duration',
    upload: 'upload',
    status: 'status'
  }
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

  const columns = getColumnIndexes_(sheet);
  const isUploadCell =
    editedRange.getColumn() === columns.upload && editedRange.getRow() >= CONFIG.OUTPUT_START_ROW;
  if (!isUploadCell) return;

  const isChecked =
    typeof editedRange.isChecked === 'function'
      ? editedRange.isChecked()
      : e.value === 'TRUE' || e.value === true;
  if (!isChecked) return;

  handleUploadCheckboxEdit_(sheet, editedRange.getRow(), editedRange, columns);
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
  const columns = getColumnIndexes_(sheet);
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
    [columns.id, columns.event, columns.date, columns.time, columns.duration].forEach((column) => {
      sheet.getRange(CONFIG.OUTPUT_START_ROW, column, rowsToClear, 1).clearContent();
    });
  }

  if (events.length === 0) return;

  const idValues = [];
  const titleValues = [];
  const dateValues = [];
  const timeValues = [];
  const durationValues = [];

  events.forEach((event) => {
    const startTime = event.getStartTime();
    const durationHours = (event.getEndTime().getTime() - startTime.getTime()) / (1000 * 60 * 60);

    idValues.push([event.getId()]);
    titleValues.push([event.getTitle()]);
    dateValues.push([startOfDay_(startTime)]);
    timeValues.push([createTimeOnly_(startTime)]);
    durationValues.push([durationHours]);
  });

  sheet.getRange(CONFIG.OUTPUT_START_ROW, columns.id, idValues.length, 1).setValues(idValues);
  sheet.getRange(CONFIG.OUTPUT_START_ROW, columns.event, titleValues.length, 1).setValues(titleValues);
  sheet.getRange(CONFIG.OUTPUT_START_ROW, columns.date, dateValues.length, 1).setValues(dateValues);
  sheet.getRange(CONFIG.OUTPUT_START_ROW, columns.time, timeValues.length, 1).setValues(timeValues);
  sheet.getRange(CONFIG.OUTPUT_START_ROW, columns.duration, durationValues.length, 1).setValues(durationValues);

  sheet.getRange(CONFIG.OUTPUT_START_ROW, columns.date, dateValues.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(CONFIG.OUTPUT_START_ROW, columns.time, timeValues.length, 1).setNumberFormat('HH:mm');
  sheet.getRange(CONFIG.OUTPUT_START_ROW, columns.duration, durationValues.length, 1).setNumberFormat('0.##');
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {GoogleAppsScript.Spreadsheet.Range} checkboxRange
 * @param {{id:number,event:number,date:number,time:number,duration:number,upload:number,status:number}} columns
 */
function handleUploadCheckboxEdit_(sheet, row, checkboxRange, columns) {
  const statusCell = sheet.getRange(row, columns.status);

  try {
    const resultMessage = upsertOrDeleteCalendarEventFromRow_(sheet, row, columns);
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
 * @param {{id:number,event:number,date:number,time:number,duration:number,upload:number,status:number}} columns
 * @returns {string}
 */
function upsertOrDeleteCalendarEventFromRow_(sheet, row, columns) {
  const calendar = CalendarApp.getDefaultCalendar();
  const idValue = String(sheet.getRange(row, columns.id).getValue() || '').trim();
  const title = String(sheet.getRange(row, columns.event).getValue() || '').trim();

  if (!title) {
    if (!idValue) {
      return 'Skipped: empty title and no event ID to delete.';
    }

    const existing = calendar.getEventById(idValue);
    if (!existing) {
      sheet.getRange(row, columns.id).clearContent();
      return 'Skipped: event ID not found; nothing deleted.';
    }

    existing.deleteEvent();
    [columns.id, columns.event, columns.date, columns.time, columns.duration].forEach((column) => {
      sheet.getRange(row, column).clearContent();
    });
    return 'Deleted event (empty title).';
  }

  const dateValue = parseSheetDate_(sheet.getRange(row, columns.date).getValue());
  if (!dateValue) {
    throw new Error(`Row ${row}: date must be a valid date.`);
  }

  const timeValue = parseSheetTime_(sheet.getRange(row, columns.time).getValue());
  if (!timeValue) {
    throw new Error(`Row ${row}: time must be a valid time.`);
  }

  const durationHours = Number(sheet.getRange(row, columns.duration).getValue());
  if (!Number.isFinite(durationHours) || durationHours <= 0) {
    throw new Error(`Row ${row}: duration must be a positive number of hours.`);
  }

  const start = combineDateAndTime_(dateValue, timeValue);
  const end = new Date(start.getTime() + durationHours * 60 * 60 * 1000);

  let event = idValue ? calendar.getEventById(idValue) : null;
  if (event) {
    event.setTitle(title);
    event.setTime(start, end);
    sheet.getRange(row, columns.id).setValue(event.getId());
    return 'Updated existing event.';
  }

  event = calendar.createEvent(title, start, end);
  sheet.getRange(row, columns.id).setValue(event.getId());
  return 'Created new event.';
}

/**
 * Resolves required columns by header labels in CONFIG.HEADER_ROW.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{id:number,event:number,date:number,time:number,duration:number,upload:number,status:number}}
 */
function getColumnIndexes_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet
    .getRange(CONFIG.HEADER_ROW, 1, 1, Math.max(lastColumn, 1))
    .getValues()[0]
    .map((value) => String(value || '').trim().toLowerCase());

  const getRequiredColumn = (name) => {
    const index = headers.indexOf(name.toLowerCase());
    if (index === -1) {
      throw new Error(`Missing required header: "${name}".`);
    }
    return index + 1;
  };

  const upload = getRequiredColumn(CONFIG.HEADER_NAMES.upload);
  const statusIndex = headers.indexOf(CONFIG.HEADER_NAMES.status.toLowerCase());

  return {
    id: getRequiredColumn(CONFIG.HEADER_NAMES.id),
    event: getRequiredColumn(CONFIG.HEADER_NAMES.event),
    date: getRequiredColumn(CONFIG.HEADER_NAMES.date),
    time: getRequiredColumn(CONFIG.HEADER_NAMES.time),
    duration: getRequiredColumn(CONFIG.HEADER_NAMES.duration),
    upload,
    status: statusIndex === -1 ? upload + 1 : statusIndex + 1
  };
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
