// Column indices (1-based)
const COL_DATE       = 1;
const COL_START_TIME = 2;
const COL_END_TIME   = 3;
const COL_TITLE      = 4;
const COL_LOCATION   = 5;
const COL_NOTES      = 6;
const COL_EVENT_ID   = 7;

function syncSheetToCalendar() {
  const sheetName  = PROPS.getProperty("sheetName")  || "Sheet1";
  const calendarId = PROPS.getProperty("calendarId") || "primary";

  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const calendar = CalendarApp.getCalendarById(calendarId);

  if (!sheet)    throw new Error(`Sheet "${sheetName}" not found`);
  if (!calendar) throw new Error(`Calendar "${calendarId}" not found or not accessible`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // nothing but a header

  const range = sheet.getRange(2, 1, lastRow - 1, COL_EVENT_ID);
  const rows  = range.getValues();

  rows.forEach((row, i) => {
    const sheetRow = i + 2; // 1-based, offset for header

    const dateVal  = row[COL_DATE - 1];
    const title    = String(row[COL_TITLE - 1]).trim();
    const eventId  = String(row[COL_EVENT_ID - 1]).trim();

    // Skip rows without a date or title
    if (!dateVal || !title) {
      if (eventId) deleteEvent(calendar, sheet, sheetRow, eventId);
      return;
    }

    const date     = new Date(dateVal);
    const isAllDay = !row[COL_START_TIME - 1];
    const location = String(row[COL_LOCATION - 1] || "").trim();
    const notes    = String(row[COL_NOTES - 1]    || "").trim();

    const options = {};
    if (location) options.location    = location;
    if (notes)    options.description = notes;

    if (eventId) {
      // Update existing event
      updateEvent(calendar, sheet, sheetRow, eventId, date, title, isAllDay, row, options);
    } else {
      // Create new event
      createEvent(calendar, sheet, sheetRow, date, title, isAllDay, row, options);
    }
  });
}

function createEvent(calendar, sheet, sheetRow, date, title, isAllDay, row, options) {
  let event;
  if (isAllDay) {
    event = calendar.createAllDayEvent(title, date, options);
  } else {
    const { start, end } = buildDateTimes(date, row);
    event = calendar.createEvent(title, start, end, options);
  }
  sheet.getRange(sheetRow, COL_EVENT_ID).setValue(event.getId());
}

function updateEvent(calendar, sheet, sheetRow, eventId, date, title, isAllDay, row, options) {
  let event;
  try {
    event = calendar.getEventById(eventId);
  } catch (e) {
    event = null;
  }

  if (!event) {
    // Event was deleted from the calendar — recreate it
    sheet.getRange(sheetRow, COL_EVENT_ID).setValue("");
    createEvent(calendar, sheet, sheetRow, date, title, isAllDay, row, options);
    return;
  }

  event.setTitle(title);
  if (options.location)    event.setLocation(options.location);
  if (options.description) event.setDescription(options.description);

  if (isAllDay) {
    event.setAllDayDate(date);
  } else {
    const { start, end } = buildDateTimes(date, row);
    event.setTime(start, end);
  }
}

function deleteEvent(calendar, sheet, sheetRow, eventId) {
  try {
    const event = calendar.getEventById(eventId);
    if (event) event.deleteEvent();
  } catch (e) {
    // Already gone — that's fine
  }
  sheet.getRange(sheetRow, COL_EVENT_ID).setValue("");
}

function buildDateTimes(date, row) {
  const startTimeVal = row[COL_START_TIME - 1];
  const endTimeVal   = row[COL_END_TIME - 1];

  const start = combineDateAndTime(date, startTimeVal);
  // Default end to 1 hour after start if missing
  const end   = endTimeVal ? combineDateAndTime(date, endTimeVal) : new Date(start.getTime() + 60 * 60 * 1000);

  return { start, end };
}

function combineDateAndTime(date, timeVal) {
  const base = new Date(date);
  if (!timeVal) return base;

  // timeVal from Sheets is a Date object where only the time portion matters
  if (timeVal instanceof Date) {
    base.setHours(timeVal.getHours(), timeVal.getMinutes(), timeVal.getSeconds(), 0);
  } else {
    // Handle string like "14:30" or "2:30 PM"
    const t = new Date(`1970/01/01 ${timeVal}`);
    if (!isNaN(t)) base.setHours(t.getHours(), t.getMinutes(), 0, 0);
  }
  return base;
}

