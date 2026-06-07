// Returns an array of result objects: { action, label, detail }
// action: "created" | "updated" | "deleted" | "skipped" | "failed"
function syncSheetToCalendar() {
  const sheetName  = PROPS.getProperty("sheetName")  || "Sheet1";
  const calendarId = PROPS.getProperty("calendarId") || "primary";

  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const calendar = CalendarApp.getCalendarById(calendarId);

  if (!sheet)    throw new Error(`Sheet "${sheetName}" not found`);
  if (!calendar) throw new Error(`Calendar "${calendarId}" not found or not accessible`);

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers   = allValues[0].map(h => String(h).trim());

  const idx = {
    date:      colIdx(headers, "col_date"),
    title:     colIdx(headers, "col_title"),
    startTime: colIdx(headers, "col_startTime"),
    endTime:   colIdx(headers, "col_endTime"),
    location:  colIdx(headers, "col_location"),
    notes:     colIdx(headers, "col_notes"),
    eventId:   colIdx(headers, "col_eventId"),
  };

  if (idx.date  < 0) throw new Error("Date column is not mapped. Open the Sheet2Cal sidebar to configure.");
  if (idx.title < 0) throw new Error("Event Title column is not mapped. Open the Sheet2Cal sidebar to configure.");

  const results = [];

  allValues.slice(1).forEach((row, i) => {
    const sheetRow = i + 2;
    const dateVal  = idx.date  >= 0 ? row[idx.date]  : null;
    const title    = idx.title >= 0 ? String(row[idx.title]).trim() : "";
    const eventId  = idx.eventId >= 0 ? String(row[idx.eventId]).trim() : "";

    if (!dateVal && !title) return; // silently skip fully empty rows

    if (!dateVal || !title) {
      const missing = [!dateVal && "date", !title && "title"].filter(Boolean).join(" & ");
      results.push({ action: "skipped", label: `Row ${sheetRow}`, detail: `Missing ${missing}` });
      if (eventId) {
        try {
          deleteEvent(calendar, sheet, sheetRow, idx.eventId);
        } catch (e) {
          results[results.length - 1].action = "failed";
          results[results.length - 1].detail += ` (delete failed: ${e.message})`;
        }
      }
      return;
    }

    const date     = new Date(dateVal);
    const dateStr  = date.toLocaleDateString();
    const isAllDay = !(idx.startTime >= 0 && row[idx.startTime]);
    const location = idx.location >= 0 ? String(row[idx.location] || "").trim() : "";
    const notes    = idx.notes    >= 0 ? String(row[idx.notes]    || "").trim() : "";

    const options = {};
    if (location) options.location    = location;
    if (notes)    options.description = notes;

    try {
      if (eventId) {
        const outcome = updateEvent(calendar, sheet, sheetRow, eventId, date, title, isAllDay, row, idx, options);
        results.push({ action: outcome, label: title, detail: dateStr });
      } else {
        createEvent(calendar, sheet, sheetRow, date, title, isAllDay, row, idx, options);
        results.push({ action: "created", label: title, detail: dateStr });
      }
    } catch (e) {
      results.push({ action: "failed", label: title, detail: `${dateStr} — ${e.message}` });
    }
  });

  return results;
}

// Returns "updated" or "unchanged"
function createEvent(calendar, sheet, sheetRow, date, title, isAllDay, row, idx, options) {
  let event;
  if (isAllDay) {
    event = calendar.createAllDayEvent(title, date, options);
  } else {
    const { start, end } = buildDateTimes(date, row, idx);
    event = calendar.createEvent(title, start, end, options);
  }
  if (idx.eventId >= 0) {
    sheet.getRange(sheetRow, idx.eventId + 1).setValue(event.getId());
  }
}

function updateEvent(calendar, sheet, sheetRow, eventId, date, title, isAllDay, row, idx, options) {
  let event;
  try { event = calendar.getEventById(eventId); } catch (e) { event = null; }

  if (!event) {
    if (idx.eventId >= 0) sheet.getRange(sheetRow, idx.eventId + 1).setValue("");
    createEvent(calendar, sheet, sheetRow, date, title, isAllDay, row, idx, options);
    return "created";
  }

  event.setTitle(title);
  if (options.location !== undefined) event.setLocation(options.location);
  if (options.description !== undefined) event.setDescription(options.description);

  if (isAllDay) {
    event.setAllDayDate(date);
  } else {
    const { start, end } = buildDateTimes(date, row, idx);
    event.setTime(start, end);
  }
  return "updated";
}

function deleteEvent(calendar, sheet, sheetRow, eventIdColIdx) {
  const eventId = String(sheet.getRange(sheetRow, eventIdColIdx + 1).getValue()).trim();
  try {
    const event = calendar.getEventById(eventId);
    if (event) event.deleteEvent();
  } catch (e) { /* already gone */ }
  sheet.getRange(sheetRow, eventIdColIdx + 1).setValue("");
}

function buildDateTimes(date, row, idx) {
  const startVal = idx.startTime >= 0 ? row[idx.startTime] : null;
  const endVal   = idx.endTime   >= 0 ? row[idx.endTime]   : null;

  const start = combineDateAndTime(date, startVal);
  const end   = endVal
    ? combineDateAndTime(date, endVal)
    : new Date(start.getTime() + 60 * 60 * 1000);

  return { start, end };
}

function combineDateAndTime(date, timeVal) {
  const base = new Date(date);
  if (!timeVal) return base;
  if (timeVal instanceof Date) {
    base.setHours(timeVal.getHours(), timeVal.getMinutes(), timeVal.getSeconds(), 0);
  } else {
    const t = new Date(`1970/01/01 ${timeVal}`);
    if (!isNaN(t)) base.setHours(t.getHours(), t.getMinutes(), 0, 0);
  }
  return base;
}

function colIdx(headers, propKey) {
  const saved = (PROPS.getProperty(propKey) || "").trim();
  if (!saved) return -1;
  return headers.indexOf(saved);
}
