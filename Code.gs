// ── Shared index builder ──────────────────────────────────────────────────────

function buildIdx(headers) {
  return {
    date:      colIdx(headers, "col_date"),
    title:     colIdx(headers, "col_title"),
    startTime: colIdx(headers, "col_startTime"),
    endTime:   colIdx(headers, "col_endTime"),
    location:  colIdx(headers, "col_location"),
    notes:     colIdx(headers, "col_notes"),
    eventId:   colIdx(headers, "col_eventId"),
  };
}

// ── Calendar primitives ───────────────────────────────────────────────────────

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

// Returns "updated"
function updateEvent(calendar, sheet, sheetRow, eventId, date, title, isAllDay, row, idx, options) {
  let event;
  try { event = calendar.getEventById(eventId); } catch (e) { event = null; }

  if (!event) {
    // Event was deleted from calendar since preview ran — recreate it
    if (idx.eventId >= 0) sheet.getRange(sheetRow, idx.eventId + 1).setValue("");
    createEvent(calendar, sheet, sheetRow, date, title, isAllDay, row, idx, options);
    return "created";
  }

  event.setTitle(title);
  if (options.location    !== undefined) event.setLocation(options.location);
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

// ── Date/time helpers ─────────────────────────────────────────────────────────

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
