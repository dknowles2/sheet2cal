const PROPS = PropertiesService.getDocumentProperties();

// ── Card builder ──────────────────────────────────────────────────────────────

function buildHomepage() {
  const sheetName  = PROPS.getProperty("sheetName")  || "";
  const calendarId = PROPS.getProperty("calendarId") || "";
  const lastSync   = PROPS.getProperty("lastSync")   || "Never";

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Sheet2Cal").setSubtitle("Sync a sheet to Google Calendar"))
    .addSection(buildConfigSection(sheetName, calendarId))
    .addSection(buildSyncSection(lastSync))
    .addSection(buildCalendarPickerSection());

  return card.build();
}

function buildConfigSection(sheetName, calendarId) {
  const sheetInput = CardService.newTextInput()
    .setFieldName("sheetName")
    .setTitle("Sheet name")
    .setHint("The tab name to sync, e.g. \"Sheet1\"")
    .setValue(sheetName);

  const calInput = CardService.newTextInput()
    .setFieldName("calendarId")
    .setTitle("Calendar ID")
    .setHint("\"primary\" or a calendar's email address")
    .setValue(calendarId);

  const saveButton = CardService.newTextButton()
    .setText("Save settings")
    .setOnClickAction(CardService.newAction().setFunctionName("saveConfig"))
    .setTextButtonStyle(CardService.TextButtonStyle.OUTLINED);

  return CardService.newCardSection()
    .setHeader("Configuration")
    .addWidget(sheetInput)
    .addWidget(calInput)
    .addWidget(saveButton);
}

function buildSyncSection(lastSync) {
  const syncButton = CardService.newTextButton()
    .setText("Sync now")
    .setOnClickAction(CardService.newAction().setFunctionName("runSyncFromCard"))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  const lastSyncText = CardService.newDecoratedText()
    .setTopLabel("Last sync")
    .setText(lastSync);

  return CardService.newCardSection()
    .setHeader("Sync")
    .addWidget(syncButton)
    .addWidget(lastSyncText);
}

function buildCalendarPickerSection() {
  const calendars = CalendarApp.getAllCalendars();
  const rows = calendars.map(cal =>
    `${cal.getName()}  —  ${cal.getId()}`
  ).join("\n");

  const calList = CardService.newTextParagraph()
    .setText(rows || "No calendars found.");

  return CardService.newCardSection()
    .setHeader("Your calendars")
    .setCollapsible(true)
    .addWidget(calList);
}

// ── Actions ───────────────────────────────────────────────────────────────────

function saveConfig(e) {
  const sheetName  = (e.formInput.sheetName  || "").trim();
  const calendarId = (e.formInput.calendarId || "").trim();

  if (!sheetName)  return notify("Sheet name is required.");
  if (!calendarId) return notify("Calendar ID is required.");

  PROPS.setProperties({ sheetName, calendarId });

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Settings saved."))
    .setNavigation(CardService.newNavigation().updateCard(buildHomepage()))
    .build();
}

function runSyncFromCard() {
  try {
    syncSheetToCalendar();
    const ts = new Date().toLocaleString();
    PROPS.setProperty("lastSync", ts);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Sync complete — ${ts}`))
      .setNavigation(CardService.newNavigation().updateCard(buildHomepage()))
      .build();
  } catch (err) {
    return notify(`Error: ${err.message}`);
  }
}

function notify(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .build();
}
