const PROPS = PropertiesService.getDocumentProperties();

const FIELD_DEFS = [
  { key: "col_date",       label: "Date",        required: true  },
  { key: "col_title",      label: "Event Title", required: true  },
  { key: "col_startTime",  label: "Start Time",  required: false },
  { key: "col_endTime",    label: "End Time",    required: false },
  { key: "col_location",   label: "Location",    required: false },
  { key: "col_notes",      label: "Notes",       required: false },
  { key: "col_eventId",    label: "Event ID",    required: false },
];

// ── Entry points ──────────────────────────────────────────────────────────────

function buildHomepage() {
  return isConfigured() ? buildMainCard() : buildSettingsCard(null);
}

// Called when the sheet dropdown changes so column headers can refresh.
function onSheetChange(e) {
  const nav = CardService.newNavigation().updateCard(buildSettingsCard(e));
  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}

function openSettings() {
  const nav = CardService.newNavigation().pushCard(buildSettingsCard(null));
  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}

// ── Main (home) card ──────────────────────────────────────────────────────────

function buildMainCard() {
  const sheetName    = PROPS.getProperty("sheetName")  || "";
  const calendarId   = PROPS.getProperty("calendarId") || "";
  const lastSync     = PROPS.getProperty("lastSync")   || "Never";

  const calendarName = getCalendarName(calendarId);

  const configSummary = CardService.newCardSection()
    .addWidget(
      CardService.newDecoratedText()
        .setTopLabel("Sheet")
        .setText(sheetName)
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION))
    )
    .addWidget(
      CardService.newDecoratedText()
        .setTopLabel("Calendar")
        .setText(calendarName)
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.INVITE))
    );

  const previewButton = CardService.newTextButton()
    .setText("Preview")
    .setOnClickAction(CardService.newAction().setFunctionName("runPreviewFromMain"))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  const settingsButton = CardService.newTextButton()
    .setText("Settings")
    .setOnClickAction(CardService.newAction().setFunctionName("openSettings"))
    .setTextButtonStyle(CardService.TextButtonStyle.OUTLINED);

  const buttonSet = CardService.newButtonSet()
    .addButton(settingsButton)
    .addButton(previewButton);

  const syncSection = CardService.newCardSection()
    .addWidget(buttonSet)
    .addWidget(
      CardService.newDecoratedText()
        .setTopLabel("Last sync")
        .setText(lastSync)
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CLOCK))
    );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle("Sheet2Cal")
      .setSubtitle("Sync a sheet to Google Calendar"))
    .addSection(configSummary)
    .addSection(syncSection)
    .build();
}

// ── Settings card ─────────────────────────────────────────────────────────────

function buildSettingsCard(e) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ss.getSheets().map(s => s.getName());

  const savedSheet    = PROPS.getProperty("sheetName") || sheetNames[0];
  const selectedSheet = (e && e.formInput && e.formInput.sheetName) || savedSheet;

  const headers   = getHeaders(selectedSheet);
  const calendars = CalendarApp.getAllCalendars();

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle("Settings")
      .setSubtitle("Sheet2Cal configuration"))
    .addSection(buildSheetSection(sheetNames, selectedSheet))
    .addSection(buildCalendarSection(calendars))
    .addSection(buildColumnSection(headers))
    .addSection(buildSaveSection())
    .build();
}

// ── Section builders ──────────────────────────────────────────────────────────

function buildSheetSection(sheetNames, selectedSheet) {
  const dropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName("sheetName")
    .setTitle("Sheet to sync")
    .setOnChangeAction(CardService.newAction().setFunctionName("onSheetChange"));

  sheetNames.forEach(name => dropdown.addItem(name, name, name === selectedSheet));

  return CardService.newCardSection()
    .setHeader("Sheet")
    .addWidget(dropdown);
}

function buildCalendarSection(calendars) {
  const savedCalId = PROPS.getProperty("calendarId") || "";

  const dropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName("calendarId")
    .setTitle("Target calendar");

  const primary = calendars.find(c => c.isMyPrimaryCalendar());
  if (primary) {
    dropdown.addItem(`${primary.getName()} (primary)`, primary.getId(), primary.getId() === savedCalId);
  }
  calendars
    .filter(c => !c.isMyPrimaryCalendar())
    .forEach(cal => dropdown.addItem(cal.getName(), cal.getId(), cal.getId() === savedCalId));

  return CardService.newCardSection()
    .setHeader("Calendar")
    .addWidget(dropdown);
}

function buildColumnSection(headers) {
  const section = CardService.newCardSection()
    .setHeader("Column mapping")
    .setCollapsible(false);

  FIELD_DEFS.forEach(({ key, label, required }) => {
    const saved = PROPS.getProperty(key) || "";

    const dropdown = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName(key)
      .setTitle(required ? `${label} *` : label);

    if (!required) dropdown.addItem("— not mapped —", "", saved === "");
    headers.forEach(h => dropdown.addItem(h, h, h === saved));

    section.addWidget(dropdown);
  });

  return section;
}

function buildSaveSection() {
  const saveButton = CardService.newTextButton()
    .setText("Save settings")
    .setOnClickAction(CardService.newAction().setFunctionName("saveConfig"))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  return CardService.newCardSection().addWidget(saveButton);
}

// ── Actions ───────────────────────────────────────────────────────────────────

function saveConfig(e) {
  const err = applyConfig(e);
  if (err) return notify(err);

  // Pop back to (or replace with) the main card
  const nav = CardService.newNavigation().popToRoot().updateCard(buildMainCard());
  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .setNotification(CardService.newNotification().setText("Settings saved."))
    .build();
}

// Entry point from main card's Preview button (no form input to save first)
function runPreviewFromMain() {
  let actions;
  try {
    actions = previewSync();
  } catch (ex) {
    return notify(`Error: ${ex.message}`);
  }
  const nav = CardService.newNavigation().pushCard(buildPreviewCard(actions));
  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function isConfigured() {
  return !!(PROPS.getProperty("sheetName") && PROPS.getProperty("calendarId") && PROPS.getProperty("col_date") && PROPS.getProperty("col_title"));
}

function applyConfig(e) {
  const fi = e.formInput;

  if (!fi.sheetName)  return "Sheet name is required.";
  if (!fi.calendarId) return "Calendar is required.";
  if (!fi.col_date)   return "A 'Date' column mapping is required.";
  if (!fi.col_title)  return "An 'Event Title' column mapping is required.";

  const props = { sheetName: fi.sheetName, calendarId: fi.calendarId };
  FIELD_DEFS.forEach(({ key }) => { props[key] = fi[key] || ""; });
  PROPS.setProperties(props);

  return null;
}

function getCalendarName(calendarId) {
  if (!calendarId) return "Not configured";
  try {
    const cal = CalendarApp.getCalendarById(calendarId);
    return cal ? cal.getName() : calendarId;
  } catch (e) {
    return calendarId;
  }
}

function getHeaders(sheetName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).trim())
    .filter(h => h !== "");
}

function notify(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .build();
}
