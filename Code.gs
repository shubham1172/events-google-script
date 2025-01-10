const ANNIVERSARIES_SHEET_NAME = "Anniversaries";
const BIRTHDAYS_SHEET_NAME = "Birthdays";
const CALENDAR_ID = "primary"; // This cannot be changed as birthdays can only be synced to the primary calendar.
const SORT_DATA = true; // If true, the data in the sheets will be sorted by the first column.

/**
 * Syncs all events to the calendar.
 * This is the main function that should be called from the trigger.
 */
function syncEvents() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    syncEventsFromSheet(ss, BIRTHDAYS_SHEET_NAME)
    syncEventsFromSheet(ss, ANNIVERSARIES_SHEET_NAME, isAnniversary = true)
}

/**
 * Syncs events from the specified sheet to the calendar. 
 */
function syncEventsFromSheet(spreadsheet, sheetName, isAnniversary = false) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        Logger.log(`INFO: Sheet ${sheetName} not found, skipping.`);
        return;
    }

    if (SORT_DATA) {
        Logger.log(`INFO: Sorting data in ${sheetName}.`);
        sheet.getDataRange().sort(1);
    }

    const range = sheet.getDataRange()
    const data = range.getValues();
    const events = data.map(row => ({ name: row[0], date: row[1] }));

    const existingEvents = listBirthdays();
    Logger.log('DEBUG: Total events in calendar: ' + Object.keys(existingEvents).length);

    Logger.log(`INFO: Syncing ${events.length} events from ${sheetName}.`);
    createEventsIfNotExists(events, existingEvents, isAnniversary);
    Logger.log('INFO: Sync events completed.');
}

/**
 * Creates events for names which do not have an existing event.
 * For simplicity, this never overwrites existing events.
 * 
 * events: Array of objects with name and date properties.
 * existingEvents: Map of event names to event IDs.
 * isAnniversary: If true, the event type is set to 'anniversary'.
 */
function createEventsIfNotExists(events, existingEvents, isAnniversary) {
    events.forEach(event => {
        // Use the appropriate event name based on the event type.
        const eventName = isAnniversary ?
            getAnniversaryEventName(event.name) :
            getBirthdayEventName(event.name);

        // If the event does not exist, create it.
        // Note, anniversary events are also birthday events with a special property to differentiate them,
        // but Google API does not support providing this property as of today.
        if (!existingEvents[eventName]) {
            createBirthday(eventName, event.date);
        } else {
            Logger.log(`DEBUG: ${eventName} already exists, skipping.`);
        }
    });
}

/** 
 * Creates a birthday event. 
 */
function createBirthday(eventName, date) {
    const [startDate, endDate] = getEventDates(date);
    const event = {
        start: { date: startDate },
        end: { date: endDate },
        eventType: 'birthday',
        recurrence: ["RRULE:FREQ=YEARLY"],
        summary: eventName,
        transparency: "transparent",
        visibility: "private",
        reminders: {
            useDefault: false,
            overrides: [
                { method: 'popup', minutes: 60 * 24 * 1 }, // 1 days before
                { method: 'popup', minutes: 60 * 24 * 7 }, // 7 days before
            ]
        }
    }

    createEvent(event);
}

/**
 * Returns the event name based on the input.
 * E.g., "John" -> "John's birthday"
 */
function getBirthdayEventName(personName) {
    return personName + "'s birthday";
}

/**
 * Returns the event name based on the input.
 * E.g., "John and Jane" -> "John and Jane's anniversary"
 */
function getAnniversaryEventName(coupleName) {
    return coupleName + "'s anniversary";
}

// ******************************
// Date functions
// ******************************

/**
 * Returns the start and end dates of an event based on the input.
 * E.g., 
 * "12/05" -> ["2025-12-05", "2025-12-06"],
 * "31/10" -> ["2025-10-31", "2025-11-01"]
 */
function getEventDates(input) {
    const currentYear = new Date().getFullYear();
    const [day, month] = input.split('/').map(Number);

    if (!day || !month || day > 31 || month > 12) {
        throw new Error("Invalid input format. Use DD/MM with valid values.");
    }

    const firstDate = new Date(currentYear, month - 1, day); // JS months are 0-indexed
    const secondDate = new Date(currentYear, month - 1, day + 1); // Adds 1 day

    const formatDate = (date) =>
        `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;

    return [formatDate(firstDate), formatDate(secondDate)];
}

/**
 * Returns the start and end dates of the current year.
 * E.g.,
 * { start: "2025-01-01T00:00:00+00:00", end: "2025-12-31T23:59:59+00:00"}
 */
function getYearStartAndEnd() {
    const currentYear = new Date().getFullYear();
    return [`${currentYear}-01-01T00:00:00+00:00`, `${currentYear}-12-31T23:59:59+00:00`];
}

// ******************************
// Calendar API functions
// ******************************

/**
  * Creates a Calendar event.
  * See https://developers.google.com/calendar/api/v3/reference/events/insert
  */
function createEvent(event) {
    try {
        var response = Calendar.Events.insert(event, CALENDAR_ID);
        Logger.log('INFO: Created event for ' + event.summary + ', htmlLink: ' + response.htmlLink);
    } catch (exception) {
        Logger.log('ERROR: Error creating event: ' + exception);
    }
}


/**
  * Lists all birthday events in current year.
  * See https://developers.google.com/calendar/api/v3/reference/events/list
  * 
  * This also includes anniversary events (as they are birthday events).
  * 
  * It returns a map of event names to event IDs.
  * e.g., { "John's birthday": "1234", "John and Jane's anniversary": "5678" }
  */
function listBirthdays() {
    const [start, end] = getYearStartAndEnd();

    const optionalArgs = {
        eventTypes: ['birthday'],
        singleEvents: true,
        timeMax: end,
        timeMin: start,
    };

    const birthdays = {};

    try {
        var response = Calendar.Events.list(CALENDAR_ID, optionalArgs);
        response.items.forEach(event => {
            birthdays[event.summary] = event.id;
        });
    } catch (exception) {
        Logger.log('ERROR: Error listing events: ' + exception);
    }

    return birthdays;
}