const { google } = require('googleapis');
const { isEqual } = require('lodash');
const crypto = require('crypto');
const moment = require('moment-timezone');

const { CALENDAR_ID, SPREADSHEET_ID, TIMEZONE } = require('./config.json')

const SCOPES = [
  'https://www.googleapis.com/auth/calendar',
  'https://www.googleapis.com/auth/spreadsheets'
];
const RANGE = 'WN25!A2:I';
const LAST_UPDATED_CELL = 'WN25!K1';

async function wait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function retryWithBackoff(operation, maxRetries = 5, baseDelay = 1000) {
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return await operation();
    } catch (error) {
      if (error.response?.status === 429) {
        const delay = baseDelay * Math.pow(2, attempt);
        const jitter = Math.random() * 1000;
        console.log(`Rate limited. Attempt ${attempt + 1}/${maxRetries}. Retrying in ${delay}ms`);
        await wait(delay + jitter);
        continue;
      }
      throw error;
    }
  }

  throw new Error("Exceeded max retries.");
}

async function authorize() {
  const auth = new google.auth.GoogleAuth({
    keyFile: 'credentials.json',
    scopes: SCOPES,
  });
  return auth.getClient();
}

async function updateLastSyncTime(sheets) {
  const timestamp = moment().tz(TIMEZONE).format('M/D/YYYY H:mm:ss z');

  await retryWithBackoff(() =>
    sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: LAST_UPDATED_CELL,
      valueInputOption: 'RAW',
      requestBody: {
        values: [[timestamp]]
      }
    })
  );

  console.log(`Updated last sync time to ${timestamp}`);
}

async function throwIfSpreadsheetChanged(originalData) {
  const checkResponse = await retryWithBackoff(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: RANGE,
    })
  );

  if (!isEqual(originalData, checkResponse.data.values)) {
    throw new Error('Spreadsheet required updating but content changed during sync!');
  }
}

async function getSpreadsheetData(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const response = await retryWithBackoff(() =>
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: RANGE,
    })
  );

  const data = [];
  const uuidsToBeUpdated = [];

  for (const [index, row] of Object.entries(response.data.values)) {
    const uuid = row[8] || crypto.randomUUID();

    // If UUID was generated, update spreadsheet
    if (!row[8]) {
      await throwIfSpreadsheetChanged(response.data.values);

      // Add one to index for one-based indexing, then another one to skip the header
      uuidsToBeUpdated.push({ rowIndex: index + 2, uuid: uuid })
    }

    data.push({
      Origin: row[0] ?? "Unknown",
      Name: row[1] ?? "Unknown",
      'Due Date': row[2] ?? "Unknown",
      'Due Time': row[3] ?? "Unknown",
      Status: row[4] ?? "Unknown",
      Difficulty: row[5] ?? "Unknown",
      Priority: row[6] ?? "Unknown",
      Notes: row[7] ?? "Unknown",
      UUID: uuid
    });
  }

  for (const { rowIndex, uuid } of uuidsToBeUpdated) {
    await throwIfSpreadsheetChanged(response.data.values);

    await retryWithBackoff(() =>
      sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `WN25!I${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: {
          values: [[uuid]]
        }
      })
    );
  }

  return { data, sheets };
}

function getEventData(assignment) {
  const originalDateTimeString = `${assignment['Due Date']}|${assignment['Due Time']}`;

  const dueDateAndTime = moment.tz(originalDateTimeString, 'L|LTS', TIMEZONE);

  if (!dueDateAndTime.isValid()) {
    throw new Error(`Moment could not parse time "${originalDateTimeString}"`)
  }

  const eventDateTime = dueDateAndTime.toISOString(true);

  return {
    summary: `${assignment.Origin}: ${assignment.Name}`,
    description: `Difficulty: ${assignment.Difficulty}\nPriority: ${assignment.Priority}\nNotes: ${assignment.Notes}`,
    start: {
      dateTime: eventDateTime,
      timeZone: TIMEZONE
    },
    end: {
      dateTime: eventDateTime,
      timeZone: TIMEZONE
    },
    extendedProperties: {
      private: { uuid: assignment.UUID }
    }
  };
}

function eventsAreEqual(event1, event2) {
  const normalizeEvent = (event) => ({
    summary: event.summary,
    description: event.description,
    start: moment.tz(event.start.dateTime, event.start.timeZone).unix(),
    end: moment.tz(event.end.dateTime, event.end.timeZone).unix()
  });

  const a = normalizeEvent(event1);
  const b = normalizeEvent(event2);

  return isEqual(a, b);
}

async function syncWithCalendar() {
  try {
    console.log(`Syncing calendar at ${moment().toLocaleString()}...`)

    const auth = await authorize();
    const { data: assignments, sheets } = await getSpreadsheetData(auth);
    const calendar = google.calendar({ version: 'v3', auth });

    const existingEvents = await retryWithBackoff(() =>
      calendar.events.list({
        calendarId: CALENDAR_ID,
      })
    );

    const existingEventMap = new Map(
      existingEvents.data.items.map(event => [event.extendedProperties?.private?.uuid, event])
    );

    for (const assignment of assignments) {
      const eventData = getEventData(assignment);
      const existingEvent = existingEventMap.get(assignment.UUID);

      if (existingEvent) {
        if (!eventsAreEqual(existingEvent, eventData)) {
          await retryWithBackoff(() =>
            calendar.events.update({
              calendarId: CALENDAR_ID,
              eventId: existingEvent.id,
              requestBody: eventData,
            })
          );
          console.log('Updated event:', eventData.summary);
        }
        existingEventMap.delete(assignment.UUID);
      } else {
        await retryWithBackoff(() =>
          calendar.events.insert({
            calendarId: CALENDAR_ID,
            requestBody: eventData,
          })
        );
        console.log('Created event:', eventData.summary);
      }
    }

    for (const [_, event] of existingEventMap) {
      await retryWithBackoff(() =>
        calendar.events.delete({
          calendarId: CALENDAR_ID,
          eventId: event.id,
        })
      );
      console.log('Deleted event:', event.summary);
    }

    // Update the last sync time
    await updateLastSyncTime(sheets);

    console.log(`Sync complete at ${moment().toLocaleString()}!`)
  } catch (error) {
    console.error('Error:', error.response?.data || error);
  }
}

syncWithCalendar();