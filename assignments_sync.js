const { google } = require('googleapis');
const { isEqual } = require('lodash');
const crypto = require('crypto');
const moment = require('moment-timezone');

const { CALENDAR_ID, SPREADSHEET_ID, SEMESTER_ID, TIMEZONE } = require('./config.json')

let logMessages = [];

function getLogMessageFromLogCall(args) {
  const timestamp = moment().tz(TIMEZONE).format('M/D/YYYY H:mm:ss z');
  const message = args.map(arg => 
    typeof arg === 'object' ? JSON.stringify(arg) : String(arg)
  ).join(' ');

  return [timestamp, message];
}

function logAndSendToSheet(...args) {
  logMessages.push(getLogMessageFromLogCall(args));
  console.log(...args);
};

function logError(...args) {
  const [timestamp, message] = getLogMessageFromLogCall(args);
  logMessages.push([timestamp, `ERROR: ${message}`]);
  console.error(...args);
};

async function appendToLog(auth) {
  if (logMessages.length === 0) return;

  const sheets = google.sheets({ version: 'v4', auth });
  
  // Append logs to the Log sheet
  await retryWithBackoff(() =>
    sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Log!A:B',
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: logMessages
      }
    })
  );
}

const SCOPES = [
  'https://www.googleapis.com/auth/calendar',
  'https://www.googleapis.com/auth/spreadsheets'
];
const RANGE = `${SEMESTER_ID}!A2:I`;
const LAST_UPDATED_CELL = `${SEMESTER_ID}!K1`;

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
        logAndSendToSheet(`Rate limited. Attempt ${attempt + 1}/${maxRetries}. Retrying in ${delay}ms`);
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

// Fill empty columns with ""
async function getNormalizedSpreadsheet(sheets) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: RANGE,
  });

  response.data.values = response.data.values ?? [];

  for (let i = 0; i < response.data.values.length; i++) {
    if (response.data.values[i].length !== 9) {
      response.data.values[i].push(...(new Array(9-response.data.values[i].length).fill("")))
    }
  }

  return response
}

async function throwIfSpreadsheetChanged(originalResponse, sheets) {
  const checkResponse = await retryWithBackoff(() =>
    getNormalizedSpreadsheet(sheets)
  );

  if (!isEqual(originalResponse, checkResponse.data.values)) {
    throw new Error('Spreadsheet required updating but content changed during sync!');
  }
}

async function getSpreadsheetData(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const response = await retryWithBackoff(() =>
    getNormalizedSpreadsheet(sheets)
  );

  const originalValues = response.data.values;

  const data = [];
  const uuidsToBeUpdated = [];

  for (let i = 0; i < originalValues.length; i++) {
    const row = originalValues[i];
    
    const uuid = row[8] || crypto.randomUUID();

    // If UUID was generated, update spreadsheet
    if (!row[8]) {
      await throwIfSpreadsheetChanged(originalValues, sheets);

      // Add one to index for one-based indexing, then another one to skip the header
      uuidsToBeUpdated.push({ rowIndex: i + 2, uuid: uuid })
    }

    data.push({
      Origin: row[0] || "Unknown",
      Name: row[1] || "Unknown",
      'Due Time': row[2] || "Unknown",
      'Due Date': row[3] || "Unknown",
      Status: row[4] || "Unknown",
      Difficulty: row[5] || "Unknown",
      Priority: row[6] || "Unknown",
      Notes: row[7] || "Unknown",
      UUID: uuid
    });
  }

  const newValues = originalValues;

  for (const { rowIndex, uuid } of uuidsToBeUpdated) {
    await throwIfSpreadsheetChanged(originalValues, sheets);

    await retryWithBackoff(() =>
      sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SEMESTER_ID}!I${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: {
          values: [[uuid]]
        }
      })
    );

    newValues[rowIndex-2][8] = uuid;
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
  const summary = `${assignment.Status === '2 - Done' ? "DONE - " : ""}${assignment.Origin}: ${assignment.Name}`

  return {
    summary: summary,
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
          logAndSendToSheet('Updated event:', eventData.summary);
        }
        existingEventMap.delete(assignment.UUID);
      } else {
        await retryWithBackoff(() =>
          calendar.events.insert({
            calendarId: CALENDAR_ID,
            requestBody: eventData,
          })
        );
        logAndSendToSheet('Created event:', eventData.summary);
      }
    }

    for (const [_, event] of existingEventMap) {
      await retryWithBackoff(() =>
        calendar.events.delete({
          calendarId: CALENDAR_ID,
          eventId: event.id,
        })
      );
      logAndSendToSheet('Deleted event:', event.summary);
    }

    // Update the last sync time
    await updateLastSyncTime(sheets);

    console.log(`Sync complete at ${moment().toLocaleString()}!`)

    // Append all collected logs before terminating
    await appendToLog(auth);
  } catch (error) {
    logError('Error:', error.response?.data || error);

    // Try to log the error even if sync failed
    try {
      const auth = await authorize();
      await appendToLog(auth);
    } catch (logError) {
      console.error('Failed to write to log:', logError);
    }
  }
}

syncWithCalendar();