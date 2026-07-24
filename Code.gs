/**
 * J3 Creation - Lead Form Backend
 * Handles GET and POST requests from the lead capture form and
 * appends every submission as a new row into a "Leads" sheet.
 *
 * Deploy as Web App:
 *   Deploy > New deployment > Type: Web app
 *   Execute as: Me
 *   Who has access: Anyone
 */

const SHEET_NAME = 'Leads';
const HEADERS = [
  'Timestamp',
  'Full Name',
  'Phone Number',
  'Location',
  'Service Needed',
  'Message'
];

/**
 * Returns the "Leads" sheet, creating it (and its bold header row)
 * if it does not already exist.
 */
function getOrCreateLeadsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  ensureHeaders_(sheet);

  return sheet;
}

/**
 * Ensures the header row exists and is bold. If the first row is empty
 * or does not match the expected headers, it (re)writes them.
 */
function ensureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = firstRow.some(function (cell) {
    return cell !== '' && cell !== null;
  });

  if (!hasHeaders) {
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    for (let i = 1; i <= HEADERS.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
}

/**
 * Builds a standard JSON response.
 */
function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Extracts and normalizes lead fields from a request's parameters,
 * supporting both e.parameter (GET query string / form-urlencoded)
 * and a JSON body (POST with JSON payload).
 */
function extractLeadData_(e) {
  let data = {};

  if (e && e.parameter) {
    data = e.parameter;
  }

  if (e && e.postData && e.postData.type === 'application/json' && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      data = Object.assign({}, data, body);
    } catch (err) {
      // Ignore malformed JSON, fall back to query params if any.
    }
  }

  return {
    name: data.name || '',
    phone: data.phone || '',
    location: data.location || '',
    service: data.service || '',
    message: data.message || ''
  };
}

/**
 * Appends a lead row to the sheet and returns the JSON response.
 */
function saveLead_(e) {
  try {
    const sheet = getOrCreateLeadsSheet_();
    const lead = extractLeadData_(e);

    sheet.appendRow([
      new Date(),
      lead.name,
      lead.phone,
      lead.location,
      lead.service,
      lead.message
    ]);

    return jsonResponse_({
      success: true,
      message: 'Lead Saved'
    });
  } catch (error) {
    return jsonResponse_({
      success: false,
      message: 'Error: ' + error.message
    });
  }
}

/**
 * Handles GET requests (used by the front-end via a no-cors GET fetch
 * with query-string parameters).
 */
function doGet(e) {
  // Allow simple health-check pings with no params.
  if (!e || !e.parameter || Object.keys(e.parameter).length === 0) {
    return jsonResponse_({
      success: true,
      message: 'J3 Creation Lead API is live'
    });
  }
  return saveLead_(e);
}

/**
 * Handles POST requests (JSON body or form-urlencoded).
 */
function doPost(e) {
  return saveLead_(e);
}
