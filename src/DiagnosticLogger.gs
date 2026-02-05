/**
 * DiagnosticLogger.gs - Diagnostic Mode Logging System
 *
 * Provides comprehensive logging capabilities for debugging issues with:
 * - Agenda generation
 * - Meeting notes appending
 * - API calls (Claude, Todoist, etc.)
 * - Data collection from various sources
 *
 * All logging is conditional - only runs when DIAGNOSTIC_MODE is enabled in Script Properties.
 */

// ============================================================================
// DIAGNOSTIC MODE CHECK
// ============================================================================

/**
 * Checks if diagnostic mode is currently enabled.
 * @returns {boolean} True if diagnostic mode is on
 */
function isDiagnosticModeEnabled() {
  const mode = PropertiesService.getScriptProperties().getProperty('DIAGNOSTIC_MODE');
  return mode === 'true';
}

/**
 * Enables diagnostic mode.
 */
function enableDiagnosticMode() {
  PropertiesService.getScriptProperties().setProperty('DIAGNOSTIC_MODE', 'true');
  Logger.log('Diagnostic mode ENABLED');
}

/**
 * Disables diagnostic mode.
 */
function disableDiagnosticMode() {
  PropertiesService.getScriptProperties().setProperty('DIAGNOSTIC_MODE', 'false');
  Logger.log('Diagnostic mode DISABLED');
}

/**
 * Toggles diagnostic mode on/off and returns the new state.
 * @returns {boolean} The new diagnostic mode state
 */
function toggleDiagnosticMode() {
  const currentState = isDiagnosticModeEnabled();
  if (currentState) {
    disableDiagnosticMode();
  } else {
    enableDiagnosticMode();
  }
  return !currentState;
}

// ============================================================================
// SHEET INITIALIZATION
// ============================================================================

/**
 * Creates all diagnostic sheets with proper headers.
 * Sheets are created hidden by default.
 */
function initializeDiagnosticSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let created = 0;

  // Sheet 1: API_Request_Log
  created += createDiagnosticSheet(
    ss,
    CONFIG.SHEETS.API_REQUEST_LOG,
    [
      'Timestamp', 'Request_ID', 'API_Name', 'Endpoint', 'Method',
      'Headers', 'Payload', 'Client_ID', 'Event_ID', 'Context'
    ]
  );

  // Sheet 2: API_Response_Log
  created += createDiagnosticSheet(
    ss,
    CONFIG.SHEETS.API_RESPONSE_LOG,
    [
      'Timestamp', 'Request_ID', 'API_Name', 'Status_Code', 'Response_Headers',
      'Response_Body', 'Parse_Success', 'Extracted_Data', 'Error_Message', 'Duration_MS'
    ]
  );

  // Sheet 3: Data_Collection_Log
  created += createDiagnosticSheet(
    ss,
    CONFIG.SHEETS.DATA_COLLECTION_LOG,
    [
      'Timestamp', 'Client_ID', 'Event_ID', 'Source', 'Query_Details',
      'Items_Found', 'Sample_Data', 'Is_Empty', 'Error'
    ]
  );

  // Sheet 4: Agenda_Generation_Trace
  created += createDiagnosticSheet(
    ss,
    CONFIG.SHEETS.AGENDA_GENERATION_TRACE,
    [
      'Trace_ID', 'Timestamp', 'Event_ID', 'Event_Title', 'Client_ID',
      'Step_Number', 'Step_Name', 'Step_Status', 'Step_Details', 'Data_Summary', 'Duration_MS'
    ]
  );

  // Sheet 5: Notes_Append_Trace
  created += createDiagnosticSheet(
    ss,
    CONFIG.SHEETS.NOTES_APPEND_TRACE,
    [
      'Trace_ID', 'Timestamp', 'Client_ID', 'Message_ID', 'Doc_ID',
      'Step_Number', 'Step_Name', 'Step_Status', 'Step_Details', 'Content_Length', 'Duration_MS'
    ]
  );

  // Sheet 6: Doc_Append_Log
  created += createDiagnosticSheet(
    ss,
    CONFIG.SHEETS.DOC_APPEND_LOG,
    [
      'Timestamp', 'Client_ID', 'Doc_ID', 'Doc_URL', 'Content_Type',
      'Content_Preview', 'Content_Full_Length', 'Append_Success',
      'Error_Details', 'Before_Doc_Length', 'After_Doc_Length', 'Verification_Status'
    ]
  );

  Logger.log(`Diagnostic sheets initialized: ${created} sheet(s) created`);
  return created;
}

/**
 * Helper function to create a single diagnostic sheet.
 * @param {Spreadsheet} ss - The spreadsheet object
 * @param {string} sheetName - Name of the sheet to create
 * @param {Array<string>} headers - Column headers
 * @returns {number} 1 if created, 0 if already exists
 */
function createDiagnosticSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
    Logger.log(`Created hidden diagnostic sheet: ${sheetName}`);
    return 1;
  }

  return 0;
}

/**
 * Clears all diagnostic sheets (keeps headers).
 */
function clearDiagnosticSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const diagnosticSheets = [
    CONFIG.SHEETS.API_REQUEST_LOG,
    CONFIG.SHEETS.API_RESPONSE_LOG,
    CONFIG.SHEETS.DATA_COLLECTION_LOG,
    CONFIG.SHEETS.AGENDA_GENERATION_TRACE,
    CONFIG.SHEETS.NOTES_APPEND_TRACE,
    CONFIG.SHEETS.DOC_APPEND_LOG
  ];

  let cleared = 0;
  diagnosticSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
      cleared++;
    }
  });

  Logger.log(`Cleared ${cleared} diagnostic sheet(s)`);
  return cleared;
}

/**
 * Shows all diagnostic sheets (unhides them for viewing).
 */
function showDiagnosticSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const diagnosticSheets = [
    CONFIG.SHEETS.API_REQUEST_LOG,
    CONFIG.SHEETS.API_RESPONSE_LOG,
    CONFIG.SHEETS.DATA_COLLECTION_LOG,
    CONFIG.SHEETS.AGENDA_GENERATION_TRACE,
    CONFIG.SHEETS.NOTES_APPEND_TRACE,
    CONFIG.SHEETS.DOC_APPEND_LOG
  ];

  let shown = 0;
  diagnosticSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.isSheetHidden()) {
      sheet.showSheet();
      shown++;
    }
  });

  Logger.log(`Showed ${shown} diagnostic sheet(s)`);
  return shown;
}

/**
 * Hides all diagnostic sheets.
 */
function hideDiagnosticSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const diagnosticSheets = [
    CONFIG.SHEETS.API_REQUEST_LOG,
    CONFIG.SHEETS.API_RESPONSE_LOG,
    CONFIG.SHEETS.DATA_COLLECTION_LOG,
    CONFIG.SHEETS.AGENDA_GENERATION_TRACE,
    CONFIG.SHEETS.NOTES_APPEND_TRACE,
    CONFIG.SHEETS.DOC_APPEND_LOG
  ];

  let hidden = 0;
  diagnosticSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && !sheet.isSheetHidden()) {
      sheet.hideSheet();
      hidden++;
    }
  });

  Logger.log(`Hidden ${hidden} diagnostic sheet(s)`);
  return hidden;
}

// ============================================================================
// API REQUEST/RESPONSE LOGGING
// ============================================================================

/**
 * Ensures a diagnostic sheet exists, creates it if it doesn't.
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} sheetName - Name of the sheet
 * @param {Array<string>} headers - Column headers
 * @returns {Sheet} The sheet object
 */
function ensureDiagnosticSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
    Logger.log(`Auto-created diagnostic sheet: ${sheetName}`);
  }
  return sheet;
}

/**
 * Logs an API request with full details.
 * Only logs if diagnostic mode is enabled.
 * Auto-creates sheet if it doesn't exist.
 *
 * @param {string} apiName - Name of the API (e.g., "Claude API", "Todoist API")
 * @param {string} endpoint - Full URL endpoint
 * @param {string} method - HTTP method (GET, POST, etc.)
 * @param {Object} headers - Request headers
 * @param {Object|string} payload - Request payload
 * @param {Object} context - Additional context (clientId, eventId, flow)
 * @returns {string|null} Request ID for linking to response, or null if not logged
 */
function logAPIRequest(apiName, endpoint, method, headers, payload, context) {
  if (!isDiagnosticModeEnabled()) return null;

  try {
    const requestId = Utilities.getUuid();
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ensureDiagnosticSheet(
      ss,
      CONFIG.SHEETS.API_REQUEST_LOG,
      ['Timestamp', 'Request_ID', 'API_Name', 'Endpoint', 'Method',
       'Headers', 'Payload', 'Client_ID', 'Event_ID', 'Context']
    );

    // Sanitize headers to remove sensitive data for logging
    const sanitizedHeaders = sanitizeHeaders(headers);

    sheet.appendRow([
      new Date().toISOString(),
      requestId,
      apiName,
      endpoint,
      method,
      JSON.stringify(sanitizedHeaders, null, 2),
      truncateString(typeof payload === 'string' ? payload : JSON.stringify(payload, null, 2), 50000),
      context.clientId || '',
      context.eventId || '',
      context.flow || ''
    ]);

    return requestId;
  } catch (e) {
    Logger.log(`Error logging API request: ${e.message}`);
    return null;
  }
}

/**
 * Logs an API response with full details.
 * Only logs if diagnostic mode is enabled.
 * Auto-creates sheet if it doesn't exist.
 *
 * @param {string} requestId - Request ID from logAPIRequest
 * @param {string} apiName - Name of the API
 * @param {number} statusCode - HTTP status code
 * @param {Object} headers - Response headers
 * @param {string|Object} responseBody - Response body
 * @param {boolean} parseSuccess - Whether JSON parsing succeeded
 * @param {Object} extractedData - Key fields extracted from response
 * @param {string} error - Error message if any
 * @param {number} durationMs - Request duration in milliseconds
 */
function logAPIResponse(requestId, apiName, statusCode, headers, responseBody, parseSuccess, extractedData, error, durationMs) {
  if (!isDiagnosticModeEnabled()) return;

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ensureDiagnosticSheet(
      ss,
      CONFIG.SHEETS.API_RESPONSE_LOG,
      ['Timestamp', 'Request_ID', 'API_Name', 'Status_Code', 'Response_Headers',
       'Response_Body', 'Parse_Success', 'Extracted_Data', 'Error_Message', 'Duration_MS']
    );

    sheet.appendRow([
      new Date().toISOString(),
      requestId || '',
      apiName,
      statusCode,
      JSON.stringify(headers, null, 2),
      truncateString(typeof responseBody === 'string' ? responseBody : JSON.stringify(responseBody, null, 2), 50000),
      parseSuccess ? 'TRUE' : 'FALSE',
      extractedData ? JSON.stringify(extractedData, null, 2) : '',
      error || '',
      durationMs || 0
    ]);
  } catch (e) {
    Logger.log(`Error logging API response: ${e.message}`);
  }
}

// ============================================================================
// DATA COLLECTION LOGGING
// ============================================================================

/**
 * Logs data collection from any source (Todoist, Gmail, Calendar, Google Docs).
 * Only logs if diagnostic mode is enabled.
 * Auto-creates sheet if it doesn't exist.
 *
 * @param {string} clientId - Client identifier
 * @param {string} eventId - Event identifier
 * @param {string} source - Data source name (e.g., "Todoist", "Gmail", "Calendar", "Google Doc")
 * @param {string} queryDetails - Query or search details
 * @param {Array} items - Array of items found
 * @param {Array} sampleData - Sample of first few items
 * @param {string} error - Error message if any
 */
function logDataCollection(clientId, eventId, source, queryDetails, items, sampleData, error) {
  if (!isDiagnosticModeEnabled()) return;

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ensureDiagnosticSheet(
      ss,
      CONFIG.SHEETS.DATA_COLLECTION_LOG,
      ['Timestamp', 'Client_ID', 'Event_ID', 'Source', 'Query_Details',
       'Items_Found', 'Sample_Data', 'Is_Empty', 'Error']
    );

    const itemCount = items ? items.length : 0;
    const isEmpty = !items || items.length === 0;

    sheet.appendRow([
      new Date().toISOString(),
      clientId || '',
      eventId || '',
      source,
      queryDetails || '',
      itemCount,
      sampleData ? JSON.stringify(sampleData, null, 2) : '',
      isEmpty ? 'TRUE' : 'FALSE',
      error || ''
    ]);
  } catch (e) {
    Logger.log(`Error logging data collection: ${e.message}`);
  }
}

// ============================================================================
// AGENDA GENERATION TRACE LOGGING
// ============================================================================

/**
 * Starts a new agenda generation trace and returns the trace ID.
 * Only logs if diagnostic mode is enabled.
 *
 * @param {CalendarEvent} event - Calendar event object
 * @param {Object} client - Client object (can be null if not matched yet)
 * @returns {string|null} Trace ID or null if diagnostic mode disabled
 */
function startAgendaTrace(event, client) {
  if (!isDiagnosticModeEnabled()) return null;

  const traceId = Utilities.getUuid();
  logAgendaStep(traceId, event, client, 1, 'STARTED', 'started', 'Agenda generation initiated');
  return traceId;
}

/**
 * Logs a step in the agenda generation process.
 * Only logs if diagnostic mode is enabled.
 * Auto-creates sheet if it doesn't exist.
 *
 * @param {string} traceId - Trace ID from startAgendaTrace
 * @param {CalendarEvent} event - Calendar event object
 * @param {Object} client - Client object
 * @param {number} stepNumber - Sequential step number
 * @param {string} stepName - Name of the step (e.g., "CLIENT_MATCH", "FETCH_TODOIST")
 * @param {string} status - Step status: "started", "success", "failed", "skipped"
 * @param {string} details - Description of what happened
 * @param {string} dataSummary - Optional summary of data collected
 * @param {number} durationMs - Optional duration in milliseconds
 */
function logAgendaStep(traceId, event, client, stepNumber, stepName, status, details, dataSummary, durationMs) {
  if (!isDiagnosticModeEnabled() || !traceId) return;

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ensureDiagnosticSheet(
      ss,
      CONFIG.SHEETS.AGENDA_GENERATION_TRACE,
      ['Trace_ID', 'Timestamp', 'Event_ID', 'Event_Title', 'Client_ID',
       'Step_Number', 'Step_Name', 'Step_Status', 'Step_Details', 'Data_Summary', 'Duration_MS']
    );

    sheet.appendRow([
      traceId,
      new Date().toISOString(),
      event ? event.getId() : '',
      event ? event.getTitle() : '',
      client ? client.client_name : 'NO_CLIENT',
      stepNumber,
      stepName,
      status,
      details || '',
      dataSummary || '',
      durationMs || 0
    ]);
  } catch (e) {
    Logger.log(`Error logging agenda step: ${e.message}`);
  }
}

// ============================================================================
// NOTES APPEND TRACE LOGGING
// ============================================================================

/**
 * Starts a new notes append trace and returns the trace ID.
 * Only logs if diagnostic mode is enabled.
 *
 * @param {Object} client - Client object
 * @param {string} messageId - Gmail message ID
 * @returns {string|null} Trace ID or null if diagnostic mode disabled
 */
function startNotesAppendTrace(client, messageId) {
  if (!isDiagnosticModeEnabled()) return null;

  const traceId = Utilities.getUuid();
  logNotesAppendStep(traceId, client, messageId, null, 1, 'STARTED', 'started', 'Notes append process initiated');
  return traceId;
}

/**
 * Logs a step in the notes appending process.
 * Only logs if diagnostic mode is enabled.
 *
 * @param {string} traceId - Trace ID from startNotesAppendTrace
 * @param {Object} client - Client object
 * @param {string} messageId - Gmail message ID
 * @param {string} docId - Google Doc ID
 * @param {number} stepNumber - Sequential step number
 * @param {string} stepName - Name of the step (e.g., "DETECT_SENT", "EXTRACT_PLAIN_TEXT")
 * @param {string} status - Step status: "started", "success", "failed", "skipped"
 * @param {string} details - Description of what happened
 * @param {number} contentLength - Optional length of content
 * @param {number} durationMs - Optional duration in milliseconds
 */
function logNotesAppendStep(traceId, client, messageId, docId, stepNumber, stepName, status, details, contentLength, durationMs) {
  if (!isDiagnosticModeEnabled() || !traceId) return;

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ensureDiagnosticSheet(
      ss,
      CONFIG.SHEETS.NOTES_APPEND_TRACE,
      ['Trace_ID', 'Timestamp', 'Client_ID', 'Message_ID', 'Doc_ID',
       'Step_Number', 'Step_Name', 'Step_Status', 'Step_Details', 'Content_Length', 'Duration_MS']
    );

    sheet.appendRow([
      traceId,
      new Date().toISOString(),
      client ? client.client_name : '',
      messageId || '',
      docId || '',
      stepNumber,
      stepName,
      status,
      details || '',
      contentLength || 0,
      durationMs || 0
    ]);
  } catch (e) {
    Logger.log(`Error logging notes append step: ${e.message}`);
  }
}

// ============================================================================
// DOCUMENT APPEND LOGGING
// ============================================================================

/**
 * Logs a document append operation with full details.
 * Only logs if diagnostic mode is enabled.
 * Auto-creates sheet if it doesn't exist.
 *
 * @param {Object} client - Client object
 * @param {string} docId - Google Doc ID
 * @param {string} docUrl - Google Doc URL
 * @param {string} contentType - "Agenda" or "Notes"
 * @param {string} content - Full content being appended
 * @param {boolean} success - Whether append succeeded
 * @param {string} error - Error message if any
 * @param {number} beforeLength - Doc length before append
 * @param {number} afterLength - Doc length after append
 * @param {string} verificationStatus - "verified", "failed", "skipped"
 */
function logDocAppend(client, docId, docUrl, contentType, content, success, error, beforeLength, afterLength, verificationStatus) {
  if (!isDiagnosticModeEnabled()) return;

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ensureDiagnosticSheet(
      ss,
      CONFIG.SHEETS.DOC_APPEND_LOG,
      ['Timestamp', 'Client_ID', 'Doc_ID', 'Doc_URL', 'Content_Type',
       'Content_Preview', 'Content_Full_Length', 'Append_Success',
       'Error_Details', 'Before_Doc_Length', 'After_Doc_Length', 'Verification_Status']
    );

    const contentPreview = content ? content.substring(0, 200) : '';
    const contentLength = content ? content.length : 0;

    sheet.appendRow([
      new Date().toISOString(),
      client ? client.client_name : '',
      docId || '',
      docUrl || '',
      contentType,
      contentPreview,
      contentLength,
      success ? 'TRUE' : 'FALSE',
      error || '',
      beforeLength || 0,
      afterLength || 0,
      verificationStatus || 'skipped'
    ]);
  } catch (e) {
    Logger.log(`Error logging doc append: ${e.message}`);
  }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Truncates a string to a maximum length.
 * @param {string} str - String to truncate
 * @param {number} maxLength - Maximum length
 * @returns {string} Truncated string
 */
function truncateString(str, maxLength) {
  if (!str) return '';
  if (str.length <= maxLength) return str;
  return str.substring(0, maxLength) + '... [TRUNCATED]';
}

/**
 * Sanitizes headers to remove sensitive data.
 * @param {Object} headers - Headers object
 * @returns {Object} Sanitized headers
 */
function sanitizeHeaders(headers) {
  if (!headers) return {};

  const sanitized = {...headers};

  // Remove or mask sensitive headers
  const sensitiveKeys = ['x-api-key', 'authorization', 'api-key', 'token'];

  Object.keys(sanitized).forEach(key => {
    if (sensitiveKeys.some(sensitive => key.toLowerCase().includes(sensitive))) {
      sanitized[key] = '[REDACTED]';
    }
  });

  return sanitized;
}
