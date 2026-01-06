/**
 * Code.gs - Entry Points and Trigger Handlers
 *
 * This file contains the main entry points for the Client Management Automation System.
 * It handles Web App endpoints, trigger setup, and routes requests to appropriate modules.
 */

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

const CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),
  SHEETS: {
    CLIENT_REGISTRY: 'Client_Registry',
    GENERATED_AGENDAS: 'Generated_Agendas',
    PROCESSING_LOG: 'Processing_Log',
    UNMATCHED: 'Unmatched'
  },
  BUSINESS_HOURS: {
    START: 8,  // 8:00 AM
    END: 18    // 6:00 PM
  }
};

// ============================================================================
// WEB APP ENDPOINTS
// ============================================================================

/**
 * Handles HTTP GET requests to the Web App.
 * Used for health checks and status verification.
 *
 * @param {Object} e - The event object containing request parameters
 * @returns {TextOutput} JSON response with status
 */
function doGet(e) {
  const response = {
    status: 'ok',
    message: 'Client Management Automation System is running',
    timestamp: new Date().toISOString()
  };

  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles HTTP POST requests to the Web App.
 * Primary endpoint for Fathom webhook integration.
 *
 * @param {Object} e - The event object containing POST data
 * @returns {TextOutput} JSON response with processing result
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // Process the Fathom webhook
    const result = processFathomWebhook(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', result: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    logProcessing('WEBHOOK_ERROR', null, error.message, 'error');

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================================
// TRIGGER SETUP
// ============================================================================

/**
 * Creates all required time-based triggers for the automation system.
 * Should be run once during initial setup.
 *
 * Triggers created:
 * - Label and filter creation: daily at 6:00 AM
 * - Sent meeting summary monitor: every 10 minutes
 * - Agenda generation: every 1 hour (limited to business hours in handler)
 * - Daily outlook: daily at 7:00 AM
 * - Weekly outlook: weekly on Monday at 7:00 AM
 * - Client onboarding check: daily at 6:30 AM
 */
function setupTriggers() {
  // First, remove any existing triggers to avoid duplicates
  removeAllTriggers();

  // Label and filter creation - daily at 6:00 AM
  ScriptApp.newTrigger('runLabelAndFilterCreation')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  // Client onboarding check - daily at 6:30 AM
  ScriptApp.newTrigger('runClientOnboarding')
    .timeBased()
    .atHour(6)
    .nearMinute(30)
    .everyDays(1)
    .create();

  // Sent meeting summary monitor - every 10 minutes
  ScriptApp.newTrigger('runSentMeetingSummaryMonitor')
    .timeBased()
    .everyMinutes(10)
    .create();

  // Agenda generation - every hour (business hours check done in handler)
  ScriptApp.newTrigger('runAgendaGeneration')
    .timeBased()
    .everyHours(1)
    .create();

  // Daily outlook - daily at 7:00 AM
  ScriptApp.newTrigger('runDailyOutlook')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  // Weekly outlook - Monday at 7:00 AM
  ScriptApp.newTrigger('runWeeklyOutlook')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();

  Logger.log('All triggers have been set up successfully.');
}

/**
 * Removes all existing triggers for this script.
 * Useful for cleanup or before re-creating triggers.
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  Logger.log(`Removed ${triggers.length} existing triggers.`);
}

// ============================================================================
// TRIGGER HANDLERS
// ============================================================================

/**
 * Handler for daily label and filter creation trigger.
 * Runs at 6:00 AM daily.
 */
function runLabelAndFilterCreation() {
  try {
    logProcessing('FILTER_SYNC', null, 'Starting label and filter synchronization', 'processing');
    syncLabelsAndFilters();
    logProcessing('FILTER_SYNC', null, 'Label and filter synchronization completed', 'success');
  } catch (error) {
    logProcessing('FILTER_SYNC', null, `Error: ${error.message}`, 'error');
  }
}

/**
 * Handler for client onboarding trigger.
 * Runs at 6:30 AM daily to check for new clients needing setup.
 */
function runClientOnboarding() {
  try {
    logProcessing('CLIENT_ONBOARD', null, 'Starting client onboarding check', 'processing');
    processNewClients();
    logProcessing('CLIENT_ONBOARD', null, 'Client onboarding check completed', 'success');
  } catch (error) {
    logProcessing('CLIENT_ONBOARD', null, `Error: ${error.message}`, 'error');
  }
}

/**
 * Handler for sent meeting summary monitor trigger.
 * Runs every 10 minutes.
 */
function runSentMeetingSummaryMonitor() {
  try {
    monitorSentMeetingSummaries();
  } catch (error) {
    logProcessing('SUMMARY_MONITOR', null, `Error: ${error.message}`, 'error');
  }
}

/**
 * Handler for agenda generation trigger.
 * Runs every hour, but only processes during business hours (8 AM - 6 PM).
 */
function runAgendaGeneration() {
  const currentHour = new Date().getHours();

  // Only run during business hours
  if (currentHour < CONFIG.BUSINESS_HOURS.START || currentHour >= CONFIG.BUSINESS_HOURS.END) {
    Logger.log(`Skipping agenda generation - outside business hours (${currentHour}:00)`);
    return;
  }

  try {
    logProcessing('AGENDA_GEN', null, 'Starting agenda generation check', 'processing');
    generateAgendas();
    logProcessing('AGENDA_GEN', null, 'Agenda generation check completed', 'success');
  } catch (error) {
    logProcessing('AGENDA_GEN', null, `Error: ${error.message}`, 'error');
  }
}

/**
 * Handler for daily outlook trigger.
 * Runs at 7:00 AM daily.
 */
function runDailyOutlook() {
  try {
    logProcessing('DAILY_OUTLOOK', null, 'Generating daily outlook', 'processing');
    generateDailyOutlook();
    logProcessing('DAILY_OUTLOOK', null, 'Daily outlook generated', 'success');
  } catch (error) {
    logProcessing('DAILY_OUTLOOK', null, `Error: ${error.message}`, 'error');
  }
}

/**
 * Handler for weekly outlook trigger.
 * Runs at 7:00 AM every Monday.
 */
function runWeeklyOutlook() {
  try {
    logProcessing('WEEKLY_OUTLOOK', null, 'Generating weekly outlook', 'processing');
    generateWeeklyOutlook();
    logProcessing('WEEKLY_OUTLOOK', null, 'Weekly outlook generated', 'success');
  } catch (error) {
    logProcessing('WEEKLY_OUTLOOK', null, `Error: ${error.message}`, 'error');
  }
}

// ============================================================================
// MANUAL EXECUTION FUNCTIONS
// ============================================================================

/**
 * Manually runs label and filter synchronization.
 * Useful for initial setup or testing.
 */
function manualSyncLabelsAndFilters() {
  runLabelAndFilterCreation();
}

/**
 * Manually runs agenda generation.
 * Bypasses business hours check for testing purposes.
 */
function manualGenerateAgendas() {
  try {
    logProcessing('AGENDA_GEN', null, 'Manual agenda generation started', 'processing');
    generateAgendas();
    logProcessing('AGENDA_GEN', null, 'Manual agenda generation completed', 'success');
  } catch (error) {
    logProcessing('AGENDA_GEN', null, `Error: ${error.message}`, 'error');
  }
}

/**
 * Manually runs daily outlook generation.
 * Useful for testing.
 */
function manualDailyOutlook() {
  runDailyOutlook();
}

/**
 * Manually runs weekly outlook generation.
 * Useful for testing.
 */
function manualWeeklyOutlook() {
  runWeeklyOutlook();
}

// ============================================================================
// INITIALIZATION
// ============================================================================

/**
 * Initializes the spreadsheet with required sheets if they don't exist.
 * Should be run once during initial setup.
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Client_Registry sheet
  createSheetIfNotExists(ss, CONFIG.SHEETS.CLIENT_REGISTRY, [
    'client_id', 'client_name', 'email_domains', 'contact_emails',
    'google_doc_url', 'todoist_project_id'
  ]);

  // Generated_Agendas sheet
  createSheetIfNotExists(ss, CONFIG.SHEETS.GENERATED_AGENDAS, [
    'event_id', 'event_title', 'client_id', 'generated_timestamp'
  ]);

  // Processing_Log sheet
  createSheetIfNotExists(ss, CONFIG.SHEETS.PROCESSING_LOG, [
    'timestamp', 'action_type', 'client_id', 'details', 'status'
  ]);

  // Unmatched sheet
  createSheetIfNotExists(ss, CONFIG.SHEETS.UNMATCHED, [
    'timestamp', 'item_type', 'item_details', 'participant_emails', 'manually_resolved'
  ]);

  Logger.log('Spreadsheet initialization completed.');
}

/**
 * Creates a sheet with headers if it doesn't already exist.
 *
 * @param {Spreadsheet} ss - The spreadsheet object
 * @param {string} sheetName - Name of the sheet to create
 * @param {string[]} headers - Array of column headers
 */
function createSheetIfNotExists(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log(`Created sheet: ${sheetName}`);
  } else {
    Logger.log(`Sheet already exists: ${sheetName}`);
  }
}
