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
    UNMATCHED: 'Unmatched',
    FOLDERS: 'Folders',
    PROCESSED_FATHOM: 'Processed_Fathom_Meetings'
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
    // Log all incoming webhooks for debugging
    logProcessing('WEBHOOK_RECEIVED', null, 'Received POST request', 'info');

    // Verify webhook signature if secret is configured
    const webhookSecret = PropertiesService.getScriptProperties().getProperty('FATHOM_WEBHOOK_SECRET');
    const enforceSignature = PropertiesService.getScriptProperties().getProperty('FATHOM_ENFORCE_SIGNATURE');

    if (webhookSecret) {
      const isValid = verifyFathomWebhookSignature(e, webhookSecret);
      if (!isValid) {
        // Log warning but only reject if enforcement is explicitly enabled
        logProcessing('WEBHOOK_INVALID_SIGNATURE', null, 'Webhook signature verification failed - processing anyway (set FATHOM_ENFORCE_SIGNATURE=true to reject)', 'warning');

        // Only reject if enforcement is explicitly enabled
        if (enforceSignature === 'true') {
          return ContentService
            .createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid signature' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      } else {
        logProcessing('WEBHOOK_SIGNATURE', null, 'Webhook signature verified successfully', 'success');
      }
    } else {
      logProcessing('WEBHOOK_NO_SECRET', null, 'No webhook secret configured - skipping signature verification', 'info');
    }

    let payload = JSON.parse(e.postData.contents);

    // Log payload summary for debugging
    logProcessing('WEBHOOK_PAYLOAD', null, `Processing meeting: ${payload.meeting_title || payload.title || 'Unknown'}`, 'info');

    // Normalize webhook payload to ensure consistent format
    // Fathom webhooks might have different field names than API responses
    payload = normalizeFathomPayload(payload);

    // Process the Fathom webhook
    const result = processFathomWebhook(payload);

    logProcessing('WEBHOOK_SUCCESS', null, `Successfully processed webhook for: ${payload.meeting_title || 'Unknown'}`, 'success');

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', result: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    const errorDetail = `${error.message}\nStack: ${error.stack || 'N/A'}`;
    logProcessing('WEBHOOK_ERROR', null, errorDetail, 'error');

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Verifies the Fathom webhook signature.
 * Fathom signs webhooks using HMAC SHA-256 with the webhook secret.
 *
 * @param {Object} e - The event object containing POST data and headers
 * @param {string} secret - The webhook secret
 * @returns {boolean} True if signature is valid
 */
function verifyFathomWebhookSignature(e, secret) {
  try {
    // Get signature from header - Fathom uses 'webhook-signature' or 'x-fathom-signature'
    const signatureHeader = e.parameter['webhook-signature'] ||
                           e.parameter['x-fathom-signature'] ||
                           (e.headers && (e.headers['webhook-signature'] || e.headers['x-fathom-signature']));

    if (!signatureHeader) {
      Logger.log('No webhook signature header found');
      return false;
    }

    // Get raw body
    const rawBody = e.postData.contents;

    // Fathom signature format: "v1,<base64-signature> <base64-signature2>..."
    // Extract signatures after the version prefix
    const parts = signatureHeader.split(',');
    if (parts.length < 2) {
      Logger.log('Invalid signature header format');
      return false;
    }

    const signatures = parts[1].trim().split(' ');

    // Compute expected signature using HMAC SHA-256
    const expectedSignature = Utilities.base64Encode(
      Utilities.computeHmacSha256Signature(rawBody, secret)
    );

    // Check if any of the provided signatures match
    for (const sig of signatures) {
      if (sig === expectedSignature) {
        return true;
      }
    }

    Logger.log('Webhook signature mismatch');
    return false;

  } catch (error) {
    Logger.log('Webhook verification error: ' + error.message);
    return false;
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
 * - Google Drive folder sync: daily at 5:30 AM
 * - Label and filter creation: daily at 6:00 AM
 * - Client onboarding check: daily at 6:30 AM
 * - Daily outlook: daily at 7:00 AM
 * - Weekly outlook: weekly on Monday at 7:00 AM
 * - Sent meeting summary monitor: every 10 minutes
 * - Agenda generation: every 1 hour (limited to business hours in handler)
 */
function setupTriggers() {
  // First, remove any existing triggers to avoid duplicates
  removeAllTriggers();

  // Google Drive folder sync - daily at 5:30 AM
  ScriptApp.newTrigger('runFolderSync')
    .timeBased()
    .atHour(5)
    .nearMinute(30)
    .everyDays(1)
    .create();

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

  // Fathom API polling - every 30 minutes (PRIMARY METHOD)
  // Automatically fetches meetings from Fathom API
  ScriptApp.newTrigger('runFathomPolling')
    .timeBased()
    .everyMinutes(30)
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

/**
 * Creates the Fathom polling trigger (30-minute intervals).
 * Called when user enables polling backup in settings.
 */
function createFathomPollingTrigger() {
  // First remove any existing polling triggers
  removeFathomPollingTrigger();

  // Create new trigger
  ScriptApp.newTrigger('runFathomPolling')
    .timeBased()
    .everyMinutes(30)
    .create();

  Logger.log('Fathom polling trigger created');
}

/**
 * Removes the Fathom polling trigger.
 * Called when user disables polling backup in settings.
 */
function removeFathomPollingTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runFathomPolling') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  if (removed > 0) {
    Logger.log(`Removed ${removed} Fathom polling trigger(s)`);
  }
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
 * Handler for Fathom API polling trigger.
 * Runs every 30 minutes - PRIMARY METHOD for fetching meetings.
 */
function runFathomPolling() {
  try {
    pollFathomForNewMeetings();
  } catch (error) {
    logProcessing('FATHOM_POLL', null, `Error: ${error.message}`, 'error');
  }
}

/**
 * Handler for agenda generation trigger.
 * Runs every hour, but only processes during business hours.
 * Business hours are configurable via Script Properties (default 8 AM - 6 PM).
 */
function runAgendaGeneration() {
  const currentHour = new Date().getHours();

  // Read business hours from Script Properties (fall back to CONFIG defaults)
  const props = PropertiesService.getScriptProperties();
  const startHour = parseInt(props.getProperty('BUSINESS_HOURS_START'), 10) || CONFIG.BUSINESS_HOURS.START;
  const endHour = parseInt(props.getProperty('BUSINESS_HOURS_END'), 10) || CONFIG.BUSINESS_HOURS.END;

  // Only run during business hours
  if (currentHour < startHour || currentHour >= endHour) {
    Logger.log(`Skipping agenda generation - outside business hours (${currentHour}:00, configured: ${startHour}-${endHour})`);
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
 * Note: Prefer using SETUP_RUN_THIS_FIRST() in Utilities.gs for full setup.
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Use the createAllSheets function from Utilities.gs
  createAllSheets(ss);

  Logger.log('Spreadsheet initialization completed.');

  // Run initial folder sync to populate the Folders sheet
  Logger.log('Running initial folder sync...');
  syncDriveFolders();
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
