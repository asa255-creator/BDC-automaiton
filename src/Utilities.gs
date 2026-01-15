/**
 * Utilities.gs - Shared Helpers
 *
 * This module contains shared utility functions used across the
 * Client Management Automation System.
 */

// ============================================================================
// DATE AND TIME FORMATTING
// ============================================================================

/**
 * Formats a date as YYYY-MM-DD.
 *
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  if (!date || !(date instanceof Date)) {
    return '';
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

/**
 * Formats a date with a human-readable format (e.g., "January 6, 2025").
 *
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDateLong(date) {
  if (!date || !(date instanceof Date)) {
    return '';
  }

  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return date.toLocaleDateString('en-US', options);
}

/**
 * Formats a date in short format with ordinal (e.g., "Jan 6th").
 *
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDateShort(date) {
  if (!date || !(date instanceof Date)) {
    return '';
  }

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const month = months[date.getMonth()];
  const day = date.getDate();
  const ordinal = getOrdinalSuffix(day);

  return `${month} ${day}${ordinal}`;
}

/**
 * Gets the ordinal suffix for a number (st, nd, rd, th).
 *
 * @param {number} n - The number
 * @returns {string} The ordinal suffix
 */
function getOrdinalSuffix(n) {
  const s = ['th', 'st', 'nd', 'rd'];
  const v = n % 100;
  return s[(v - 20) % 10] || s[v] || s[0];
}

/**
 * Formats a time as HH:MM AM/PM.
 *
 * @param {Date} date - The date/time to format
 * @returns {string} Formatted time string
 */
function formatTime(date) {
  if (!date || !(date instanceof Date)) {
    return '';
  }

  let hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const ampm = hours >= 12 ? 'PM' : 'AM';

  hours = hours % 12;
  hours = hours ? hours : 12; // 0 should be 12

  return `${hours}:${minutes} ${ampm}`;
}

/**
 * Formats a full date and time.
 *
 * @param {Date} date - The date/time to format
 * @returns {string} Formatted date and time string
 */
function formatDateTime(date) {
  if (!date || !(date instanceof Date)) {
    return '';
  }

  return `${formatDateLong(date)} at ${formatTime(date)}`;
}

/**
 * Parses an ISO date string into a Date object.
 *
 * @param {string} isoString - ISO date string
 * @returns {Date|null} Parsed Date or null
 */
function parseISODate(isoString) {
  if (!isoString) {
    return null;
  }

  try {
    const date = new Date(isoString);
    return isNaN(date.getTime()) ? null : date;
  } catch (e) {
    return null;
  }
}

/**
 * Gets the start of today (midnight).
 *
 * @returns {Date} Start of today
 */
function getStartOfToday() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return today;
}

/**
 * Gets the end of today (23:59:59.999).
 *
 * @returns {Date} End of today
 */
function getEndOfToday() {
  const today = new Date();
  today.setHours(23, 59, 59, 999);
  return today;
}

/**
 * Gets the start of the current week (Monday).
 *
 * @returns {Date} Start of week
 */
function getStartOfWeek() {
  const today = new Date();
  const day = today.getDay();
  const diff = today.getDate() - day + (day === 0 ? -6 : 1);
  const monday = new Date(today);
  monday.setDate(diff);
  monday.setHours(0, 0, 0, 0);
  return monday;
}

/**
 * Gets the end of the current week (Sunday).
 *
 * @returns {Date} End of week
 */
function getEndOfWeek() {
  const startOfWeek = getStartOfWeek();
  const sunday = new Date(startOfWeek);
  sunday.setDate(sunday.getDate() + 6);
  sunday.setHours(23, 59, 59, 999);
  return sunday;
}

// ============================================================================
// LOGGING
// ============================================================================

/**
 * Logs a processing action to the Processing_Log sheet.
 *
 * @param {string} actionType - Type of action (e.g., 'WEBHOOK_PROCESS', 'AGENDA_GEN')
 * @param {string|null} clientId - The client ID if applicable
 * @param {string} details - Details about the action
 * @param {string} status - Status ('success', 'error', 'processing', 'unmatched', 'warning')
 */
function logProcessing(actionType, clientId, details, status) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PROCESSING_LOG);

    if (!sheet) {
      Logger.log('Processing_Log sheet not found');
      return;
    }

    sheet.appendRow([
      new Date().toISOString(),
      actionType,
      clientId || '',
      details,
      status
    ]);

  } catch (error) {
    Logger.log(`Failed to log processing: ${error.message}`);
  }
}

/**
 * Gets recent processing logs.
 *
 * @param {number} limit - Maximum number of logs to return
 * @param {string} actionType - Optional filter by action type
 * @returns {Object[]} Array of log entries
 */
function getRecentLogs(limit = 50, actionType = null) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROCESSING_LOG);

  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const logs = [];

  for (let i = data.length - 1; i >= 1 && logs.length < limit; i--) {
    const row = data[i];
    const log = {};

    headers.forEach((header, index) => {
      log[header] = row[index];
    });

    if (!actionType || log.action_type === actionType) {
      logs.push(log);
    }
  }

  return logs;
}

/**
 * Clears old processing logs (older than specified days).
 *
 * @param {number} daysToKeep - Number of days of logs to retain
 */
function clearOldLogs(daysToKeep = 30) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROCESSING_LOG);

  if (!sheet) {
    return;
  }

  const data = sheet.getDataRange().getValues();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const timestamp = new Date(data[i][0]);
    if (timestamp < cutoffDate) {
      rowsToDelete.push(i + 1); // 1-based row index
    }
  }

  // Delete rows from bottom to top to preserve indices
  rowsToDelete.reverse().forEach(rowIndex => {
    sheet.deleteRow(rowIndex);
  });

  Logger.log(`Cleared ${rowsToDelete.length} old log entries`);
}

// ============================================================================
// STRING UTILITIES
// ============================================================================

/**
 * Truncates a string to a specified length with ellipsis.
 *
 * @param {string} str - The string to truncate
 * @param {number} maxLength - Maximum length
 * @returns {string} Truncated string
 */
function truncate(str, maxLength) {
  if (!str || str.length <= maxLength) {
    return str || '';
  }
  return str.substring(0, maxLength - 3) + '...';
}

/**
 * Sanitizes HTML by escaping special characters.
 *
 * @param {string} text - Text to sanitize
 * @returns {string} Sanitized text
 */
function escapeHtml(text) {
  if (!text) return '';

  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Strips HTML tags from a string.
 *
 * @param {string} html - HTML string
 * @returns {string} Plain text
 */
function stripHtml(html) {
  if (!html) return '';
  return html.replace(/<[^>]*>/g, '').trim();
}

/**
 * Capitalizes the first letter of a string.
 *
 * @param {string} str - String to capitalize
 * @returns {string} Capitalized string
 */
function capitalize(str) {
  if (!str) return '';
  return str.charAt(0).toUpperCase() + str.slice(1);
}

/**
 * Converts a string to title case.
 *
 * @param {string} str - String to convert
 * @returns {string} Title cased string
 */
function toTitleCase(str) {
  if (!str) return '';
  return str.replace(/\w\S*/g, txt =>
    txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()
  );
}

// ============================================================================
// ARRAY UTILITIES
// ============================================================================

/**
 * Removes duplicates from an array.
 *
 * @param {Array} arr - Array with potential duplicates
 * @returns {Array} Array with duplicates removed
 */
function uniqueArray(arr) {
  return [...new Set(arr)];
}

/**
 * Groups an array of objects by a key.
 *
 * @param {Object[]} arr - Array of objects
 * @param {string} key - Key to group by
 * @returns {Object} Grouped object
 */
function groupBy(arr, key) {
  return arr.reduce((grouped, item) => {
    const groupKey = item[key] || 'undefined';
    if (!grouped[groupKey]) {
      grouped[groupKey] = [];
    }
    grouped[groupKey].push(item);
    return grouped;
  }, {});
}

/**
 * Sorts an array of objects by a key.
 *
 * @param {Object[]} arr - Array of objects
 * @param {string} key - Key to sort by
 * @param {boolean} ascending - Sort direction (default true)
 * @returns {Object[]} Sorted array
 */
function sortBy(arr, key, ascending = true) {
  return [...arr].sort((a, b) => {
    const valA = a[key];
    const valB = b[key];

    if (valA < valB) return ascending ? -1 : 1;
    if (valA > valB) return ascending ? 1 : -1;
    return 0;
  });
}

// ============================================================================
// ERROR HANDLING
// ============================================================================

/**
 * Wraps a function with error handling and logging.
 *
 * @param {Function} fn - Function to wrap
 * @param {string} actionType - Action type for logging
 * @returns {Function} Wrapped function
 */
function withErrorHandling(fn, actionType) {
  return function(...args) {
    try {
      return fn.apply(this, args);
    } catch (error) {
      logProcessing(actionType, null, `Error: ${error.message}`, 'error');
      Logger.log(`${actionType} error: ${error.message}`);
      throw error;
    }
  };
}

/**
 * Retries a function with exponential backoff.
 *
 * @param {Function} fn - Function to retry
 * @param {number} maxRetries - Maximum number of retries
 * @param {number} baseDelay - Base delay in milliseconds
 * @returns {*} Result of the function
 */
function retryWithBackoff(fn, maxRetries = 3, baseDelay = 1000) {
  let lastError;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return fn();
    } catch (error) {
      lastError = error;

      if (attempt < maxRetries) {
        const delay = baseDelay * Math.pow(2, attempt);
        Logger.log(`Retry attempt ${attempt + 1} after ${delay}ms`);
        Utilities.sleep(delay);
      }
    }
  }

  throw lastError;
}

// ============================================================================
// SPREADSHEET UTILITIES
// ============================================================================

/**
 * Gets data from a sheet as an array of objects.
 *
 * @param {string} sheetName - Name of the sheet
 * @returns {Object[]} Array of row objects
 */
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return [];
  }

  const headers = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((header, index) => {
      row[header] = data[i][index];
    });
    rows.push(row);
  }

  return rows;
}

/**
 * Appends data to a sheet.
 *
 * @param {string} sheetName - Name of the sheet
 * @param {Object} rowData - Object with column values
 * @param {string[]} columns - Array of column names in order
 */
function appendToSheet(sheetName, rowData, columns) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`Sheet not found: ${sheetName}`);
    return;
  }

  const row = columns.map(col => rowData[col] || '');
  sheet.appendRow(row);
}

/**
 * Finds a row in a sheet by a column value.
 *
 * @param {string} sheetName - Name of the sheet
 * @param {string} columnName - Column to search
 * @param {*} value - Value to find
 * @returns {Object|null} Row data or null
 */
function findRowByColumn(sheetName, columnName, value) {
  const data = getSheetData(sheetName);

  for (const row of data) {
    if (row[columnName] === value) {
      return row;
    }
  }

  return null;
}

// ============================================================================
// VALIDATION
// ============================================================================

/**
 * Validates that required properties are set.
 *
 * @param {string[]} requiredProps - Array of required property names
 * @returns {Object} Validation result with isValid and missing properties
 */
function validateRequiredProperties(requiredProps) {
  const properties = PropertiesService.getScriptProperties();
  const missing = [];

  for (const prop of requiredProps) {
    if (!properties.getProperty(prop)) {
      missing.push(prop);
    }
  }

  return {
    isValid: missing.length === 0,
    missing: missing
  };
}

/**
 * Validates the system configuration.
 *
 * @returns {Object} Validation result
 */
function validateConfiguration() {
  const requiredProps = ['SPREADSHEET_ID', 'TODOIST_API_TOKEN', 'CLAUDE_API_KEY'];
  const propValidation = validateRequiredProperties(requiredProps);

  const result = {
    isValid: true,
    errors: []
  };

  if (!propValidation.isValid) {
    result.isValid = false;
    result.errors.push(`Missing properties: ${propValidation.missing.join(', ')}`);
  }

  // Check spreadsheet access
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const requiredSheets = Object.values(CONFIG.SHEETS);

    for (const sheetName of requiredSheets) {
      if (!ss.getSheetByName(sheetName)) {
        result.errors.push(`Missing sheet: ${sheetName}`);
        result.isValid = false;
      }
    }
  } catch (e) {
    result.isValid = false;
    result.errors.push(`Cannot access spreadsheet: ${e.message}`);
  }

  return result;
}

// ============================================================================
// CACHING HELPERS
// ============================================================================

/**
 * Gets a value from cache with automatic JSON parsing.
 *
 * @param {string} key - Cache key
 * @returns {*} Cached value or null
 */
function getCached(key) {
  const cache = CacheService.getScriptCache();
  const value = cache.get(key);

  if (!value) return null;

  try {
    return JSON.parse(value);
  } catch (e) {
    return value;
  }
}

/**
 * Sets a value in cache with automatic JSON stringification.
 *
 * @param {string} key - Cache key
 * @param {*} value - Value to cache
 * @param {number} expirationInSeconds - Cache duration (default 6 hours)
 */
function setCached(key, value, expirationInSeconds = 21600) {
  const cache = CacheService.getScriptCache();
  const stringValue = typeof value === 'string' ? value : JSON.stringify(value);
  cache.put(key, stringValue, expirationInSeconds);
}

/**
 * Removes a value from cache.
 *
 * @param {string} key - Cache key
 */
function removeCached(key) {
  const cache = CacheService.getScriptCache();
  cache.remove(key);
}

// ============================================================================
// ONE-CLICK SETUP
// ============================================================================

/**
 * MAIN SETUP FUNCTION - Run this first!
 *
 * Works in two modes:
 *
 * MODE 1 - Bound to a Sheet (recommended):
 * - Create script via Extensions > Apps Script from a Google Sheet
 * - Run this function - it will prompt for all settings via dialogs
 *
 * MODE 2 - Standalone Script:
 * - Set SPREADSHEET_ID in Script Properties first (required)
 * - Optionally set: TODOIST_API_TOKEN, CLAUDE_API_KEY, USER_NAME,
 *   BUSINESS_HOURS_START, BUSINESS_HOURS_END, DOC_NAME_TEMPLATE
 * - Run this function - it will use those properties and set defaults
 *
 * In both modes, this function:
 * - Checks for required Advanced Services (Calendar API, Gmail API)
 * - Creates all required sheets
 * - Sets up all triggers
 * - Syncs Google Drive folders
 */
function SETUP_RUN_THIS_FIRST() {
  const props = PropertiesService.getScriptProperties();

  // Detect if we're bound to a spreadsheet (can use UI)
  let ui = null;
  let ss = null;
  let isInteractive = false;

  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      ui = SpreadsheetApp.getUi();
      isInteractive = true;
    }
  } catch (e) {
    // Not bound to a spreadsheet, will use standalone mode
    Logger.log('Running in standalone mode (no UI available)');
  }

  if (isInteractive) {
    // ========== INTERACTIVE MODE (bound to sheet) ==========
    runInteractiveSetup(ui, ss, props);
  } else {
    // ========== STANDALONE MODE ==========
    runStandaloneSetup(props);
  }
}

/**
 * Runs setup in interactive mode with UI prompts.
 */
function runInteractiveSetup(ui, ss, props) {
  // Step 1: Get spreadsheet ID automatically
  const spreadsheetId = ss.getId();
  props.setProperty('SPREADSHEET_ID', spreadsheetId);
  Logger.log(`Set SPREADSHEET_ID: ${spreadsheetId}`);

  // Step 2: Prompt for Todoist API Token
  const todoistResponse = ui.prompt(
    'Todoist API Token',
    'Enter your Todoist API token (from Todoist Settings > Integrations > Developer):\n\nLeave blank to skip Todoist integration.',
    ui.ButtonSet.OK_CANCEL
  );

  if (todoistResponse.getSelectedButton() === ui.Button.CANCEL) {
    ui.alert('Setup cancelled.');
    return;
  }

  const todoistToken = todoistResponse.getResponseText().trim();
  if (todoistToken) {
    props.setProperty('TODOIST_API_TOKEN', todoistToken);
    Logger.log('Set TODOIST_API_TOKEN');
  }

  // Step 3: Prompt for Claude API Key
  const claudeResponse = ui.prompt(
    'Claude API Key',
    'Enter your Anthropic Claude API key:\n\nLeave blank to skip AI agenda generation.',
    ui.ButtonSet.OK_CANCEL
  );

  if (claudeResponse.getSelectedButton() === ui.Button.CANCEL) {
    ui.alert('Setup cancelled.');
    return;
  }

  const claudeKey = claudeResponse.getResponseText().trim();
  if (claudeKey) {
    props.setProperty('CLAUDE_API_KEY', claudeKey);
    Logger.log('Set CLAUDE_API_KEY');
  }

  // Step 4: Prompt for user's name (for email signatures)
  const nameResponse = ui.prompt(
    'Your Name',
    'Enter your name/initials for email signatures (e.g., "TC", "John"):\n\nThis appears at the end of meeting summary emails.',
    ui.ButtonSet.OK_CANCEL
  );

  if (nameResponse.getSelectedButton() === ui.Button.CANCEL) {
    ui.alert('Setup cancelled.');
    return;
  }

  const userName = nameResponse.getResponseText().trim() || 'Team';
  props.setProperty('USER_NAME', userName);
  Logger.log(`Set USER_NAME: ${userName}`);

  // Step 5: Prompt for business hours
  const hoursResponse = ui.prompt(
    'Business Hours',
    'Enter your business hours for agenda generation (format: START-END, e.g., "8-18" for 8 AM to 6 PM):\n\nAgendas are only generated during these hours.\n\nPress OK for default (8-18) or enter custom hours:',
    ui.ButtonSet.OK_CANCEL
  );

  if (hoursResponse.getSelectedButton() === ui.Button.CANCEL) {
    ui.alert('Setup cancelled.');
    return;
  }

  const hoursInput = hoursResponse.getResponseText().trim();
  let startHour = 8;
  let endHour = 18;

  if (hoursInput && hoursInput.includes('-')) {
    const parts = hoursInput.split('-');
    const parsedStart = parseInt(parts[0], 10);
    const parsedEnd = parseInt(parts[1], 10);
    if (!isNaN(parsedStart) && !isNaN(parsedEnd) && parsedStart >= 0 && parsedEnd <= 24) {
      startHour = parsedStart;
      endHour = parsedEnd;
    }
  }

  props.setProperty('BUSINESS_HOURS_START', startHour.toString());
  props.setProperty('BUSINESS_HOURS_END', endHour.toString());
  Logger.log(`Set business hours: ${startHour} to ${endHour}`);

  // Step 6: Prompt for doc naming template
  const docNameResponse = ui.prompt(
    'Document Naming',
    'Enter the naming template for client meeting notes docs:\n\nUse {client_name} as placeholder.\n\nPress OK for default or enter custom template:',
    ui.ButtonSet.OK_CANCEL
  );

  if (docNameResponse.getSelectedButton() === ui.Button.CANCEL) {
    ui.alert('Setup cancelled.');
    return;
  }

  const docTemplate = docNameResponse.getResponseText().trim() || 'Client Notes - {client_name}';
  props.setProperty('DOC_NAME_TEMPLATE', docTemplate);
  Logger.log(`Set DOC_NAME_TEMPLATE: ${docTemplate}`);

  // Step 7: Check Advanced Services
  const serviceStatus = checkAdvancedServices();

  if (serviceStatus.missing.length > 0) {
    const enableNow = ui.alert(
      'Advanced Services Required',
      'The following Advanced Services need to be enabled:\n\n' +
      serviceStatus.missing.join('\n') + '\n\n' +
      'To enable them:\n' +
      '1. Click "Services" (+ icon) in the left sidebar\n' +
      '2. Find each service listed above\n' +
      '3. Click "Add" for each one\n\n' +
      'Click OK after enabling them, or Cancel to continue without (some features will be disabled).',
      ui.ButtonSet.OK_CANCEL
    );

    if (enableNow === ui.Button.OK) {
      // Re-check after user says they enabled
      const recheck = checkAdvancedServices();
      if (recheck.missing.length > 0) {
        ui.alert('Warning', 'Some services still not detected:\n' + recheck.missing.join('\n') + '\n\nContinuing setup - some features may not work.', ui.ButtonSet.OK);
      }
    }
  }

  // Step 8: Create all sheets
  ui.alert('Creating Sheets', 'Creating all required sheets...', ui.ButtonSet.OK);

  createAllSheets(ss);

  // Step 9: Set up triggers
  setupAllTriggers();

  // Step 10: Sync Drive folders
  ui.alert('Syncing Folders', 'Scanning Google Drive folders (this may take a moment)...', ui.ButtonSet.OK);

  try {
    syncDriveFolders();
  } catch (e) {
    Logger.log(`Folder sync error: ${e.message}`);
  }

  // Done!
  let completionMessage = 'Your Client Management Automation System is ready.\n\n' +
    'Next steps:\n' +
    '1. Add clients to the Client_Registry sheet\n' +
    '2. Select a folder from the docs_folder_path dropdown\n' +
    '3. The system will automatically create docs and projects';

  if (serviceStatus.missing.length > 0) {
    completionMessage += '\n\nNote: Some Advanced Services were not enabled. Check the log for details.';
  }

  ui.alert('Setup Complete!', completionMessage, ui.ButtonSet.OK);
}

/**
 * Runs setup in standalone mode using Script Properties.
 */
function runStandaloneSetup(props) {
  Logger.log('=== STANDALONE SETUP MODE ===');

  // Step 1: Check for required SPREADSHEET_ID
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    Logger.log('ERROR: SPREADSHEET_ID not set in Script Properties.');
    Logger.log('');
    Logger.log('To fix this:');
    Logger.log('1. Go to Project Settings (gear icon)');
    Logger.log('2. Scroll to Script Properties');
    Logger.log('3. Click "Add Script Property"');
    Logger.log('4. Property: SPREADSHEET_ID');
    Logger.log('5. Value: Your Google Sheet ID (from the URL)');
    Logger.log('');
    Logger.log('Optional properties you can also set:');
    Logger.log('- TODOIST_API_TOKEN: Your Todoist API token');
    Logger.log('- CLAUDE_API_KEY: Your Anthropic Claude API key');
    Logger.log('- USER_NAME: Your name for email signatures (default: "Team")');
    Logger.log('- BUSINESS_HOURS_START: Start hour 0-23 (default: 8)');
    Logger.log('- BUSINESS_HOURS_END: End hour 0-24 (default: 18)');
    Logger.log('- DOC_NAME_TEMPLATE: Doc naming template (default: "Client Notes - {client_name}")');
    throw new Error('SPREADSHEET_ID is required. See logs for instructions.');
  }

  Logger.log(`Using SPREADSHEET_ID: ${spreadsheetId}`);

  // Step 2: Apply defaults for optional properties if not set
  if (!props.getProperty('USER_NAME')) {
    props.setProperty('USER_NAME', 'Team');
    Logger.log('Set default USER_NAME: Team');
  } else {
    Logger.log(`Using USER_NAME: ${props.getProperty('USER_NAME')}`);
  }

  if (!props.getProperty('BUSINESS_HOURS_START')) {
    props.setProperty('BUSINESS_HOURS_START', '8');
    Logger.log('Set default BUSINESS_HOURS_START: 8');
  } else {
    Logger.log(`Using BUSINESS_HOURS_START: ${props.getProperty('BUSINESS_HOURS_START')}`);
  }

  if (!props.getProperty('BUSINESS_HOURS_END')) {
    props.setProperty('BUSINESS_HOURS_END', '18');
    Logger.log('Set default BUSINESS_HOURS_END: 18');
  } else {
    Logger.log(`Using BUSINESS_HOURS_END: ${props.getProperty('BUSINESS_HOURS_END')}`);
  }

  if (!props.getProperty('DOC_NAME_TEMPLATE')) {
    props.setProperty('DOC_NAME_TEMPLATE', 'Client Notes - {client_name}');
    Logger.log('Set default DOC_NAME_TEMPLATE: Client Notes - {client_name}');
  } else {
    Logger.log(`Using DOC_NAME_TEMPLATE: ${props.getProperty('DOC_NAME_TEMPLATE')}`);
  }

  // Log optional integrations status
  if (props.getProperty('TODOIST_API_TOKEN')) {
    Logger.log('TODOIST_API_TOKEN: Set');
  } else {
    Logger.log('TODOIST_API_TOKEN: Not set (Todoist integration disabled)');
  }

  if (props.getProperty('CLAUDE_API_KEY')) {
    Logger.log('CLAUDE_API_KEY: Set');
  } else {
    Logger.log('CLAUDE_API_KEY: Not set (AI agenda generation disabled)');
  }

  // Step 3: Check Advanced Services
  Logger.log('');
  Logger.log('Checking Advanced Services...');
  const serviceStatus = checkAdvancedServices();

  if (serviceStatus.missing.length > 0) {
    Logger.log('WARNING: Missing Advanced Services:');
    serviceStatus.missing.forEach(s => Logger.log(s));
    Logger.log('');
    Logger.log('To enable them:');
    Logger.log('1. Click "Services" (+ icon) in the left sidebar');
    Logger.log('2. Find each service listed above');
    Logger.log('3. Click "Add" for each one');
    Logger.log('');
    Logger.log('Continuing setup - some features may not work.');
  } else {
    Logger.log('All Advanced Services are enabled.');
  }

  // Step 4: Open the spreadsheet and create sheets
  Logger.log('');
  Logger.log('Creating sheets...');
  const ss = SpreadsheetApp.openById(spreadsheetId);
  createAllSheets(ss);
  Logger.log('Sheets created.');

  // Step 5: Set up triggers
  Logger.log('');
  Logger.log('Setting up triggers...');
  setupAllTriggers();
  Logger.log('Triggers created.');

  // Step 6: Sync Drive folders
  Logger.log('');
  Logger.log('Syncing Google Drive folders...');
  try {
    syncDriveFolders();
    Logger.log('Folder sync complete.');
  } catch (e) {
    Logger.log(`Folder sync error: ${e.message}`);
  }

  // Done!
  Logger.log('');
  Logger.log('=== SETUP COMPLETE ===');
  Logger.log('');
  Logger.log('Next steps:');
  Logger.log('1. Add clients to the Client_Registry sheet');
  Logger.log('2. Select a folder from the docs_folder_path dropdown');
  Logger.log('3. The system will automatically create docs and projects');
  Logger.log('');
  Logger.log(`Spreadsheet URL: https://docs.google.com/spreadsheets/d/${spreadsheetId}`);

  if (serviceStatus.missing.length > 0) {
    Logger.log('');
    Logger.log('NOTE: Some Advanced Services were not enabled. See warnings above.');
  }
}

/**
 * Checks if required Advanced Services are enabled.
 *
 * @returns {Object} Object with 'available' and 'missing' arrays
 */
function checkAdvancedServices() {
  const status = {
    available: [],
    missing: []
  };

  // Check Calendar API
  try {
    if (typeof Calendar !== 'undefined' && Calendar.Events) {
      status.available.push('Calendar API - for attaching docs to calendar events');
    } else {
      status.missing.push('• Calendar API - for attaching docs to calendar events');
    }
  } catch (e) {
    status.missing.push('• Calendar API - for attaching docs to calendar events');
  }

  // Check Gmail API
  try {
    if (typeof Gmail !== 'undefined' && Gmail.Users) {
      status.available.push('Gmail API - for creating email filters');
    } else {
      status.missing.push('• Gmail API - for creating email filters');
    }
  } catch (e) {
    status.missing.push('• Gmail API - for creating email filters');
  }

  Logger.log(`Advanced Services - Available: ${status.available.length}, Missing: ${status.missing.length}`);
  return status;
}

/**
 * Creates all required sheets with headers.
 *
 * @param {Spreadsheet} ss - The spreadsheet object
 */
function createAllSheets(ss) {
  // Client_Registry - with checkbox column for setup_complete
  createClientRegistrySheet(ss);

  // Generated_Agendas
  createSheetWithHeaders(ss, 'Generated_Agendas', [
    'event_id', 'event_title', 'client_name', 'generated_timestamp'
  ]);

  // Processing_Log
  createSheetWithHeaders(ss, 'Processing_Log', [
    'timestamp', 'action_type', 'client_name', 'details', 'status'
  ]);

  // Unmatched
  createSheetWithHeaders(ss, 'Unmatched', [
    'timestamp', 'item_type', 'item_details', 'participant_emails', 'manually_resolved'
  ]);

  // Folders
  createSheetWithHeaders(ss, 'Folders', [
    'folder_path', 'folder_id', 'folder_url'
  ]);

  // Hidden Prompts sheet
  createPromptsSheet(ss);

  // Hidden JSON Formats sheet (API documentation)
  createJsonFormatsSheet(ss);

  Logger.log('All sheets created.');
}

/**
 * Creates the Client_Registry sheet with checkbox column.
 *
 * @param {Spreadsheet} ss - The spreadsheet object
 */
function createClientRegistrySheet(ss) {
  const sheetName = 'Client_Registry';
  const headers = [
    'client_name', 'contact_emails', 'docs_folder_path',
    'setup_complete', 'google_doc_url', 'todoist_project_id'
  ];

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Add checkbox data validation to setup_complete column (column 4)
    // Apply to rows 2-1000 to cover future entries
    const checkboxColumn = 4;
    const checkboxRange = sheet.getRange(2, checkboxColumn, 999, 1);
    const checkboxRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    checkboxRange.setDataValidation(checkboxRule);

    Logger.log(`Created sheet: ${sheetName} with checkbox column`);
  }
}

/**
 * Creates a sheet with headers if it doesn't exist.
 *
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {string} sheetName - Name of the sheet
 * @param {string[]} headers - Column headers
 */
function createSheetWithHeaders(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log(`Created sheet: ${sheetName}`);
  }
}

/**
 * Sets up all automated triggers.
 */
function setupAllTriggers() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // Folder sync - 5:30 AM
  ScriptApp.newTrigger('runFolderSync')
    .timeBased()
    .atHour(5)
    .nearMinute(30)
    .everyDays(1)
    .create();

  // Label/filter sync - 6:00 AM
  ScriptApp.newTrigger('runLabelAndFilterCreation')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  // Client onboarding - 6:30 AM
  ScriptApp.newTrigger('runClientOnboarding')
    .timeBased()
    .atHour(6)
    .nearMinute(30)
    .everyDays(1)
    .create();

  // Meeting summary monitor - every 10 min
  ScriptApp.newTrigger('runSentMeetingSummaryMonitor')
    .timeBased()
    .everyMinutes(10)
    .create();

  // Agenda generation - hourly
  ScriptApp.newTrigger('runAgendaGeneration')
    .timeBased()
    .everyHours(1)
    .create();

  // Daily outlook - 7:00 AM
  ScriptApp.newTrigger('runDailyOutlook')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  // Weekly outlook - Monday 7:00 AM
  ScriptApp.newTrigger('runWeeklyOutlook')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();

  // Spreadsheet-based triggers
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (spreadsheetId) {
    const ss = SpreadsheetApp.openById(spreadsheetId);

    // onEdit trigger for checkbox-based client setup
    ScriptApp.newTrigger('onClientRegistryEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
    Logger.log('Created onEdit trigger for client setup');

    // onOpen trigger for custom menu
    ScriptApp.newTrigger('onOpen')
      .forSpreadsheet(ss)
      .onOpen()
      .create();
    Logger.log('Created onOpen trigger for menu');
  }

  Logger.log('All triggers created.');
}

/**
 * Handles edits to the Client_Registry sheet.
 * When setup_complete checkbox is checked, creates all client resources.
 *
 * @param {Object} e - The edit event object
 */
function onClientRegistryEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;

    // Only process edits to Client_Registry sheet
    if (sheet.getName() !== 'Client_Registry') {
      return;
    }

    // Only process edits to the setup_complete column (column 4)
    const setupCompleteCol = 4;
    if (range.getColumn() !== setupCompleteCol) {
      return;
    }

    // Only process if checkbox was checked (value is TRUE)
    if (e.value !== 'TRUE' && e.value !== true) {
      return;
    }

    const row = range.getRow();

    // Skip header row
    if (row === 1) {
      return;
    }

    // Get the row data
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Map to object
    const client = {};
    headers.forEach((header, index) => {
      client[header] = rowData[index];
    });

    // Check if already set up (has google_doc_url)
    if (client.google_doc_url) {
      Logger.log(`Client ${client.client_name} already set up, skipping`);
      return;
    }

    // Validate required fields
    if (!client.client_name) {
      Logger.log('Cannot set up client: client_name is required');
      // Uncheck the box
      range.setValue(false);
      return;
    }

    Logger.log(`Setting up client: ${client.client_name}`);

    // Get column indices for updating
    const colIndex = {
      google_doc_url: headers.indexOf('google_doc_url') + 1,
      todoist_project_id: headers.indexOf('todoist_project_id') + 1
    };

    // 1. Create Google Doc
    const folderId = client.docs_folder_path ? getFolderIdFromPath(client.docs_folder_path) : null;
    const docUrl = createClientDoc(client.client_name, folderId);
    if (docUrl && colIndex.google_doc_url > 0) {
      sheet.getRange(row, colIndex.google_doc_url).setValue(docUrl);
      Logger.log(`Created Google Doc: ${docUrl}`);
    }

    // 2. Create Todoist project
    const projectId = createTodoistProject(client.client_name);
    if (projectId && colIndex.todoist_project_id > 0) {
      sheet.getRange(row, colIndex.todoist_project_id).setValue(projectId);
      Logger.log(`Created Todoist project: ${projectId}`);
    }

    // 3. Create Gmail labels
    syncClientLabels({
      client_name: client.client_name,
      contact_emails: client.contact_emails
    });
    Logger.log(`Created Gmail labels for: ${client.client_name}`);

    // 4. Log success
    logProcessing('CLIENT_SETUP', client.client_name, 'Client setup complete via checkbox', 'success');

    Logger.log(`Client setup complete: ${client.client_name}`);

  } catch (error) {
    Logger.log(`Error in onClientRegistryEdit: ${error.message}`);
    logProcessing('CLIENT_SETUP_ERROR', null, error.message, 'error');
  }
}

// ============================================================================
// MIGRATION FROM EXISTING SYSTEM
// ============================================================================

/**
 * Scans for existing Gmail labels and Todoist projects to pre-populate clients.
 * Run this once during migration from an old system.
 *
 * Discovers:
 * - Gmail labels matching "Client: *" pattern
 * - Todoist projects
 * - Attempts to match them by name similarity
 *
 * @returns {Object} Migration results with discovered clients
 */
function migrateFromExistingSystem() {
  Logger.log('=== MIGRATION: Scanning for existing clients ===');

  const discovered = {
    gmailLabels: [],
    todoistProjects: [],
    matchedClients: [],
    unmatchedLabels: [],
    unmatchedProjects: []
  };

  // 1. Scan Gmail labels for "Client: *" pattern
  Logger.log('Scanning Gmail labels...');
  try {
    const allLabels = GmailApp.getUserLabels();

    for (const label of allLabels) {
      const labelName = label.getName();

      // Match "Client: ClientName" but not sublabels like "Client: Name/Meeting Summaries"
      if (labelName.startsWith('Client: ') && !labelName.includes('/')) {
        const clientName = labelName.replace('Client: ', '');
        discovered.gmailLabels.push({
          labelName: labelName,
          clientName: clientName
        });
        Logger.log(`Found label: ${labelName} -> Client: ${clientName}`);
      }
    }
  } catch (error) {
    Logger.log(`Failed to scan Gmail labels: ${error.message}`);
  }

  // 2. Scan Todoist projects
  Logger.log('Scanning Todoist projects...');
  const todoistToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (todoistToken) {
    try {
      const response = UrlFetchApp.fetch('https://api.todoist.com/rest/v2/projects', {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${todoistToken}`
        }
      });

      const projects = JSON.parse(response.getContentText());

      for (const project of projects) {
        // Skip inbox and other system projects
        if (project.is_inbox_project) continue;

        discovered.todoistProjects.push({
          projectId: project.id,
          projectName: project.name
        });
        Logger.log(`Found Todoist project: ${project.name} (${project.id})`);
      }
    } catch (error) {
      Logger.log(`Failed to scan Todoist projects: ${error.message}`);
    }
  } else {
    Logger.log('No Todoist API token - skipping Todoist scan');
  }

  // 3. Match Gmail labels to Todoist projects by name similarity
  Logger.log('Matching labels to projects...');

  for (const labelInfo of discovered.gmailLabels) {
    const clientName = labelInfo.clientName;
    let matched = false;

    // Try exact match first
    for (const projectInfo of discovered.todoistProjects) {
      if (projectInfo.projectName.toLowerCase() === clientName.toLowerCase()) {
        discovered.matchedClients.push({
          client_name: clientName,
          todoist_project_id: projectInfo.projectId,
          source: 'exact_match'
        });
        matched = true;
        Logger.log(`Exact match: ${clientName} -> Project ${projectInfo.projectId}`);
        break;
      }
    }

    // Try partial match if no exact match
    if (!matched) {
      for (const projectInfo of discovered.todoistProjects) {
        const labelLower = clientName.toLowerCase();
        const projectLower = projectInfo.projectName.toLowerCase();

        if (labelLower.includes(projectLower) || projectLower.includes(labelLower)) {
          discovered.matchedClients.push({
            client_name: clientName,
            todoist_project_id: projectInfo.projectId,
            source: 'partial_match',
            todoist_name: projectInfo.projectName
          });
          matched = true;
          Logger.log(`Partial match: ${clientName} -> Project ${projectInfo.projectName} (${projectInfo.projectId})`);
          break;
        }
      }
    }

    if (!matched) {
      discovered.unmatchedLabels.push(labelInfo);
      // Still add as a client, just without Todoist project
      discovered.matchedClients.push({
        client_name: clientName,
        todoist_project_id: '',
        source: 'label_only'
      });
    }
  }

  // Find Todoist projects without matching labels
  for (const projectInfo of discovered.todoistProjects) {
    const hasMatch = discovered.matchedClients.some(
      c => c.todoist_project_id === projectInfo.projectId
    );
    if (!hasMatch) {
      discovered.unmatchedProjects.push(projectInfo);
    }
  }

  Logger.log('');
  Logger.log('=== MIGRATION SUMMARY ===');
  Logger.log(`Gmail labels found: ${discovered.gmailLabels.length}`);
  Logger.log(`Todoist projects found: ${discovered.todoistProjects.length}`);
  Logger.log(`Matched clients: ${discovered.matchedClients.length}`);
  Logger.log(`Unmatched labels: ${discovered.unmatchedLabels.length}`);
  Logger.log(`Unmatched projects: ${discovered.unmatchedProjects.length}`);

  return discovered;
}

/**
 * Imports discovered clients into the Client_Registry sheet.
 * Run migrateFromExistingSystem() first to see what will be imported.
 *
 * @param {boolean} dryRun - If true, only logs what would be imported without making changes
 */
function importDiscoveredClients(dryRun = false) {
  Logger.log(`=== IMPORT CLIENTS ${dryRun ? '(DRY RUN)' : ''} ===`);

  // Run discovery
  const discovered = migrateFromExistingSystem();

  if (discovered.matchedClients.length === 0) {
    Logger.log('No clients discovered to import.');
    return;
  }

  // Get existing clients to avoid duplicates
  const existingClients = getClientRegistry();
  const existingNames = existingClients.map(c => c.client_name.toLowerCase());

  // Open spreadsheet
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    Logger.log('ERROR: SPREADSHEET_ID not set');
    return;
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('Client_Registry');

  if (!sheet) {
    Logger.log('ERROR: Client_Registry sheet not found');
    return;
  }

  let imported = 0;
  let skipped = 0;

  for (const client of discovered.matchedClients) {
    // Check for duplicate
    if (existingNames.includes(client.client_name.toLowerCase())) {
      Logger.log(`SKIP (duplicate): ${client.client_name}`);
      skipped++;
      continue;
    }

    if (dryRun) {
      Logger.log(`WOULD IMPORT: ${client.client_name} (Todoist: ${client.todoist_project_id || 'none'})`);
    } else {
      // Add to sheet: client_name, contact_emails, docs_folder_path, setup_complete, google_doc_url, todoist_project_id
      sheet.appendRow([
        client.client_name,  // client_name
        '',                   // contact_emails (to be filled manually)
        '',                   // docs_folder_path (to be selected)
        false,                // setup_complete (unchecked - user needs to complete setup)
        '',                   // google_doc_url (to be created)
        client.todoist_project_id || ''  // todoist_project_id (if matched)
      ]);
      Logger.log(`IMPORTED: ${client.client_name}`);
    }
    imported++;
  }

  Logger.log('');
  Logger.log(`=== IMPORT COMPLETE ===`);
  Logger.log(`Imported: ${imported}`);
  Logger.log(`Skipped (duplicates): ${skipped}`);

  if (!dryRun && imported > 0) {
    Logger.log('');
    Logger.log('Next steps:');
    Logger.log('1. Add contact_emails for each client');
    Logger.log('2. Select docs_folder_path from dropdown');
    Logger.log('3. Check setup_complete to create Google Docs');
  }

  // Log unmatched Todoist projects
  if (discovered.unmatchedProjects.length > 0) {
    Logger.log('');
    Logger.log('Todoist projects without matching Gmail labels:');
    for (const project of discovered.unmatchedProjects) {
      Logger.log(`  - ${project.projectName} (${project.projectId})`);
    }
    Logger.log('You may want to create Gmail labels for these or add them manually.');
  }
}

/**
 * Preview what will be imported without making changes.
 * Run this first before importDiscoveredClients().
 */
function previewMigration() {
  importDiscoveredClients(true);
}

// ============================================================================
// ON OPEN MENU
// ============================================================================

/**
 * Creates custom menu when the spreadsheet is opened.
 * This is an installable trigger - set up via setupAllTriggers().
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Client Automation')
      .addItem('Run Setup', 'SETUP_RUN_THIS_FIRST')
      .addItem('Import Existing Clients...', 'showMigrationWizard')
      .addSeparator()
      .addItem('Update Settings...', 'showSettingsEditor')
      .addItem('Adjust Prompts...', 'showPromptsEditor')
      .addSeparator()
      .addItem('Sync Drive Folders', 'syncDriveFolders')
      .addItem('Sync Labels & Filters', 'runLabelAndFilterCreation')
      .addSeparator()
      .addItem('View Processing Log', 'showProcessingLog')
      .addSeparator()
      .addItem('Disable Automation...', 'disableAutomationWithConfirmation')
      .addToUi();
  } catch (e) {
    // Not bound to a spreadsheet - skip menu creation
    Logger.log('onOpen: Not bound to spreadsheet, skipping menu');
  }
}

/**
 * Shows the processing log to the user.
 */
function showProcessingLog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Processing_Log');

  if (sheet) {
    ss.setActiveSheet(sheet);
    ui.alert('Processing Log', 'Showing the Processing_Log sheet with recent activity.', ui.ButtonSet.OK);
  } else {
    ui.alert('Not Found', 'Processing_Log sheet not found. Run setup first.', ui.ButtonSet.OK);
  }
}

// ============================================================================
// DISABLE AUTOMATION
// ============================================================================

/**
 * Disables all automation by removing all triggers.
 * Requires user confirmation before executing.
 */
function disableAutomationWithConfirmation() {
  const ui = SpreadsheetApp.getUi();

  // Get current trigger count
  const triggers = ScriptApp.getProjectTriggers();
  const triggerCount = triggers.length;

  if (triggerCount === 0) {
    ui.alert('No Triggers', 'Automation is already disabled. No triggers are currently active.', ui.ButtonSet.OK);
    return;
  }

  // Show confirmation dialog
  const response = ui.alert(
    'Disable Automation?',
    `This will remove all ${triggerCount} active trigger(s) and stop all automated processes:\n\n` +
    '• Daily/Weekly outlook emails\n' +
    '• Agenda generation\n' +
    '• Meeting summary monitoring\n' +
    '• Label/filter sync\n' +
    '• Folder sync\n' +
    '• Client onboarding\n\n' +
    'You can re-enable by running Setup again.\n\n' +
    'Are you sure you want to disable automation?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Cancelled', 'Automation was not disabled.', ui.ButtonSet.OK);
    return;
  }

  // Remove all triggers
  try {
    for (const trigger of triggers) {
      ScriptApp.deleteTrigger(trigger);
    }

    logProcessing('AUTOMATION_DISABLED', null, `Removed ${triggerCount} triggers`, 'warning');

    ui.alert(
      'Automation Disabled',
      `Successfully removed ${triggerCount} trigger(s).\n\n` +
      'All automated processes have been stopped.\n\n' +
      'To re-enable automation, go to:\n' +
      'Client Automation > Run Setup',
      ui.ButtonSet.OK
    );

    Logger.log(`Disabled automation: removed ${triggerCount} triggers`);

  } catch (error) {
    ui.alert('Error', `Failed to disable automation: ${error.message}`, ui.ButtonSet.OK);
    Logger.log(`Error disabling automation: ${error.message}`);
  }
}

// ============================================================================
// SETTINGS EDITOR
// ============================================================================

/**
 * Shows the settings editor dialog.
 */
function showSettingsEditor() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsEditor')
    .setWidth(450)
    .setHeight(550)
    .setTitle('Update Settings');

  SpreadsheetApp.getUi().showModalDialog(html, 'Update Settings');
}

/**
 * Gets current settings for the editor.
 * Masks sensitive values for display.
 *
 * @returns {Object} Current settings
 */
function getSettingsForEditor() {
  const props = PropertiesService.getScriptProperties();

  return {
    TODOIST_API_TOKEN: props.getProperty('TODOIST_API_TOKEN') || '',
    CLAUDE_API_KEY: props.getProperty('CLAUDE_API_KEY') || '',
    USER_NAME: props.getProperty('USER_NAME') || '',
    BUSINESS_HOURS_START: props.getProperty('BUSINESS_HOURS_START') || '8',
    BUSINESS_HOURS_END: props.getProperty('BUSINESS_HOURS_END') || '18',
    DOC_NAME_TEMPLATE: props.getProperty('DOC_NAME_TEMPLATE') || 'Client Notes - {client_name}'
  };
}

/**
 * Saves settings from the editor.
 * Only updates non-empty values.
 *
 * @param {Object} settings - Settings object from the editor
 * @returns {Object} Result with success status
 */
function saveSettingsFromEditor(settings) {
  const props = PropertiesService.getScriptProperties();

  // Only update settings that have values (don't clear existing ones if left blank)
  if (settings.TODOIST_API_TOKEN) {
    props.setProperty('TODOIST_API_TOKEN', settings.TODOIST_API_TOKEN);
  }

  if (settings.CLAUDE_API_KEY) {
    props.setProperty('CLAUDE_API_KEY', settings.CLAUDE_API_KEY);
  }

  if (settings.USER_NAME) {
    props.setProperty('USER_NAME', settings.USER_NAME);
  }

  if (settings.BUSINESS_HOURS_START) {
    props.setProperty('BUSINESS_HOURS_START', settings.BUSINESS_HOURS_START);
  }

  if (settings.BUSINESS_HOURS_END) {
    props.setProperty('BUSINESS_HOURS_END', settings.BUSINESS_HOURS_END);
  }

  if (settings.DOC_NAME_TEMPLATE) {
    props.setProperty('DOC_NAME_TEMPLATE', settings.DOC_NAME_TEMPLATE);
  }

  Logger.log('Settings updated via editor');

  return { success: true };
}

// ============================================================================
// MIGRATION WIZARD
// ============================================================================

/**
 * Shows the migration wizard dialog.
 */
function showMigrationWizard() {
  const html = HtmlService.createHtmlOutputFromFile('MigrationWizard')
    .setWidth(500)
    .setHeight(600)
    .setTitle('Import Existing Clients');

  SpreadsheetApp.getUi().showModalDialog(html, 'Import Existing Clients');
}

/**
 * Scans for existing Gmail labels and Todoist projects.
 * Called by the migration wizard HTML.
 *
 * @returns {Object} Discovered data for the wizard
 */
function scanForMigration() {
  const discovered = {
    gmailLabels: [],
    todoistProjects: []
  };

  // Scan Gmail labels for "Client: *" pattern
  try {
    const allLabels = GmailApp.getUserLabels();

    for (const label of allLabels) {
      const labelName = label.getName();

      // Match "Client: ClientName" but not sublabels
      if (labelName.startsWith('Client: ') && !labelName.includes('/')) {
        const clientName = labelName.replace('Client: ', '');
        discovered.gmailLabels.push({
          labelName: labelName,
          clientName: clientName
        });
      }
    }
  } catch (error) {
    Logger.log(`Failed to scan Gmail labels: ${error.message}`);
  }

  // Scan Todoist projects
  const todoistToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (todoistToken) {
    try {
      const response = UrlFetchApp.fetch('https://api.todoist.com/rest/v2/projects', {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${todoistToken}`
        }
      });

      const projects = JSON.parse(response.getContentText());

      for (const project of projects) {
        // Skip inbox and other system projects
        if (project.is_inbox_project) continue;

        discovered.todoistProjects.push({
          projectId: project.id,
          projectName: project.name
        });
      }
    } catch (error) {
      Logger.log(`Failed to scan Todoist projects: ${error.message}`);
    }
  }

  return discovered;
}

/**
 * Gets folders for the wizard dropdown.
 *
 * @returns {Object[]} Array of folder objects with folder_path
 */
function getFoldersForWizard() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) return [];

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('Folders');

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const folders = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        folders.push({
          folder_path: data[i][0],
          folder_id: data[i][1]
        });
      }
    }

    return folders;
  } catch (e) {
    Logger.log(`Error getting folders: ${e.message}`);
    return [];
  }
}

/**
 * Searches Google Docs by name.
 *
 * @param {string} query - Search query
 * @returns {Object[]} Array of matching docs with name and url
 */
function searchGoogleDocs(query) {
  if (!query || query.length < 2) return [];

  try {
    // Search for Google Docs containing the query
    const files = DriveApp.searchFiles(
      `mimeType='application/vnd.google-apps.document' and title contains '${query.replace(/'/g, "\\'")}'`
    );

    const results = [];
    let count = 0;
    const maxResults = 20;

    while (files.hasNext() && count < maxResults) {
      const file = files.next();
      results.push({
        name: file.getName(),
        url: file.getUrl(),
        id: file.getId()
      });
      count++;
    }

    return results;
  } catch (e) {
    Logger.log(`Error searching docs: ${e.message}`);
    return [];
  }
}

/**
 * Searches Gmail for contacts by name or email.
 * Extracts unique email addresses from sent emails.
 *
 * @param {string} query - Search query (name or email)
 * @returns {Object[]} Array of contact objects with name and email
 */
function searchGmailContacts(query) {
  if (!query || query.length < 2) return [];

  try {
    const results = [];
    const seenEmails = new Set();
    const maxResults = 15;
    const queryLower = query.toLowerCase();

    // Helper function to parse and add contacts from email headers
    function parseAndAddContacts(headerValue) {
      if (!headerValue) return;

      const recipients = headerValue.split(',');
      for (const recipient of recipients) {
        if (results.length >= maxResults) return;

        const trimmed = recipient.trim();
        if (!trimmed) continue;

        // Parse "Name <email>" or just "email" format
        const match = trimmed.match(/^(?:([^<]+)\s*)?<?([^\s<>]+@[^\s<>]+)>?$/);
        if (match) {
          const name = match[1] ? match[1].trim().replace(/"/g, '') : '';
          const email = match[2].toLowerCase();

          // Skip if already seen
          if (seenEmails.has(email)) continue;

          // Match if query is found in email OR name (partial match)
          if (email.includes(queryLower) || name.toLowerCase().includes(queryLower)) {
            seenEmails.add(email);
            results.push({ name: name, email: email });
          }
        }
      }
    }

    // Strategy 1: Search sent emails with to: operator
    try {
      const toThreads = GmailApp.search(`in:sent to:${query}`, 0, 30);
      for (const thread of toThreads) {
        if (results.length >= maxResults) break;
        const messages = thread.getMessages();
        for (const message of messages) {
          if (results.length >= maxResults) break;
          parseAndAddContacts(message.getTo());
          parseAndAddContacts(message.getCc());
        }
      }
    } catch (e) { /* continue to next strategy */ }

    // Strategy 2: Search sent emails with the query as a general keyword
    // This catches names that to: operator misses
    if (results.length < maxResults) {
      try {
        const keywordThreads = GmailApp.search(`in:sent ${query}`, 0, 30);
        for (const thread of keywordThreads) {
          if (results.length >= maxResults) break;
          const messages = thread.getMessages();
          for (const message of messages) {
            if (results.length >= maxResults) break;
            parseAndAddContacts(message.getTo());
            parseAndAddContacts(message.getCc());
          }
        }
      } catch (e) { /* continue to next strategy */ }
    }

    // Strategy 3: Search received emails with from: operator
    if (results.length < maxResults) {
      try {
        const fromThreads = GmailApp.search(`from:${query}`, 0, 30);
        for (const thread of fromThreads) {
          if (results.length >= maxResults) break;
          const messages = thread.getMessages();
          for (const message of messages) {
            if (results.length >= maxResults) break;
            const from = message.getFrom();
            parseAndAddContacts(from);
          }
        }
      } catch (e) { /* continue to next strategy */ }
    }

    // Strategy 4: Search received emails with keyword (catches names in body/subject)
    if (results.length < maxResults) {
      try {
        const receivedThreads = GmailApp.search(`${query}`, 0, 30);
        for (const thread of receivedThreads) {
          if (results.length >= maxResults) break;
          const messages = thread.getMessages();
          for (const message of messages) {
            if (results.length >= maxResults) break;
            parseAndAddContacts(message.getFrom());
          }
        }
      } catch (e) { /* continue */ }
    }

    return results;

  } catch (e) {
    Logger.log(`Error searching contacts: ${e.message}`);
    return [];
  }
}

/**
 * Imports clients from the migration wizard.
 *
 * @param {Object[]} importData - Array of client import configurations
 * @returns {Object} Result with imported count
 */
function importClientsFromWizard(importData) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID not set');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('Client_Registry');

  if (!sheet) {
    throw new Error('Client_Registry sheet not found');
  }

  // Get existing clients to avoid duplicates
  const existingData = sheet.getDataRange().getValues();
  const existingNames = [];
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) {
      existingNames.push(existingData[i][0].toLowerCase());
    }
  }

  let imported = 0;

  for (const client of importData) {
    // Skip duplicates
    if (existingNames.includes(client.client_name.toLowerCase())) {
      Logger.log(`Skip duplicate: ${client.client_name}`);
      continue;
    }

    let googleDocUrl = '';
    let todoistProjectId = client.todoist_project_id || '';

    // Handle doc creation/linking
    if (client.doc_mode === 'existing' && client.existing_doc_url) {
      googleDocUrl = client.existing_doc_url;
    } else if (client.doc_mode === 'create') {
      // Create new doc
      const folderId = client.create_folder ? getFolderIdFromPath(client.create_folder) : null;
      googleDocUrl = createClientDoc(client.client_name, folderId) || '';
    }

    // Handle Todoist project creation
    if (client.create_todoist) {
      todoistProjectId = createTodoistProject(client.client_name) || '';
    }

    // Add to sheet: client_name, contact_emails, docs_folder_path, setup_complete, google_doc_url, todoist_project_id
    const contactEmails = client.contact_emails || '';
    sheet.appendRow([
      client.client_name,
      contactEmails,
      client.create_folder || '',
      true,  // setup_complete - checked since we're setting up now
      googleDocUrl,
      todoistProjectId
    ]);

    // Create Gmail labels and filters
    if (client.create_gmail_label || !client.gmail_label) {
      syncClientLabels({
        client_name: client.client_name,
        contact_emails: contactEmails
      });
    }

    Logger.log(`Imported: ${client.client_name}`);
    imported++;
  }

  logProcessing('MIGRATION_WIZARD', null, `Imported ${imported} clients via wizard`, 'success');

  return { imported: imported };
}
