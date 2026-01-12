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
  // Client_Registry
  createSheetWithHeaders(ss, 'Client_Registry', [
    'client_id', 'client_name', 'email_domains', 'contact_emails',
    'docs_folder_path', 'google_doc_url', 'todoist_project_id'
  ]);

  // Generated_Agendas
  createSheetWithHeaders(ss, 'Generated_Agendas', [
    'event_id', 'event_title', 'client_id', 'generated_timestamp'
  ]);

  // Processing_Log
  createSheetWithHeaders(ss, 'Processing_Log', [
    'timestamp', 'action_type', 'client_id', 'details', 'status'
  ]);

  // Unmatched
  createSheetWithHeaders(ss, 'Unmatched', [
    'timestamp', 'item_type', 'item_details', 'participant_emails', 'manually_resolved'
  ]);

  // Folders
  createSheetWithHeaders(ss, 'Folders', [
    'folder_path', 'folder_id', 'folder_url'
  ]);

  Logger.log('All sheets created.');
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

  Logger.log('All triggers created.');
}
