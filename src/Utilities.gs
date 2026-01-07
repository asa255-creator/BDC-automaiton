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
