/**
 * ClientIdentification.gs - Client Matching Logic
 *
 * This module handles client identification from various inputs including
 * meeting participant emails, email sender/recipient addresses, and calendar
 * event guest lists.
 *
 * Matching Order:
 * 1. Exact match against contact_emails
 * 2. Domain match against email_domains
 *
 * The first match is used. If no match is found, the item is logged to the
 * Unmatched sheet and processing stops.
 */

// ============================================================================
// MAIN CLIENT IDENTIFICATION
// ============================================================================

/**
 * Identifies a client from a list of email addresses.
 *
 * @param {string[]} emails - Array of email addresses to match against
 * @returns {Object|null} Client object if found, null otherwise
 *
 * @example
 * const client = identifyClient(['john@acmecorp.com', 'sarah@gmail.com']);
 * // Returns: { client_id: '1', client_name: 'Acme Corp', ... }
 */
function identifyClient(emails) {
  if (!emails || emails.length === 0) {
    return null;
  }

  // Normalize emails to lowercase
  const normalizedEmails = emails.map(email => email.toLowerCase().trim());

  // Get all clients from registry
  const clients = getClientRegistry();

  // First pass: Check for exact contact email match
  for (const client of clients) {
    const contactEmails = parseCommaSeparatedList(client.contact_emails);

    for (const email of normalizedEmails) {
      if (contactEmails.includes(email)) {
        Logger.log(`Client identified by contact email: ${client.client_name} (${email})`);
        return client;
      }
    }
  }

  // Second pass: Check for domain match
  for (const client of clients) {
    const emailDomains = parseCommaSeparatedList(client.email_domains);

    for (const email of normalizedEmails) {
      const domain = extractDomain(email);
      if (domain && emailDomains.includes(domain)) {
        Logger.log(`Client identified by domain: ${client.client_name} (${domain})`);
        return client;
      }
    }
  }

  // No match found
  Logger.log(`No client match found for emails: ${normalizedEmails.join(', ')}`);
  return null;
}

/**
 * Identifies a client from meeting participants.
 *
 * @param {Object[]} participants - Array of participant objects with name and email
 * @returns {Object|null} Client object if found, null otherwise
 */
function identifyClientFromParticipants(participants) {
  if (!participants || participants.length === 0) {
    return null;
  }

  const emails = participants
    .map(p => p.email)
    .filter(email => email && email.length > 0);

  return identifyClient(emails);
}

/**
 * Identifies a client from calendar event guests.
 *
 * @param {CalendarEvent} event - Google Calendar event object
 * @returns {Object|null} Client object if found, null otherwise
 */
function identifyClientFromCalendarEvent(event) {
  const guestList = event.getGuestList();
  const emails = guestList.map(guest => guest.getEmail());

  // Also check the event organizer if not the current user
  const creatorEmail = event.getCreators()[0];
  if (creatorEmail && creatorEmail !== Session.getActiveUser().getEmail()) {
    emails.push(creatorEmail);
  }

  return identifyClient(emails);
}

/**
 * Identifies a client from an email message.
 *
 * @param {GmailMessage} message - Gmail message object
 * @returns {Object|null} Client object if found, null otherwise
 */
function identifyClientFromEmail(message) {
  const emails = [];

  // Get sender
  const from = message.getFrom();
  const fromEmail = extractEmailFromHeader(from);
  if (fromEmail) {
    emails.push(fromEmail);
  }

  // Get recipients
  const to = message.getTo();
  const toEmails = extractEmailsFromHeader(to);
  emails.push(...toEmails);

  // Get CC recipients
  const cc = message.getCc();
  if (cc) {
    const ccEmails = extractEmailsFromHeader(cc);
    emails.push(...ccEmails);
  }

  return identifyClient(emails);
}

// ============================================================================
// CLIENT REGISTRY ACCESS
// ============================================================================

/**
 * Retrieves all clients from the Client_Registry sheet.
 *
 * @returns {Object[]} Array of client objects
 */
function getClientRegistry() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CLIENT_REGISTRY);

  if (!sheet) {
    Logger.log('Client_Registry sheet not found');
    return [];
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return []; // Only header row or empty
  }

  const headers = data[0];
  const clients = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const client = {};

    headers.forEach((header, index) => {
      client[header] = row[index] || '';
    });

    // Only include clients with at least a client_id and client_name
    if (client.client_id && client.client_name) {
      clients.push(client);
    }
  }

  return clients;
}

/**
 * Retrieves a specific client by ID.
 *
 * @param {string} clientId - The client ID to look up
 * @returns {Object|null} Client object if found, null otherwise
 */
function getClientById(clientId) {
  const clients = getClientRegistry();
  return clients.find(c => c.client_id === clientId) || null;
}

/**
 * Retrieves a specific client by name.
 *
 * @param {string} clientName - The client name to look up
 * @returns {Object|null} Client object if found, null otherwise
 */
function getClientByName(clientName) {
  const clients = getClientRegistry();
  return clients.find(c => c.client_name.toLowerCase() === clientName.toLowerCase()) || null;
}

// ============================================================================
// UNMATCHED LOGGING
// ============================================================================

/**
 * Logs an unmatched item to the Unmatched sheet.
 *
 * @param {string} itemType - Type of item ('meeting' or 'email')
 * @param {string} itemDetails - Description or details of the item
 * @param {string[]} participantEmails - Array of email addresses that didn't match
 */
function logUnmatched(itemType, itemDetails, participantEmails) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UNMATCHED);

  if (!sheet) {
    Logger.log('Unmatched sheet not found');
    return;
  }

  const timestamp = new Date().toISOString();
  const emailsString = Array.isArray(participantEmails)
    ? participantEmails.join(', ')
    : participantEmails;

  sheet.appendRow([
    timestamp,
    itemType,
    itemDetails,
    emailsString,
    'FALSE' // manually_resolved defaults to false
  ]);

  Logger.log(`Logged unmatched ${itemType}: ${itemDetails}`);
}

/**
 * Marks an unmatched item as resolved.
 *
 * @param {number} rowIndex - The row index in the Unmatched sheet (1-based, excluding header)
 */
function markUnmatchedResolved(rowIndex) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UNMATCHED);

  if (!sheet) {
    return;
  }

  // Row index is 1-based and accounts for header
  const actualRow = rowIndex + 1;
  const manuallyResolvedCol = 5; // Column E

  sheet.getRange(actualRow, manuallyResolvedCol).setValue('TRUE');
  Logger.log(`Marked unmatched item at row ${rowIndex} as resolved`);
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Extracts the domain from an email address.
 *
 * @param {string} email - The email address
 * @returns {string|null} The domain portion or null if invalid
 */
function extractDomain(email) {
  if (!email || typeof email !== 'string') {
    return null;
  }

  const parts = email.split('@');
  if (parts.length !== 2) {
    return null;
  }

  return parts[1].toLowerCase().trim();
}

/**
 * Parses a comma-separated list into an array of trimmed, lowercase values.
 *
 * @param {string} list - Comma-separated string
 * @returns {string[]} Array of parsed values
 */
function parseCommaSeparatedList(list) {
  if (!list || typeof list !== 'string') {
    return [];
  }

  return list
    .split(',')
    .map(item => item.toLowerCase().trim())
    .filter(item => item.length > 0);
}

/**
 * Extracts a single email address from an email header string.
 * Handles formats like "Name <email@domain.com>" or just "email@domain.com"
 *
 * @param {string} header - The email header string
 * @returns {string|null} The extracted email address or null
 */
function extractEmailFromHeader(header) {
  if (!header) {
    return null;
  }

  // Try to match email in angle brackets
  const bracketMatch = header.match(/<([^>]+)>/);
  if (bracketMatch) {
    return bracketMatch[1].toLowerCase().trim();
  }

  // Try to match a plain email
  const emailMatch = header.match(/[\w.+-]+@[\w.-]+\.\w+/);
  if (emailMatch) {
    return emailMatch[0].toLowerCase().trim();
  }

  return null;
}

/**
 * Extracts multiple email addresses from an email header string.
 * Handles comma-separated lists of emails.
 *
 * @param {string} header - The email header string
 * @returns {string[]} Array of extracted email addresses
 */
function extractEmailsFromHeader(header) {
  if (!header) {
    return [];
  }

  const emails = [];

  // Split by comma and process each part
  const parts = header.split(',');
  for (const part of parts) {
    const email = extractEmailFromHeader(part);
    if (email) {
      emails.push(email);
    }
  }

  return emails;
}

/**
 * Validates if a string is a valid email address format.
 *
 * @param {string} email - The string to validate
 * @returns {boolean} True if valid email format
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }

  const emailRegex = /^[\w.+-]+@[\w.-]+\.\w+$/;
  return emailRegex.test(email.trim());
}

/**
 * Gets the current user's email address.
 *
 * @returns {string} The current user's email
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Checks if an email belongs to the current user.
 *
 * @param {string} email - Email address to check
 * @returns {boolean} True if the email belongs to the current user
 */
function isCurrentUser(email) {
  const userEmail = getCurrentUserEmail();
  return email && email.toLowerCase().trim() === userEmail.toLowerCase().trim();
}

/**
 * Filters out the current user's email from a list of emails.
 *
 * @param {string[]} emails - Array of email addresses
 * @returns {string[]} Array without the current user's email
 */
function filterOutCurrentUser(emails) {
  const userEmail = getCurrentUserEmail().toLowerCase();
  return emails.filter(email => email.toLowerCase() !== userEmail);
}
