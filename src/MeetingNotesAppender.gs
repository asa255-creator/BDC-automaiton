/**
 * MeetingNotesAppender.gs - Append meeting notes to Google Docs.
 */

// ============================================================================
// GOOGLE DOC INTEGRATION
// ============================================================================

/**
 * Appends meeting notes to the client's running Google Doc.
 *
 * @param {GmailMessage} message - The sent meeting summary message
 * @param {Object} client - The client object
 */
function appendMeetingNotesToDoc(message, client) {
  if (!client.google_doc_url) {
    logProcessing('DOC_APPEND', client.client_name, 'No Google Doc URL configured', 'warning');
    return;
  }

  try {
    const docId = extractDocIdFromUrl(client.google_doc_url);
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Get meeting details from message
    const date = formatDate(message.getDate());

    // Add blank line before for separation from previous content
    body.appendParagraph('');

    // Add start delimiter (parseable marker)
    body.appendParagraph('═══════════════════════════════════════════════════════════');

    // Add section header with date
    body.appendParagraph(`MEETING NOTES - ${date}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    body.appendParagraph('───────────────────────────────────────────────────────────');

    // Convert email HTML to plain text and append
    const emailBody = message.getPlainBody();
    body.appendParagraph(emailBody);

    // Add end delimiter
    body.appendParagraph('───────────────────────────────────────────────────────────');
    body.appendParagraph(`END OF MEETING NOTES - ${date}`);
    body.appendParagraph('═══════════════════════════════════════════════════════════');

    // Add blank line after for separation
    body.appendParagraph('');

    doc.saveAndClose();

    logProcessing('DOC_APPEND', client.client_name, `Appended meeting notes for ${date}`, 'success');
  } catch (error) {
    logProcessing(
      'DOC_APPEND_ERROR',
      client.client_name,
      `Failed to append meeting notes: ${error.message} | URL: ${client.google_doc_url}`,
      'error'
    );
  }
}

/**
 * Extracts the document ID from a Google Docs URL.
 *
 * @param {string} url - The Google Docs URL
 * @returns {string} The document ID
 */
function extractDocIdFromUrl(url) {
  // Format: /d/DOC_ID
  const matchD = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (matchD) {
    return matchD[1];
  }

  // Format: ?id=DOC_ID or &id=DOC_ID
  const matchId = url.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (matchId) {
    return matchId[1];
  }

  // Assume it's already a doc ID if not a URL
  return url;
}

// ============================================================================
// MANUAL TEST FUNCTIONS
// ============================================================================

/**
 * TEST: Append most recent meeting summary to Google Doc.
 * Run from Apps Script editor. That's it - just appends to doc.
 */
function testLastMeetingSummary() {
  Logger.log('=== APPENDING LAST MEETING SUMMARY TO DOC ===\n');

  const allClients = getClientRegistry();
  const clients = allClients.filter(client => client.setup_complete === true);

  let mostRecentMessage = null;
  let mostRecentClient = null;
  let mostRecentDate = new Date(0);

  for (const client of clients) {
    const labelName = `Client: ${client.client_name}/Meeting Summaries`;
    const label = GmailApp.getUserLabelByName(labelName);
    if (!label) continue;

    const threads = label.getThreads(0, 1);
    if (threads.length === 0) continue;

    const messages = threads[0].getMessages();
    if (messages.length === 0) continue;

    const message = messages[0];
    if (message.getDate() > mostRecentDate) {
      mostRecentDate = message.getDate();
      mostRecentMessage = message;
      mostRecentClient = client;
    }
  }

  if (!mostRecentMessage || !mostRecentClient) {
    Logger.log('ERROR: No meeting summaries found');
    return;
  }

  Logger.log(`Client: ${mostRecentClient.client_name}`);
  Logger.log(`Subject: ${mostRecentMessage.getSubject()}`);
  Logger.log(`Doc URL: ${mostRecentClient.google_doc_url}`);

  // Append to doc
  appendMeetingNotesToDoc(mostRecentMessage, mostRecentClient);

  Logger.log('\nDone. Check the Google Doc.');
}

/**
 * Manual function to retry appending meeting notes for the most recent summary.
 * Call this from spreadsheet menu (BDC Automation > Retry Meeting Notes Append).
 * Shows dialog to select client, then tries to append their most recent meeting summary.
 */
function retryLastMeetingNotesAppend() {
  const ui = SpreadsheetApp.getUi();

  // Get all clients with setup_complete
  const allClients = getClientRegistry();
  const clients = allClients.filter(client => client.setup_complete === true);

  if (clients.length === 0) {
    ui.alert('Error', 'No clients with setup_complete found', ui.ButtonSet.OK);
    return;
  }

  // Build client selection prompt
  let clientList = 'Enter client number:\n\n';
  clients.forEach((client, index) => {
    clientList += `${index + 1}. ${client.client_name}\n`;
  });

  const response = ui.prompt('Select Client', clientList, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const clientIndex = parseInt(response.getResponseText(), 10) - 1;
  if (isNaN(clientIndex) || clientIndex < 0 || clientIndex >= clients.length) {
    ui.alert('Error', 'Invalid client number', ui.ButtonSet.OK);
    return;
  }

  const client = clients[clientIndex];

  // Check if client has a Google Doc configured
  if (!client.google_doc_url) {
    ui.alert('Error', `${client.client_name} has no Google Doc URL configured`, ui.ButtonSet.OK);
    return;
  }

  // Find the most recent message in their Meeting Summaries label
  const labelName = `Client: ${client.client_name}/Meeting Summaries`;
  const label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    ui.alert('Error', `Label "${labelName}" not found`, ui.ButtonSet.OK);
    return;
  }

  const threads = label.getThreads(0, 1);
  if (threads.length === 0) {
    ui.alert('Error', 'No threads found in Meeting Summaries label', ui.ButtonSet.OK);
    return;
  }

  const messages = threads[0].getMessages();
  if (messages.length === 0) {
    ui.alert('Error', 'No messages found in thread', ui.ButtonSet.OK);
    return;
  }

  const message = messages[0];

  // Show confirmation with details
  const confirmResult = ui.alert(
    'Confirm Retry',
    `Client: ${client.client_name}\n` +
    `Subject: ${message.getSubject()}\n` +
    `Date: ${message.getDate()}\n` +
    `Doc URL: ${client.google_doc_url}\n\n` +
    'Attempt to append meeting notes to document?',
    ui.ButtonSet.YES_NO
  );

  if (confirmResult !== ui.Button.YES) {
    return;
  }

  // Try to append
  try {
    const docId = extractDocIdFromUrl(client.google_doc_url);
    ui.alert('Debug Info', `Extracted Doc ID: ${docId}`, ui.ButtonSet.OK);

    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Get meeting details from message
    const date = formatDate(message.getDate());

    // Add content
    body.appendParagraph('');
    body.appendParagraph('═══════════════════════════════════════════════════════════');
    body.appendParagraph(`MEETING NOTES - ${date}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('───────────────────────────────────────────────────────────');

    const emailBody = message.getPlainBody();
    body.appendParagraph(emailBody);

    body.appendParagraph('───────────────────────────────────────────────────────────');
    body.appendParagraph(`END OF MEETING NOTES - ${date}`);
    body.appendParagraph('═══════════════════════════════════════════════════════════');
    body.appendParagraph('');

    doc.saveAndClose();

    ui.alert('Success', `Meeting notes appended successfully for ${client.client_name}`, ui.ButtonSet.OK);
    logProcessing('DOC_APPEND_MANUAL', client.client_name, `Manual append successful for ${date}`, 'success');

  } catch (error) {
    ui.alert(
      'Error',
      `Failed to append meeting notes:\n\n` +
      `Error: ${error.message}\n\n` +
      `Client: ${client.client_name}\n` +
      `Doc URL: ${client.google_doc_url}`,
      ui.ButtonSet.OK
    );
    logProcessing('DOC_APPEND_MANUAL', client.client_name, `Manual append failed: ${error.message} | URL: ${client.google_doc_url}`, 'error');
  }
}
