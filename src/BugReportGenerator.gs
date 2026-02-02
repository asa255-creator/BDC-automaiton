// ============================================================================
// BUG REPORT GENERATOR
// ============================================================================

/**
 * Shows UI to generate a bug report for troubleshooting.
 * User can select by client, time range, or specific log entry.
 */
function showBugReportGenerator() {
  const html = HtmlService.createHtmlOutputFromFile('BugReportGeneratorUI')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Generate Bug Report');

  SpreadsheetApp.getUi().showModalDialog(html, 'Bug Report Generator');
}

/**
 * Gets recent processing log entries for the UI.
 * @returns {Array} Recent log entries
 */
function getRecentProcessingLogForBugReport() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Processing_Log');

  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Get last 50 entries
  const recent = data.slice(-50).reverse();

  return recent.map(row => {
    const entry = {};
    headers.forEach((header, i) => {
      entry[header] = row[i];
    });
    return entry;
  });
}

/**
 * Gets all clients for the dropdown.
 * @returns {Array} Client names
 */
function getClientsForBugReport() {
  const clients = getClientRegistry();
  return clients.map(c => c.client_name).sort();
}

/**
 * Generates a comprehensive bug report based on search criteria.
 * @param {Object} criteria - Search criteria
 * @returns {string} Formatted bug report
 */
function generateBugReport(criteria) {
  Logger.log('Generating bug report with criteria: ' + JSON.stringify(criteria));

  const report = [];
  const startTime = new Date(criteria.startTime);
  const endTime = new Date(criteria.endTime);

  report.push('# BUG REPORT');
  report.push('');
  report.push(`**Generated:** ${new Date().toISOString()}`);
  report.push(`**Time Range:** ${startTime.toISOString()} to ${endTime.toISOString()}`);
  if (criteria.clientName) {
    report.push(`**Client:** ${criteria.clientName}`);
  }
  report.push('');
  report.push('---');
  report.push('');

  // Section 0: Issues Summary (scan all logs first)
  const issuesSummary = [];

  // Check for truncated API responses
  if (isDiagnosticModeEnabled()) {
    const apiResponses = getDiagnosticLogEntries('API_Response_Log', startTime, endTime, criteria.clientName);
    apiResponses.forEach(log => {
      try {
        const response = typeof log.response_body === 'string' ? JSON.parse(log.response_body) : log.response_body;
        if (response && response.stop_reason === 'max_tokens') {
          issuesSummary.push(`⚠️ Truncated API response: ${log.api_name} at ${log.timestamp} (max_tokens limit reached)`);
        }
      } catch (e) {
        // Ignore parse errors
      }

      if (log.status_code && log.status_code >= 400) {
        issuesSummary.push(`❌ API Error: ${log.api_name} returned ${log.status_code} at ${log.timestamp}`);
      }
    });
  }

  // Check processing log for errors
  const processingLogs = getProcessingLogEntries(startTime, endTime, criteria.clientName);
  const errorLogs = processingLogs.filter(log => log.status === 'error');
  if (errorLogs.length > 0) {
    issuesSummary.push(`❌ ${errorLogs.length} error(s) in Processing Log`);
  }

  // Check for missing filter IDs
  if (criteria.clientName) {
    const client = getClientByName(criteria.clientName);
    if (client && client.setup_complete) {
      const missingFilters = [];
      if (!client.from_filter_id) missingFilters.push('from_filter_id');
      if (!client.to_filter_id) missingFilters.push('to_filter_id');
      if (!client.summary_filter_id) missingFilters.push('summary_filter_id');
      if (!client.agenda_filter_id) missingFilters.push('agenda_filter_id');

      if (missingFilters.length > 0) {
        issuesSummary.push(`⚠️ Missing filter IDs: ${missingFilters.join(', ')}`);
      }
    }
  }

  if (issuesSummary.length > 0) {
    report.push('## ⚠️ ISSUES DETECTED');
    report.push('');
    issuesSummary.forEach(issue => {
      report.push(`- ${issue}`);
    });
    report.push('');
    report.push('---');
    report.push('');
  }

  // Section 1: Client Details
  if (criteria.clientName) {
    report.push('## 1. CLIENT DETAILS');
    report.push('');
    const client = getClientByName(criteria.clientName);
    if (client) {
      report.push('```');
      report.push(`Client Name: ${client.client_name}`);
      report.push(`Contact Emails: ${client.contact_emails || '(none)'}`);
      report.push(`Setup Complete: ${client.setup_complete}`);
      report.push(`Doc URL: ${client.meeting_notes_doc_url || '(none)'}`);
      report.push(`Todoist Project: ${client.todoist_project_id || '(none)'}`);
      report.push(`Gmail Label: ${client.gmail_label || '(default)'}`);
      report.push(`Meeting Summaries Label: ${client.meeting_summaries_label || '(default)'}`);
      report.push(`Meeting Agendas Label: ${client.meeting_agendas_label || '(default)'}`);
      report.push(`From Filter ID: ${client.from_filter_id || '(none)'}`);
      report.push(`To Filter ID: ${client.to_filter_id || '(none)'}`);
      report.push(`Summary Filter ID: ${client.summary_filter_id || '(none)'}`);
      report.push(`Agenda Filter ID: ${client.agenda_filter_id || '(none)'}`);
      report.push('```');
      report.push('');
    } else {
      report.push('**ERROR:** Client not found in registry');
      report.push('');
    }
  }

  // Section 2: Processing Log Entries
  report.push('## 2. PROCESSING LOG');
  report.push('');
  const processingLogs = getProcessingLogEntries(startTime, endTime, criteria.clientName);
  if (processingLogs.length > 0) {
    processingLogs.forEach(log => {
      report.push(`### ${log.timestamp} - ${log.action_type} [${log.status}]`);
      report.push('```');
      report.push(`Client ID: ${log.client_id || 'N/A'}`);
      report.push(`Details: ${log.details || 'N/A'}`);
      report.push('```');
      report.push('');
    });
  } else {
    report.push('No processing log entries found in time range.');
    report.push('');
  }

  // Section 3: Diagnostic Logs (if available)
  if (isDiagnosticModeEnabled()) {
    report.push('## 3. DIAGNOSTIC LOGS');
    report.push('');

    // API Request Log
    const apiRequests = getDiagnosticLogEntries('API_Request_Log', startTime, endTime, criteria.clientName);
    if (apiRequests.length > 0) {
      report.push('### API Requests');
      report.push('');
      apiRequests.forEach(log => {
        report.push(`**${log.timestamp}** - ${log.api_name} (${log.method} ${log.endpoint})`);
        report.push('```json');
        report.push(`Request ID: ${log.request_id}`);
        report.push(`Client ID: ${log.client_id || 'N/A'}`);
        report.push(`Payload: ${log.payload || 'N/A'}`);
        report.push('```');
        report.push('');
      });
    }

    // API Response Log
    const apiResponses = getDiagnosticLogEntries('API_Response_Log', startTime, endTime, criteria.clientName);
    if (apiResponses.length > 0) {
      report.push('### API Responses');
      report.push('');
      apiResponses.forEach(log => {
        // Parse response body if it's JSON
        let parsedResponse = null;
        let isTruncated = false;
        let stopReason = null;

        try {
          if (log.response_body && typeof log.response_body === 'string') {
            parsedResponse = JSON.parse(log.response_body);
            stopReason = parsedResponse.stop_reason;
            isTruncated = stopReason === 'max_tokens';
          }
        } catch (e) {
          // Not JSON or parse failed
        }

        // Build status line with warning if truncated
        let statusLine = `**${log.timestamp}** - ${log.api_name} [${log.status_code}]`;
        if (isTruncated) {
          statusLine += ' ⚠️ **TRUNCATED**';
        }
        report.push(statusLine);

        report.push('```');
        report.push(`Request ID: ${log.request_id}`);
        report.push(`Duration: ${log.duration_ms}ms`);
        report.push(`Parse Success: ${log.parse_success || 'N/A'}`);

        if (log.error_message) {
          report.push(`❌ Error: ${log.error_message}`);
        }

        // Show parsed response details if available
        if (parsedResponse) {
          report.push('');
          report.push('Response Details:');
          if (parsedResponse.model) report.push(`  Model: ${parsedResponse.model}`);
          if (parsedResponse.stop_reason) {
            report.push(`  Stop Reason: ${parsedResponse.stop_reason}${parsedResponse.stop_reason === 'max_tokens' ? ' ⚠️ INCOMPLETE RESPONSE' : ''}`);
          }
          if (parsedResponse.usage) {
            report.push(`  Tokens: ${parsedResponse.usage.input_tokens} in, ${parsedResponse.usage.output_tokens} out`);
          }
          if (parsedResponse.content && parsedResponse.content[0]) {
            const contentLength = parsedResponse.content[0].text ? parsedResponse.content[0].text.length : 0;
            report.push(`  Content Length: ${contentLength} chars`);

            // Show content preview (truncated)
            if (contentLength > 0) {
              report.push('');
              report.push('Content Preview:');
              report.push(truncateForReport(parsedResponse.content[0].text, 300));
            }
          }

          // Show error if present in response
          if (parsedResponse.error) {
            report.push('');
            report.push(`❌ API Error: ${JSON.stringify(parsedResponse.error, null, 2)}`);
          }
        } else {
          // Show raw response if not JSON
          report.push('');
          report.push('Response Body:');
          report.push(truncateForReport(log.response_body, 500));
        }

        // Show extracted data if available
        if (log.extracted_data) {
          try {
            const extracted = typeof log.extracted_data === 'string' ? JSON.parse(log.extracted_data) : log.extracted_data;
            report.push('');
            report.push('Extracted Data:');
            report.push(JSON.stringify(extracted, null, 2));
          } catch (e) {
            // Ignore parse errors
          }
        }

        report.push('```');
        report.push('');
      });
    }

    // Agenda Generation Trace
    const agendaTraces = getDiagnosticLogEntries('Agenda_Generation_Trace', startTime, endTime, criteria.clientName);
    if (agendaTraces.length > 0) {
      report.push('### Agenda Generation Trace');
      report.push('');
      agendaTraces.forEach(log => {
        report.push(`**${log.timestamp}** - Step ${log.step_number}: ${log.step_name} [${log.step_status}]`);
        report.push('```');
        report.push(`Event ID: ${log.event_id || 'N/A'}`);
        report.push(`Event Title: ${log.event_title || 'N/A'}`);
        report.push(`Client ID: ${log.client_id || 'N/A'}`);
        report.push(`Details: ${truncateForReport(log.step_details, 300)}`);
        report.push(`Data: ${truncateForReport(log.data_summary, 300)}`);
        report.push(`Duration: ${log.duration_ms || 0}ms`);
        report.push('```');
        report.push('');
      });
    }
  } else {
    report.push('## 3. DIAGNOSTIC LOGS');
    report.push('');
    report.push('*Diagnostic mode is not enabled. Enable it to capture detailed traces.*');
    report.push('');
  }

  // Section 4: Recent Emails (if Gmail API available)
  if (typeof Gmail !== 'undefined' && criteria.clientName) {
    report.push('## 4. RECENT EMAILS');
    report.push('');
    const client = getClientByName(criteria.clientName);
    if (client && client.meeting_agendas_label) {
      try {
        const emails = getRecentEmailsFromLabel(client.meeting_agendas_label, startTime, endTime);
        if (emails.length > 0) {
          emails.forEach(email => {
            report.push(`### ${email.subject || '(No Subject)'}`);
            report.push('```');
            report.push(`Date: ${email.date}`);
            report.push(`From: ${email.from}`);
            report.push(`To: ${email.to}`);
            report.push(`Snippet: ${email.snippet || '(empty)'}`);
            report.push(`Body Length: ${email.bodyLength} characters`);
            report.push('```');
            report.push('');
          });
        } else {
          report.push('No emails found in Meeting Agendas label for this time range.');
          report.push('');
        }
      } catch (e) {
        report.push(`Error fetching emails: ${e.message}`);
        report.push('');
      }
    }
  }

  // Section 5: System Configuration
  report.push('## 5. SYSTEM CONFIGURATION');
  report.push('');
  report.push('```');
  const props = PropertiesService.getScriptProperties().getProperties();
  report.push(`Diagnostic Mode: ${props.DIAGNOSTIC_MODE || 'false'}`);
  report.push(`Fathom API Key: ${props.FATHOM_API_KEY ? 'SET' : 'NOT SET'}`);
  report.push(`Claude API Key: ${props.CLAUDE_API_KEY ? 'SET' : 'NOT SET'}`);
  report.push(`Todoist API Token: ${props.TODOIST_API_TOKEN ? 'SET' : 'NOT SET'}`);
  report.push(`Agenda Subject Template: ${props.AGENDA_SUBJECT_TEMPLATE || '(default)'}`);
  report.push(`Meeting Summary Template: ${props.MEETING_SUMMARY_SUBJECT_TEMPLATE || '(default)'}`);
  report.push('```');
  report.push('');

  report.push('---');
  report.push('');
  report.push('**END OF BUG REPORT**');

  return report.join('\n');
}

/**
 * Gets client by name from registry.
 * @param {string} clientName - Client name
 * @returns {Object|null} Client object
 */
function getClientByName(clientName) {
  const clients = getClientRegistry();
  return clients.find(c => c.client_name === clientName) || null;
}

/**
 * Gets processing log entries in time range.
 * @param {Date} startTime - Start time
 * @param {Date} endTime - End time
 * @param {string} clientName - Optional client filter
 * @returns {Array} Log entries
 */
function getProcessingLogEntries(startTime, endTime, clientName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Processing_Log');

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const timestampIdx = headers.indexOf('timestamp');
  const clientIdIdx = headers.indexOf('client_id');

  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const timestamp = new Date(data[i][timestampIdx]);
    if (timestamp >= startTime && timestamp <= endTime) {
      if (clientName && data[i][clientIdIdx] !== clientName) continue;

      const entry = {};
      headers.forEach((header, idx) => {
        entry[header] = data[i][idx];
      });
      entries.push(entry);
    }
  }

  return entries;
}

/**
 * Gets diagnostic log entries from a specific sheet.
 * @param {string} sheetName - Diagnostic sheet name
 * @param {Date} startTime - Start time
 * @param {Date} endTime - End time
 * @param {string} clientName - Optional client filter
 * @returns {Array} Log entries
 */
function getDiagnosticLogEntries(sheetName, startTime, endTime, clientName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const timestampIdx = headers.indexOf('Timestamp') >= 0 ? headers.indexOf('Timestamp') : headers.indexOf('timestamp');
  const clientIdIdx = headers.indexOf('Client_ID') >= 0 ? headers.indexOf('Client_ID') : headers.indexOf('client_id');

  if (timestampIdx === -1) return [];

  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const timestamp = new Date(data[i][timestampIdx]);
    if (timestamp >= startTime && timestamp <= endTime) {
      if (clientName && clientIdIdx >= 0 && data[i][clientIdIdx] !== clientName) continue;

      const entry = {};
      headers.forEach((header, idx) => {
        entry[header.toLowerCase()] = data[i][idx];
      });
      entries.push(entry);
    }
  }

  return entries;
}

/**
 * Gets recent emails from a label in time range.
 * @param {string} labelName - Gmail label name
 * @param {Date} startTime - Start time
 * @param {Date} endTime - End time
 * @returns {Array} Email objects
 */
function getRecentEmailsFromLabel(labelName, startTime, endTime) {
  try {
    const startStr = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    const endStr = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    const query = `label:"${labelName}" after:${startStr} before:${endStr}`;
    const threads = GmailApp.search(query, 0, 10);

    const emails = [];
    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(msg => {
        const msgDate = msg.getDate();
        if (msgDate >= startTime && msgDate <= endTime) {
          emails.push({
            subject: msg.getSubject(),
            date: msgDate.toISOString(),
            from: msg.getFrom(),
            to: msg.getTo(),
            snippet: msg.getPlainBody().substring(0, 150),
            bodyLength: msg.getPlainBody().length
          });
        }
      });
    });

    return emails;
  } catch (e) {
    Logger.log(`Error fetching emails: ${e.message}`);
    return [];
  }
}

/**
 * Truncates text for report display.
 * @param {string} text - Text to truncate
 * @param {number} maxLength - Max length
 * @returns {string} Truncated text
 */
function truncateForReport(text, maxLength) {
  if (!text) return '(empty)';
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength) + '... [truncated]';
}
