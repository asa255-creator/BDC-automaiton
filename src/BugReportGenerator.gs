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
 * @returns {Object} Object with 'report' and 'issuesSummary' strings
 */
function generateBugReport(criteria) {
  Logger.log('Generating bug report with criteria: ' + JSON.stringify(criteria));

  const report = [];
  const startTime = new Date(criteria.startTime);
  const endTime = new Date(criteria.endTime);
  const reportType = criteria.reportType || 'all';

  const reportTypeNames = {
    all: 'All Automations',
    daily_briefing: 'Daily Briefing',
    weekly_briefing: 'Weekly Briefing',
    pre_meeting_agenda: 'Pre-Meeting Agenda',
    fathom_drafts: 'Fathom Meeting Drafts',
    meeting_notes: 'Meeting Notes Append & Todoist'
  };

  report.push('# BUG REPORT');
  report.push('');
  report.push(`**Generated:** ${new Date().toISOString()}`);
  report.push(`**Report Type:** ${reportTypeNames[reportType] || 'All Automations'}`);
  report.push(`**Time Range:** ${startTime.toISOString()} to ${endTime.toISOString()}`);
  if (criteria.clientName) {
    report.push(`**Client:** ${criteria.clientName}`);
  }
  report.push('');
  report.push('---');
  report.push('');

  // Section 1: Client Details (if specified)
  if (criteria.clientName && (reportType === 'all' || reportType === 'pre_meeting_agenda' || reportType === 'meeting_notes')) {
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
  let processingLogs = getProcessingLogEntries(startTime, endTime, criteria.clientName);

  // Filter by action_type based on reportType
  const actionTypeFilters = {
    daily_briefing: ['DAILY_OUTLOOK', 'OUTLOOK_SEND', 'DAILY_BRIEF_GEN', 'DAILY_BRIEF_EMAIL'],
    weekly_briefing: ['WEEKLY_OUTLOOK', 'OUTLOOK_SEND', 'WEEKLY_BRIEF_GEN', 'WEEKLY_BRIEF_EMAIL'],
    pre_meeting_agenda: ['AGENDA_GEN', 'AGENDA_EMAIL', 'AGENDA_SEND'],
    fathom_drafts: ['WEBHOOK_PROCESS', 'WEBHOOK_PAYLOAD', 'WEBHOOK_SUCCESS', 'WEBHOOK_ERROR', 'FATHOM_POLL', 'FATHOM_PROCESS', 'FATHOM_DRAFT'],
    meeting_notes: ['NOTES_APPEND', 'TODOIST_CREATE', 'TODOIST_UPDATE', 'TODOIST_TASK', 'MEETING_SUMMARY']
  };

  if (reportType !== 'all' && actionTypeFilters[reportType]) {
    processingLogs = processingLogs.filter(log =>
      actionTypeFilters[reportType].includes(log.action_type)
    );
  }

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

  // Section 3: Fathom Drafts Status (only for fathom_drafts or all)
  if (reportType === 'all' || reportType === 'fathom_drafts') {
    report.push('## 3. FATHOM DRAFTS STATUS');
    report.push('');
    const fathomDiagnostics = diagnoseFathomDrafts(startTime, endTime);
    report.push(fathomDiagnostics);
    report.push('');
  }

  // Section 4: Diagnostic Logs (if available)
  if (isDiagnosticModeEnabled()) {
    report.push('## 4. DIAGNOSTIC LOGS');
    report.push('');

    // Define API name filters based on report type
    const apiNameFilters = {
      daily_briefing: ['Claude API', 'claude-daily-briefing'],
      weekly_briefing: ['Claude API', 'claude-weekly-briefing'],
      pre_meeting_agenda: ['Claude API', 'claude-agenda', 'Google Calendar API'],
      fathom_drafts: ['Fathom API', 'Claude API', 'claude-summary'],
      meeting_notes: ['Claude API', 'Todoist API', 'Google Docs API']
    };

    // API Request Log
    let apiRequests = getDiagnosticLogEntries('API_Request_Log', startTime, endTime, criteria.clientName);

    // Filter by API name if specific report type
    if (reportType !== 'all' && apiNameFilters[reportType]) {
      apiRequests = apiRequests.filter(log =>
        apiNameFilters[reportType].some(filter =>
          log.api_name && log.api_name.includes(filter)
        )
      );
    }

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
    let apiResponses = getDiagnosticLogEntries('API_Response_Log', startTime, endTime, criteria.clientName);

    // Filter by API name if specific report type
    if (reportType !== 'all' && apiNameFilters[reportType]) {
      apiResponses = apiResponses.filter(log =>
        apiNameFilters[reportType].some(filter =>
          log.api_name && log.api_name.includes(filter)
        )
      );
    }
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

    // Agenda Generation Trace (only for pre_meeting_agenda or all)
    if (reportType === 'all' || reportType === 'pre_meeting_agenda') {
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
    }
  } else {
    report.push('## 4. DIAGNOSTIC LOGS');
    report.push('');
    report.push('*Diagnostic mode is not enabled. Enable it to capture detailed traces.*');
    report.push('');
  }

  // Section 5: Recent Emails (if Gmail API available)
  if (typeof Gmail !== 'undefined') {
    report.push('## 5. RECENT EMAILS');
    report.push('');

    // Determine which label to check based on report type
    let labelToCheck = null;
    let labelDescription = '';

    if (reportType === 'pre_meeting_agenda' && criteria.clientName) {
      const client = getClientByName(criteria.clientName);
      if (client && client.meeting_agendas_label) {
        labelToCheck = client.meeting_agendas_label;
        labelDescription = 'Meeting Agendas';
      }
    } else if (reportType === 'fathom_drafts' && criteria.clientName) {
      const client = getClientByName(criteria.clientName);
      if (client && client.meeting_summaries_label) {
        labelToCheck = client.meeting_summaries_label;
        labelDescription = 'Meeting Summaries';
      }
    } else if (reportType === 'meeting_notes' && criteria.clientName) {
      const client = getClientByName(criteria.clientName);
      if (client && client.meeting_summaries_label) {
        labelToCheck = client.meeting_summaries_label;
        labelDescription = 'Meeting Summaries';
      }
    } else if (reportType === 'daily_briefing') {
      const props = PropertiesService.getScriptProperties();
      labelToCheck = props.getProperty('DAILY_BRIEFING_LABEL') || 'Brief: Daily';
      labelDescription = 'Daily Briefing';
    } else if (reportType === 'weekly_briefing') {
      const props = PropertiesService.getScriptProperties();
      labelToCheck = props.getProperty('WEEKLY_BRIEFING_LABEL') || 'Brief: Weekly';
      labelDescription = 'Weekly Briefing';
    } else if (reportType === 'all' && criteria.clientName) {
      const client = getClientByName(criteria.clientName);
      if (client && client.gmail_label) {
        labelToCheck = client.gmail_label;
        labelDescription = 'Client Emails';
      }
    }

    if (labelToCheck) {
      try {
        const emails = getRecentEmailsFromLabel(labelToCheck, startTime, endTime);
        if (emails.length > 0) {
          report.push(`Found ${emails.length} email(s) in "${labelDescription}" label:`);
          report.push('');
          emails.forEach(email => {
            report.push(`### ${email.subject || '(No Subject)'}`);
            report.push('```');
            report.push(`Date: ${email.date}`);
            report.push(`From: ${email.from}`);
            report.push(`To: ${email.to}`);
            report.push(`Snippet: ${email.snippet || '(empty)'}`);
            report.push(`Body Length: ${email.bodyLength} characters`);
            if ((reportType === 'daily_briefing' || reportType === 'weekly_briefing') && email.htmlSnippet) {
              report.push('HTML Preview:');
              report.push(truncateForReport(email.htmlSnippet, 1000));
            }
            report.push('```');
            report.push('');
          });
        } else {
          report.push(`No emails found in "${labelDescription}" label for this time range.`);
          report.push('');
        }
      } catch (e) {
        report.push(`Error fetching emails: ${e.message}`);
        report.push('');
      }
    } else {
      // No label to check - provide helpful message based on report type
      if (reportType === 'fathom_drafts' || reportType === 'meeting_notes') {
        report.push('**Note:** No client specified. To check sent meeting summaries, select a specific client.');
        report.push('');
        report.push('Flow: Fathom meeting → draft created → you send → Gmail filters auto-label → monitoring processes it');
        report.push('This section checks for SENT emails that filters have labeled, not drafts.');
        report.push('');
        report.push('Check Section 3 (Fathom Drafts Status) above for draft creation diagnostics.');
      } else if (reportType === 'pre_meeting_agenda') {
        report.push('**Note:** No client specified. To check sent agendas, select a specific client.');
        report.push('The automation sends agendas with the client\'s Meeting Agendas label.');
      } else {
        report.push('No relevant email label to check for this report type.');
      }
      report.push('');
    }
  }

  // Section 6: System Configuration
  report.push('## 6. SYSTEM CONFIGURATION');
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

  // Generate issues summary separately
  const issuesSummary = generateIssuesSummary(startTime, endTime, criteria.clientName, reportType);

  return {
    report: report.join('\n'),
    issuesSummary: issuesSummary
  };
}

/**
 * Generates an issues summary by scanning logs for common problems.
 * @param {Date} startTime - Start time
 * @param {Date} endTime - End time
 * @param {string} clientName - Optional client filter
 * @param {string} reportType - Report type filter
 * @returns {string} Formatted issues summary
 */
function generateIssuesSummary(startTime, endTime, clientName, reportType) {
  const summary = [];
  const issues = [];
  reportType = reportType || 'all';

  const reportTypeNames = {
    all: 'All Automations',
    daily_briefing: 'Daily Briefing',
    weekly_briefing: 'Weekly Briefing',
    pre_meeting_agenda: 'Pre-Meeting Agenda',
    fathom_drafts: 'Fathom Meeting Drafts',
    meeting_notes: 'Meeting Notes Append & Todoist'
  };

  summary.push('# ISSUES SUMMARY');
  summary.push('');
  summary.push(`**Scanned:** ${new Date().toISOString()}`);
  summary.push(`**Report Type:** ${reportTypeNames[reportType] || 'All Automations'}`);
  summary.push(`**Time Range:** ${startTime.toISOString()} to ${endTime.toISOString()}`);
  if (clientName) {
    summary.push(`**Client:** ${clientName}`);
  }
  summary.push('');
  summary.push('---');
  summary.push('');

  // Check for truncated API responses
  if (isDiagnosticModeEnabled()) {
    let apiResponses = getDiagnosticLogEntries('API_Response_Log', startTime, endTime, clientName);

    // Filter by API name based on report type
    const apiNameFilters = {
      daily_briefing: ['Claude API', 'claude-daily-briefing'],
      weekly_briefing: ['Claude API', 'claude-weekly-briefing'],
      pre_meeting_agenda: ['Claude API', 'claude-agenda', 'Google Calendar API'],
      fathom_drafts: ['Fathom API', 'Claude API', 'claude-summary'],
      meeting_notes: ['Claude API', 'Todoist API', 'Google Docs API']
    };

    if (reportType !== 'all' && apiNameFilters[reportType]) {
      apiResponses = apiResponses.filter(log =>
        apiNameFilters[reportType].some(filter =>
          log.api_name && log.api_name.includes(filter)
        )
      );
    }

    const truncatedResponses = [];
    const errorResponses = [];

    apiResponses.forEach(log => {
      // Check for truncation
      try {
        const response = typeof log.response_body === 'string' ? JSON.parse(log.response_body) : log.response_body;
        if (response && response.stop_reason === 'max_tokens') {
          truncatedResponses.push({
            api: log.api_name,
            time: log.timestamp,
            tokens: response.usage ? response.usage.output_tokens : 'unknown'
          });
        }
      } catch (e) {
        // Ignore parse errors
      }

      // Check for API errors
      if (log.status_code && log.status_code >= 400) {
        errorResponses.push({
          api: log.api_name,
          status: log.status_code,
          time: log.timestamp,
          error: log.error_message || 'No error message'
        });
      }
    });

    if (truncatedResponses.length > 0) {
      issues.push({
        type: '⚠️ TRUNCATED API RESPONSES',
        count: truncatedResponses.length,
        details: truncatedResponses
      });
    }

    if (errorResponses.length > 0) {
      issues.push({
        type: '❌ API ERRORS',
        count: errorResponses.length,
        details: errorResponses
      });
    }
  }

  // Check processing log for errors
  let processingLogs = getProcessingLogEntries(startTime, endTime, clientName);

  // Filter by action_type based on reportType
  const actionTypeFilters = {
    daily_briefing: ['DAILY_OUTLOOK', 'OUTLOOK_SEND', 'DAILY_BRIEF_GEN', 'DAILY_BRIEF_EMAIL'],
    weekly_briefing: ['WEEKLY_OUTLOOK', 'OUTLOOK_SEND', 'WEEKLY_BRIEF_GEN', 'WEEKLY_BRIEF_EMAIL'],
    pre_meeting_agenda: ['AGENDA_GEN', 'AGENDA_EMAIL', 'AGENDA_SEND'],
    fathom_drafts: ['WEBHOOK_PROCESS', 'WEBHOOK_PAYLOAD', 'WEBHOOK_SUCCESS', 'WEBHOOK_ERROR', 'FATHOM_POLL', 'FATHOM_PROCESS', 'FATHOM_DRAFT'],
    meeting_notes: ['NOTES_APPEND', 'TODOIST_CREATE', 'TODOIST_UPDATE', 'TODOIST_TASK', 'MEETING_SUMMARY']
  };

  if (reportType !== 'all' && actionTypeFilters[reportType]) {
    processingLogs = processingLogs.filter(log =>
      actionTypeFilters[reportType].includes(log.action_type)
    );
  }

  const errorLogs = processingLogs.filter(log => log.status === 'error');
  if (errorLogs.length > 0) {
    issues.push({
      type: '❌ PROCESSING ERRORS',
      count: errorLogs.length,
      details: errorLogs.map(log => ({
        action: log.action_type,
        time: log.timestamp,
        client: log.client_id,
        details: log.details
      }))
    });
  }

  // Check for missing filter IDs (only for client-specific report types)
  if (clientName && (reportType === 'all' || reportType === 'pre_meeting_agenda' || reportType === 'meeting_notes')) {
    const client = getClientByName(clientName);
    if (client && client.setup_complete) {
      const missingFilters = [];
      if (!client.from_filter_id) missingFilters.push('from_filter_id');
      if (!client.to_filter_id) missingFilters.push('to_filter_id');
      if (!client.summary_filter_id) missingFilters.push('summary_filter_id');
      if (!client.agenda_filter_id) missingFilters.push('agenda_filter_id');

      if (missingFilters.length > 0) {
        issues.push({
          type: '⚠️ MISSING FILTER IDs',
          count: missingFilters.length,
          details: missingFilters.map(f => ({ filter: f }))
        });
      }
    }
  }

  // Format issues
  if (issues.length === 0) {
    summary.push('✅ **NO ISSUES DETECTED**');
    summary.push('');
    summary.push('All checks passed. No common problems found in logs.');
  } else {
    summary.push(`**${issues.length} ISSUE TYPE(S) DETECTED**`);
    summary.push('');

    issues.forEach(issue => {
      summary.push(`## ${issue.type}`);
      summary.push('');
      summary.push(`**Count:** ${issue.count}`);
      summary.push('');

      // Show details
      if (issue.type === '⚠️ TRUNCATED API RESPONSES') {
        issue.details.forEach(detail => {
          summary.push(`- **${detail.api}** at ${detail.time}`);
          summary.push(`  - Output tokens: ${detail.tokens}`);
          summary.push(`  - Issue: Response was cut off due to max_tokens limit`);
        });
      } else if (issue.type === '❌ API ERRORS') {
        issue.details.forEach(detail => {
          summary.push(`- **${detail.api}** at ${detail.time}`);
          summary.push(`  - Status: ${detail.status}`);
          summary.push(`  - Error: ${detail.error}`);
        });
      } else if (issue.type === '❌ PROCESSING ERRORS') {
        issue.details.forEach(detail => {
          summary.push(`- **${detail.action}** at ${detail.time}`);
          summary.push(`  - Client: ${detail.client || 'N/A'}`);
          summary.push(`  - Details: ${detail.details || 'No details'}`);
        });
      } else if (issue.type === '⚠️ MISSING FILTER IDs') {
        issue.details.forEach(detail => {
          summary.push(`- ${detail.filter}`);
        });
        summary.push('');
        summary.push('**Action Required:** Run "Sync Gmail Labels & Filters" from menu');
      }

      summary.push('');
    });
  }

  summary.push('---');
  summary.push('');
  summary.push('**Note:** This summary scans for common issues. Check the full bug report for complete details.');

  return summary.join('\n');
}

/**
 * Diagnoses Fathom draft creation issues.
 * @param {Date} startTime - Start time
 * @param {Date} endTime - End time
 * @returns {string} Diagnostic report
 */
function diagnoseFathomDrafts(startTime, endTime) {
  const diagnostics = [];

  try {
    // 1. Check Processing_Log for Fathom webhook/polling activity
    const processingLogs = getProcessingLogEntries(startTime, endTime, null);
    const webhookLogs = processingLogs.filter(log =>
      log.action_type === 'WEBHOOK_PROCESS' ||
      log.action_type === 'WEBHOOK_PAYLOAD' ||
      log.action_type === 'WEBHOOK_SUCCESS' ||
      log.action_type === 'WEBHOOK_ERROR' ||
      log.action_type === 'FATHOM_POLL'
    );

    diagnostics.push(`### Fathom Activity in Processing_Log`);
    diagnostics.push('');
    if (webhookLogs.length > 0) {
      diagnostics.push(`Found ${webhookLogs.length} Fathom-related log entries:`);
      diagnostics.push('');
      webhookLogs.forEach(log => {
        diagnostics.push(`- **${log.timestamp}** [${log.action_type}] ${log.status}`);
        diagnostics.push(`  Details: ${log.details || 'N/A'}`);
      });
    } else {
      diagnostics.push('❌ **NO Fathom activity found** - Webhooks/polling not triggering');
    }
    diagnostics.push('');

    // 2. Check Processed_Fathom sheet
    diagnostics.push(`### Processed Fathom Meetings`);
    diagnostics.push('');
    const processedMeetings = getProcessedFathomMeetings(startTime, endTime);
    if (processedMeetings.length > 0) {
      diagnostics.push(`Found ${processedMeetings.length} processed meetings:`);
      diagnostics.push('');
      processedMeetings.forEach(meeting => {
        diagnostics.push(`- **${meeting.meeting_title}**`);
        diagnostics.push(`  Meeting Date: ${meeting.meeting_date}`);
        diagnostics.push(`  Processed At: ${meeting.processed_at}`);
        diagnostics.push(`  Client: ${meeting.client_name || '(no match)'}`);
        diagnostics.push(`  Draft ID: ${meeting.draft_id || 'NONE'}`);
      });
    } else {
      diagnostics.push('No meetings recorded in Processed_Fathom sheet');
    }
    diagnostics.push('');

    // 3. Check actual Gmail drafts
    diagnostics.push(`### Current Gmail Drafts`);
    diagnostics.push('');
    try {
      const drafts = GmailApp.getDrafts();
      diagnostics.push(`Total drafts in Gmail: ${drafts.length}`);
      diagnostics.push('');

      if (drafts.length > 0) {
        diagnostics.push('Recent drafts (last 10):');
        const recentDrafts = drafts.slice(0, 10);
        recentDrafts.forEach(draft => {
          const message = draft.getMessage();
          const draftDate = message.getDate();
          const subject = message.getSubject();
          diagnostics.push(`- **${subject}**`);
          diagnostics.push(`  Created: ${draftDate.toISOString()}`);
          diagnostics.push(`  To: ${message.getTo() || '(no recipient)'}`);
        });
      } else {
        diagnostics.push('⚠️ **NO drafts in Gmail** - Drafts may have been sent or deleted');
      }
    } catch (e) {
      diagnostics.push(`Error checking Gmail drafts: ${e.message}`);
    }
    diagnostics.push('');

    // 4. Check actual Fathom API meetings
    diagnostics.push(`### Fathom API Recent Meetings`);
    diagnostics.push('');
    try {
      const apiKey = PropertiesService.getScriptProperties().getProperty('FATHOM_API_KEY');
      if (!apiKey) {
        diagnostics.push('⚠️ Fathom API key not configured - cannot fetch meeting data');
      } else {
        const url = 'https://api.fathom.ai/external/v1/meetings?include_transcript=false&include_summary=false&include_action_items=false&limit=5';
        const options = {
          method: 'GET',
          headers: {
            'X-Api-Key': apiKey,
            'Content-Type': 'application/json'
          },
          muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();

        if (responseCode !== 200) {
          diagnostics.push(`❌ Fathom API error (${responseCode}): ${response.getContentText().substring(0, 200)}`);
        } else {
          const data = JSON.parse(response.getContentText());
          const meetings = data.items || (Array.isArray(data) ? data : []);

          if (meetings.length === 0) {
            diagnostics.push('No meetings found in Fathom API');
          } else {
            diagnostics.push(`Found ${meetings.length} recent meetings in Fathom API:`);
            diagnostics.push('');
            meetings.forEach((meeting, index) => {
              diagnostics.push(`**Meeting ${index + 1}:**`);
              diagnostics.push('```json');
              // Show key fields to diagnose the ID issue
              const meetingInfo = {
                id: meeting.id || 'MISSING',
                meeting_id: meeting.meeting_id || 'MISSING',
                title: meeting.title || meeting.meeting_title || 'MISSING',
                created_at: meeting.created_at || 'MISSING',
                url: meeting.url || meeting.share_url || 'MISSING',
                available_fields: Object.keys(meeting)
              };
              diagnostics.push(JSON.stringify(meetingInfo, null, 2));
              diagnostics.push('```');
              diagnostics.push('');
            });
          }
        }
      }
    } catch (e) {
      diagnostics.push(`Error fetching Fathom API data: ${e.message}`);
    }
    diagnostics.push('');

    // 5. Analysis
    diagnostics.push(`### Analysis`);
    diagnostics.push('');

    if (webhookLogs.length === 0) {
      diagnostics.push('❌ **ROOT CAUSE: No Fathom webhook/polling activity detected**');
      diagnostics.push('   - Check if Fathom webhook is configured correctly');
      diagnostics.push('   - Check if polling trigger is running every 30 minutes');
      diagnostics.push('   - Verify FATHOM_API_KEY is set in Settings');
    } else if (processedMeetings.length === 0) {
      diagnostics.push('⚠️ **Webhooks received but meetings not being processed**');
      diagnostics.push('   - Check Processing_Log for WEBHOOK_ERROR entries');
      diagnostics.push('   - Verify meeting payload format from Fathom');
    } else if (processedMeetings.some(m => !m.draft_id)) {
      diagnostics.push('⚠️ **Meetings processed but draft creation failing**');
      diagnostics.push('   - Check for Gmail API errors');
      diagnostics.push('   - Verify GmailApp.createDraft() permissions');
    } else {
      diagnostics.push('✅ Meetings processed and drafts created successfully');
      if (drafts.length === 0) {
        diagnostics.push('   Note: Drafts may have been sent or manually deleted');
      }
    }

  } catch (e) {
    diagnostics.push(`Error diagnosing Fathom drafts: ${e.message}`);
  }

  return diagnostics.join('\n');
}

/**
 * Gets processed Fathom meetings from the sheet.
 * @param {Date} startTime - Start time
 * @param {Date} endTime - End time
 * @returns {Array} Array of meeting records
 */
function getProcessedFathomMeetings(startTime, endTime) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Processed_Fathom');

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const processedAtIdx = headers.indexOf('Processed At');

    if (processedAtIdx === -1) return [];

    const meetings = [];
    for (let i = 1; i < data.length; i++) {
      const processedAt = new Date(data[i][processedAtIdx]);
      if (processedAt >= startTime && processedAt <= endTime) {
        meetings.push({
          meeting_id: data[i][0],
          meeting_title: data[i][1],
          meeting_date: data[i][2],
          processed_at: data[i][3],
          client_name: data[i][4],
          draft_id: data[i][5]
        });
      }
    }

    return meetings;
  } catch (e) {
    Logger.log(`Error getting processed Fathom meetings: ${e.message}`);
    return [];
  }
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
    const timestamp = parseProcessingLogTimestamp(data[i][timestampIdx]);
    if (!timestamp) {
      continue;
    }
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
 * Parses Processing_Log timestamps (Date objects or formatted strings).
 *
 * @param {Date|string} value - Timestamp value
 * @returns {Date|null} Parsed Date or null if invalid
 */
function parseProcessingLogTimestamp(value) {
  if (value instanceof Date) {
    return value;
  }

  if (typeof value !== 'string') {
    return null;
  }

  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }

  return null;
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
    const timezone = getUserTimezone();
    const startStr = Utilities.formatDate(startTime, timezone, 'yyyy/MM/dd');
    const endStr = Utilities.formatDate(endTime, timezone, 'yyyy/MM/dd');

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
            bodyLength: msg.getPlainBody().length,
            htmlSnippet: msg.getBody().substring(0, 2000)
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
