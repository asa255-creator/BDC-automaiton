/**
 * MeetingAutomation.gs - Fathom webhook handling and draft creation.
 *
 * This module handles:
 * 1. Receiving Fathom webhooks when meetings end
 * 2. Creating draft meeting summary emails
 * 3. Polling Fathom for new meetings as a backup
 */

// ============================================================================
// FATHOM WEBHOOK PROCESSING
// ============================================================================

/**
 * Processes an incoming Fathom webhook payload.
 * ALWAYS creates a draft email - client matching happens when email is sent.
 *
 * @param {Object} payload - The webhook payload from Fathom
 * @returns {Object} Result of the processing
 *
 * Payload structure:
 * - meeting_title: string
 * - meeting_date: string (ISO date)
 * - transcript: string
 * - summary: string
 * - action_items: Array<{description, assignee, due_date}>
 * - participants: Array<{name, email}>
 * - fathom_url: string (optional)
 */
function processFathomWebhook(payload) {
  Logger.log('Processing Fathom webhook...');

  // Validate payload
  if (!payload || !payload.meeting_title) {
    throw new Error('Invalid webhook payload: missing meeting_title');
  }

  // Extract meeting ID for tracking
  let meetingId = payload.meeting_id || payload.id;

  // Try extracting from fathom_url if no direct ID
  if (!meetingId && payload.fathom_url) {
    const urlMatch = payload.fathom_url.match(/\/calls\/(\d+)/);
    if (urlMatch && urlMatch[1]) {
      meetingId = urlMatch[1];
    }
  }

  // Fall back to generated ID from title+date
  if (!meetingId) {
    meetingId = `${payload.meeting_title}_${payload.meeting_date}`;
  }

  // Extract participant emails for logging
  const participantEmails = (payload.participants || [])
    .map(p => p.email)
    .filter(e => e);

  // Try to identify client from participants (optional - for pre-filling recipient)
  const client = identifyClientFromParticipants(payload.participants);

  if (client) {
    Logger.log(`Client identified: ${client.client_name}`);
  } else {
    Logger.log('No client matched - draft will be created without recipient');
  }

  // ALWAYS create draft email (user will add recipient if needed)
  const draftId = createMeetingSummaryDraft(payload, client);

  // Record meeting as processed in spreadsheet
  recordFathomMeetingProcessed(
    meetingId,
    payload.meeting_title,
    payload.meeting_date,
    client ? client.client_name : null,
    draftId
  );

  // Log processing result
  const clientName = client ? client.client_name : null;
  const status = client ? 'success' : 'draft_created';
  const message = client
    ? `Created draft for meeting: ${payload.meeting_title} (client: ${client.client_name})`
    : `Created draft for meeting: ${payload.meeting_title} (no client matched - add recipient manually)`;

  logProcessing('WEBHOOK_PROCESS', clientName, message, status);

  return {
    status: status,
    client_name: clientName,
    draft_id: draftId,
    participants: participantEmails.length
  };
}

/**
 * Strips calendar response prefixes from meeting titles.
 * Google Calendar adds prefixes like "Confirmed:", "Accepted:", "Tentative:", etc.
 *
 * @param {string} title - The meeting title
 * @returns {string} The cleaned title
 */
function cleanMeetingTitle(title) {
  if (!title) return title;

  // Strip common calendar response prefixes
  const prefixes = [
    'Confirmed:',
    'Accepted:',
    'Tentative:',
    'Declined:',
    'Not Responded:',
    'Maybe:',
    'Yes:',
    'No:'
  ];

  let cleanedTitle = title.trim();

  for (const prefix of prefixes) {
    if (cleanedTitle.startsWith(prefix)) {
      cleanedTitle = cleanedTitle.substring(prefix.length).trim();
      break;
    }
  }

  return cleanedTitle;
}

/**
 * Creates a Gmail draft with the meeting summary.
 * Works with or without a matched client.
 *
 * @param {Object} payload - The Fathom webhook payload
 * @param {Object|null} client - The matched client object (or null if no match)
 * @returns {string} The draft ID
 */
function createMeetingSummaryDraft(payload, client) {
  const props = PropertiesService.getScriptProperties();
  const meetingDate = formatDateShort(new Date(payload.meeting_date));

  // Clean meeting title to remove calendar prefixes like "Confirmed:", "Accepted:", etc.
  const cleanedTitle = cleanMeetingTitle(payload.meeting_title);

  // Build subject and greeting based on whether we have a client
  let subject, body;

  if (client) {
    // CLIENT MEETING: Use client-specific language
    const clientName = client.client_name;

    const subjectTemplate = props.getProperty('MEETING_SUBJECT_TEMPLATE')
      || 'Team {client_name} - Meeting notes from "{meeting_title}" {date}';
    subject = subjectTemplate
      .replace('{client_name}', clientName)
      .replace('{meeting_title}', cleanedTitle)
      .replace('{date}', meetingDate);

    body = `<p>Team ${clientName} -</p>`;
    body += `<p>Here are the notes from the meeting "${cleanedTitle}" ${meetingDate}.</p>`;
    body += `<hr/>`;

  } else {
    // NON-CLIENT MEETING: Use generic professional language
    subject = `Meeting notes: ${cleanedTitle} (${meetingDate})`;

    body = `<p>Hello -</p>`;
    body += `<p>Here are the notes from "${cleanedTitle}" on ${meetingDate}.</p>`;
    body += `<hr/>`;
  }

  // Add summary - convert markdown to HTML preserving formatting (headings, bold, lists)
  const summaryHtml = markdownToHtml(payload.summary || 'No summary provided.');
  body += summaryHtml;

  // Add action items
  if (payload.action_items && payload.action_items.length > 0) {
    body += `<h3>Action Items</h3>`;
    body += `<ol>`;
    payload.action_items.forEach((item, index) => {
      body += `<li>`;
      body += `${item.description || item.text || item}`;
      if (item.assignee) {
        body += ` <em>(Assigned to: ${item.assignee})</em>`;
      }
      if (item.due_date) {
        body += ` <em>(Due: ${item.due_date})</em>`;
      }
      body += `</li>`;
    });
    body += `</ol>`;
  }

  // Get customizable signature from settings
  const userName = props.getProperty('USER_NAME') || 'Team';
  const signatureTemplate = props.getProperty('MEETING_SIGNATURE')
    || 'Did I miss anything?\n\nThanks,\n{user_name}';
  const signature = signatureTemplate.replace('{user_name}', userName);

  // Convert signature newlines to HTML
  body += `<hr/>`;
  body += `<p>${signature.replace(/\n/g, '<br/>')}</p>`;

  // Add metadata for post-send processing (hidden)
  // Client matching will happen when email is sent based on recipient
  body += `<div style="display:none;">`;
  body += `<!--MEETING_TITLE:${payload.meeting_title}-->`;
  body += `<!--MEETING_DATE:${payload.meeting_date}-->`;
  body += `<!--ACTION_ITEMS:${JSON.stringify(payload.action_items || [])}-->`;
  if (payload.fathom_url) {
    body += `<!--FATHOM_URL:${payload.fathom_url}-->`;
  }
  body += `</div>`;

  // Ensure proper UTF-8 encoding by wrapping in HTML structure with charset meta tags
  const fullHtmlBody = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Meeting Summary</title>
</head>
<body>
${body}
</body>
</html>`;

  // Get current user's email to exclude from recipients
  const myEmail = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase();

  // Get all participant emails from the meeting, excluding the current user
  const participantEmails = (payload.participants || [])
    .map(p => p.email)
    .filter(email => email && email.toLowerCase() !== myEmail);

  // Use participant emails as recipients (comma-separated if multiple)
  const toAddress = participantEmails.length > 0
    ? participantEmails.join(', ')
    : myEmail; // Fallback to own email if no other participants

  // Create draft with proper UTF-8 encoding
  const draft = GmailApp.createDraft(toAddress, subject, '', {
    htmlBody: fullHtmlBody
  });

  Logger.log(`Created draft with ID: ${draft.getId()}`);

  // Store draft info for monitoring (client may be null)
  storePendingDraft(draft.getId(), client ? client.client_name : null, payload);

  return draft.getId();
}

/**
 * Stores information about a pending draft for later monitoring.
 *
 * @param {string} draftId - The Gmail draft ID
 * @param {string} clientId - The client ID
 * @param {Object} payload - The original meeting payload
 */
function storePendingDraft(draftId, clientId, payload) {
  const cache = CacheService.getScriptCache();
  const key = `pending_draft_${draftId}`;

  const data = {
    draftId: draftId,
    clientId: clientId,
    meetingTitle: payload.meeting_title,
    meetingDate: payload.meeting_date,
    actionItems: payload.action_items || [],
    summary: payload.summary,
    createdAt: new Date().toISOString()
  };

  // Cache for 24 hours (86400 seconds)
  cache.put(key, JSON.stringify(data), 86400);

  // Also store in a list of pending drafts
  const pendingList = getPendingDraftsList();
  pendingList.push(draftId);
  cache.put('pending_drafts_list', JSON.stringify(pendingList), 86400);
}

/**
 * Gets the list of pending draft IDs from cache.
 *
 * @returns {string[]} Array of draft IDs
 */
function getPendingDraftsList() {
  const cache = CacheService.getScriptCache();
  const listJson = cache.get('pending_drafts_list');
  return listJson ? JSON.parse(listJson) : [];
}


// ============================================================================
// FATHOM API INTEGRATION
// ============================================================================

/**
 * Fetches the latest meeting from Fathom API.
 * This is used for testing the webhook processing without waiting for a real meeting.
 *
 * @returns {Object} The latest meeting data from Fathom
 */
function fetchLatestFathomMeeting() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('FATHOM_API_KEY');

  if (!apiKey) {
    throw new Error('Fathom API key not configured. Add it in Settings.');
  }

  // Fathom API endpoint - docs at https://developers.fathom.ai
  const url = 'https://api.fathom.ai/external/v1/meetings?include_transcript=true&include_summary=true&include_action_items=true';

  const options = {
    method: 'GET',
    headers: {
      'X-Api-Key': apiKey,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      logProcessing('FATHOM_API', null, `API error (${responseCode}): ${responseText.substring(0, 200)}`, 'error');
      throw new Error(`Fathom API error (${responseCode}): ${responseText}`);
    }

    const data = JSON.parse(responseText);

    // Fathom API returns meetings in 'items' array
    if (data.items && data.items.length > 0) {
      logProcessing('FATHOM_API', null, `Found ${data.items.length} meetings`, 'success');
      return data.items[0];
    } else if (Array.isArray(data) && data.length > 0) {
      logProcessing('FATHOM_API', null, `Found ${data.length} meetings (array)`, 'success');
      return data[0];
    }

    // Log what we received so user can see in Processing_Log sheet
    logProcessing('FATHOM_API', null, `No meetings found. Response: ${JSON.stringify(data).substring(0, 300)}`, 'warning');
    throw new Error('No meetings found in Fathom. Check Processing_Log sheet for details.');

  } catch (error) {
    logProcessing('FATHOM_API', null, error.message, 'error');
    throw error;
  }
}

/**
 * Strips hyperlinks from markdown but preserves formatting.
 * Converts [text](url) to just text, removes standalone URLs.
 * Keeps headings, bold, italics, lists intact.
 *
 * @param {string} markdown - The markdown text to clean
 * @returns {string} Markdown without links
 */
function stripMarkdownLinks(markdown) {
  if (!markdown) return '';

  let text = markdown;

  // Remove markdown links: [text](url) -> text (keep the text, remove the link)
  text = text.replace(/\[([^\]]+)\]\([^)]+\)/g, '$1');

  // Remove any standalone URLs
  text = text.replace(/https?:\/\/[^\s)]+/g, '');

  // Clean up extra whitespace
  text = text.replace(/\n{3,}/g, '\n\n');

  return text.trim();
}

/**
 * Converts markdown to HTML for email display.
 * Handles headings, bold, italics, lists, and line breaks.
 *
 * @param {string} markdown - The markdown text
 * @returns {string} HTML formatted text
 */
function markdownToHtml(markdown) {
  if (!markdown) return '';

  let html = markdown;

  // Convert headings: ## Heading -> <h3>Heading</h3>
  html = html.replace(/^#{1,2}\s+(.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^#{3,6}\s+(.+)$/gm, '<h4>$1</h4>');

  // Convert bold: **text** or __text__ -> <strong>text</strong>
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/__([^_]+)__/g, '<strong>$1</strong>');

  // Convert italics: *text* or _text_ -> <em>text</em>
  html = html.replace(/\*([^*]+)\*/g, '<em>$1</em>');
  html = html.replace(/_([^_]+)_/g, '<em>$1</em>');

  // Convert unordered list items: - item -> <li>item</li>
  html = html.replace(/^[-*]\s+(.+)$/gm, '<li>$1</li>');

  // Wrap consecutive <li> items in <ul>
  html = html.replace(/(<li>.*<\/li>\n?)+/g, '<ul>$&</ul>');

  // Convert double newlines to paragraph breaks
  html = html.replace(/\n\n/g, '</p><p>');

  // Convert single newlines to <br>
  html = html.replace(/\n/g, '<br/>');

  // Wrap in paragraph tags
  html = '<p>' + html + '</p>';

  // Clean up empty paragraphs
  html = html.replace(/<p><\/p>/g, '');
  html = html.replace(/<p>(<h[34]>)/g, '$1');
  html = html.replace(/(<\/h[34]>)<\/p>/g, '$1');
  html = html.replace(/<p>(<ul>)/g, '$1');
  html = html.replace(/(<\/ul>)<\/p>/g, '$1');

  return html;
}

/**
 * Normalizes Fathom payload from any source (webhook or API) to a consistent format.
 * Handles both webhook payloads and API responses.
 *
 * @param {Object} payload - The raw payload from Fathom (webhook or API)
 * @returns {Object} Normalized payload with consistent field names
 */
function normalizeFathomPayload(payload) {
  if (!payload) return null;

  // If this looks like an API response (has fields like 'default_summary', 'calendar_invitees'),
  // convert it using the full conversion function
  if (payload.default_summary || payload.calendar_invitees || payload.recorded_by) {
    return convertFathomMeetingToPayload(payload);
  }

  // Otherwise, assume it's a webhook payload and normalize field names
  return {
    meeting_title: payload.meeting_title || payload.title || 'Untitled Meeting',
    meeting_date: payload.meeting_date || payload.start_time || payload.created_at || new Date().toISOString(),
    transcript: payload.transcript || '',
    summary: payload.summary || payload.notes || '',
    action_items: payload.action_items || [],
    participants: payload.participants || payload.attendees || [],
    fathom_url: payload.fathom_url || payload.url || payload.share_url || null,
    meeting_id: payload.meeting_id || payload.id || null
  };
}

/**
 * Converts Fathom API meeting data to webhook payload format.
 * This normalizes the API response to match the expected webhook structure.
 *
 * @param {Object} fathomMeeting - The meeting data from Fathom API
 * @returns {Object} Normalized payload matching webhook format
 */
function convertFathomMeetingToPayload(fathomMeeting) {
  // Map Fathom API response to webhook payload format
  // Fathom API fields: title, created_at, default_summary, transcript (array), action_items, calendar_invitees, recorded_by

  // Extract transcript - Fathom returns array of {speaker: {display_name}, text, timestamp}
  let transcriptText = '';
  if (Array.isArray(fathomMeeting.transcript)) {
    transcriptText = fathomMeeting.transcript
      .map(entry => {
        const speaker = entry.speaker?.display_name || 'Unknown';
        return `${speaker}: ${entry.text}`;
      })
      .join('\n\n');
  } else if (typeof fathomMeeting.transcript === 'string') {
    transcriptText = fathomMeeting.transcript;
  }

  // Extract summary - Fathom uses default_summary.markdown_formatted
  // Check setting to determine if we should keep or strip video timestamp hyperlinks
  const props = PropertiesService.getScriptProperties();
  const keepLinks = props.getProperty('FATHOM_KEEP_LINKS') === 'true';

  let summaryText = '';
  if (fathomMeeting.default_summary && fathomMeeting.default_summary.markdown_formatted) {
    summaryText = keepLinks
      ? fathomMeeting.default_summary.markdown_formatted
      : stripMarkdownLinks(fathomMeeting.default_summary.markdown_formatted);
  } else if (typeof fathomMeeting.summary === 'string') {
    summaryText = keepLinks
      ? fathomMeeting.summary
      : stripMarkdownLinks(fathomMeeting.summary);
  } else if (fathomMeeting.summary && fathomMeeting.summary.markdown_formatted) {
    summaryText = keepLinks
      ? fathomMeeting.summary.markdown_formatted
      : stripMarkdownLinks(fathomMeeting.summary.markdown_formatted);
  }

  // Fathom uses calendar_invitees for participants
  // Each has: name, email, is_external
  const participants = fathomMeeting.calendar_invitees || fathomMeeting.attendees || fathomMeeting.participants || [];

  // Include recorded_by as a participant if available
  if (fathomMeeting.recorded_by && fathomMeeting.recorded_by.email) {
    const recorderExists = participants.some(p => p.email === fathomMeeting.recorded_by.email);
    if (!recorderExists) {
      participants.push({
        name: fathomMeeting.recorded_by.name,
        email: fathomMeeting.recorded_by.email,
        is_external: false
      });
    }
  }

  return {
    meeting_title: fathomMeeting.title || fathomMeeting.meeting_title || 'Untitled Meeting',
    meeting_date: fathomMeeting.created_at || fathomMeeting.scheduled_start_time || new Date().toISOString(),
    transcript: transcriptText,
    summary: summaryText,
    action_items: fathomMeeting.action_items || [],
    participants: participants,
    fathom_url: fathomMeeting.url || fathomMeeting.share_url || null
  };
}

/**
 * Menu function to load and process the latest meeting from Fathom.
 * This simulates receiving a webhook with the latest meeting data.
 */
function loadLatestFathomMeeting() {
  const ui = SpreadsheetApp.getUi();

  try {
    ui.alert('Loading Meeting', 'Fetching latest meeting from Fathom...', ui.ButtonSet.OK);

    // Fetch latest meeting
    const fathomMeeting = fetchLatestFathomMeeting();

    // Convert to webhook payload format
    const payload = convertFathomMeetingToPayload(fathomMeeting);

    // Show confirmation with meeting details
    const confirmResult = ui.alert(
      'Meeting Found',
      `Found meeting: "${payload.meeting_title}"\n` +
      `Date: ${payload.meeting_date}\n` +
      `Participants: ${payload.participants.length}\n\n` +
      'Process this meeting as if it were a webhook?',
      ui.ButtonSet.YES_NO
    );

    if (confirmResult === ui.Button.YES) {
      // Process the meeting using the same flow as webhooks
      const result = processFathomWebhook(payload);

      ui.alert(
        'Processing Complete',
        `Meeting processed successfully!\n\n` +
        `Client: ${result.client_name || 'Not matched'}\n` +
        `Draft created: ${result.draft_id ? 'Yes' : 'No'}`,
        ui.ButtonSet.OK
      );
    }

  } catch (error) {
    ui.alert('Error', `Failed to load meeting: ${error.message}`, ui.ButtonSet.OK);
    Logger.log(`loadLatestFathomMeeting error: ${error.message}`);
  }
}

/**
 * Checks if a Fathom meeting has already been processed.
 * Uses spreadsheet tracking instead of cache for visibility.
 *
 * @param {string} meetingId - The Fathom meeting ID
 * @returns {boolean} True if already processed
 */
function isFathomMeetingProcessed(meetingId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PROCESSED_FATHOM);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PROCESSED_FATHOM);
    sheet.appendRow(['Meeting ID', 'Meeting Title', 'Meeting Date', 'Processed At', 'Client Name', 'Draft ID']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }

  const data = sheet.getDataRange().getValues();

  const normalizedMeetingId = normalizeFathomMeetingId(meetingId);

  // Check if meeting ID exists in column A (skip header row)
  for (let i = 1; i < data.length; i++) {
    if (normalizeFathomMeetingId(data[i][0]) === normalizedMeetingId) {
      return true;
    }
  }

  return false;
}

/**
 * Records that a Fathom meeting has been processed.
 *
 * @param {string} meetingId - The Fathom meeting ID
 * @param {string} meetingTitle - The meeting title
 * @param {string} meetingDate - The meeting date
 * @param {string|null} clientName - The matched client name (or null)
 * @param {string} draftId - The created draft ID
 */
function recordFathomMeetingProcessed(meetingId, meetingTitle, meetingDate, clientName, draftId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PROCESSED_FATHOM);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PROCESSED_FATHOM);
    sheet.appendRow(['Meeting ID', 'Meeting Title', 'Meeting Date', 'Processed At', 'Client Name', 'Draft ID']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }

  sheet.appendRow([
    normalizeFathomMeetingId(meetingId),
    meetingTitle,
    meetingDate,
    new Date().toISOString(),
    clientName || 'No client match',
    draftId || 'N/A'
  ]);
}

/**
 * Normalizes Fathom meeting IDs for consistent comparisons/storage.
 * Handles numeric values stored in sheets and ensures trim.
 *
 * @param {string|number|null} meetingId - The meeting ID to normalize
 * @returns {string} Normalized meeting ID
 */
function normalizeFathomMeetingId(meetingId) {
  if (meetingId === null || meetingId === undefined) {
    return '';
  }
  return String(meetingId).trim();
}

/**
 * Polls Fathom API for new meetings and processes them automatically.
 * This is a backup mechanism in case webhooks fail.
 * Should be called periodically (e.g., every 30 minutes).
 */
function pollFathomForNewMeetings() {
  logProcessing('FATHOM_POLL', null, 'Starting Fathom polling check', 'info');

  const apiKey = PropertiesService.getScriptProperties().getProperty('FATHOM_API_KEY');

  if (!apiKey) {
    logProcessing('FATHOM_POLL', null, 'Fathom API key not configured - skipping poll', 'warning');
    return;
  }

  try {
    // Fetch latest meetings from Fathom
    const url = 'https://api.fathom.ai/external/v1/meetings?include_transcript=true&include_summary=true&include_action_items=true&limit=10';

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
      logProcessing('FATHOM_POLL', null, `API error (${responseCode})`, 'error');
      return;
    }

    const data = JSON.parse(response.getContentText());
    const meetings = data.items || (Array.isArray(data) ? data : []);

    if (meetings.length === 0) {
      logProcessing('FATHOM_POLL', null, 'No meetings found', 'info');
      return;
    }

    logProcessing('FATHOM_POLL', null, `Found ${meetings.length} recent meetings`, 'info');

    // Check each meeting to see if we've already processed it
    let newMeetingsCount = 0;
    let skippedCount = 0;
    let missingIdCount = 0;

    for (const meeting of meetings) {
      // Extract meeting ID from various sources
      let meetingId = meeting.id || meeting.meeting_id;

      // If no direct ID field, try extracting from URL
      if (!meetingId && meeting.url) {
        // Extract numeric ID from URL: https://fathom.video/calls/553083152
        const urlMatch = meeting.url.match(/\/calls\/(\d+)/);
        if (urlMatch && urlMatch[1]) {
          meetingId = urlMatch[1];
          logProcessing('FATHOM_POLL', null, `Extracted meeting ID "${meetingId}" from URL for "${meeting.title || 'Unknown'}"`, 'info');
        }
      }

      // Fall back to recording_id if still no ID
      if (!meetingId && meeting.recording_id) {
        meetingId = meeting.recording_id;
        logProcessing('FATHOM_POLL', null, `Using recording_id "${meetingId}" for "${meeting.title || 'Unknown'}"`, 'info');
      }

      if (!meetingId) {
        missingIdCount++;
        const title = meeting.title || meeting.meeting_title || 'Unknown';
        logProcessing('FATHOM_POLL', null, `Skipped meeting "${title}" - no ID field found. Available fields: ${Object.keys(meeting).join(', ')}`, 'warning');
        continue; // Skip meetings without IDs
      }

      // Check if we've already processed this meeting (using SPREADSHEET tracking)
      if (isFathomMeetingProcessed(meetingId)) {
        skippedCount++;
        continue; // Skip already processed meetings
      }

      // Convert to webhook payload format and process
      try {
        const payload = convertFathomMeetingToPayload(meeting);
        logProcessing('FATHOM_POLL', null, `Processing new meeting: ${payload.meeting_title}`, 'info');

        const result = processFathomWebhook(payload);

        // Record as processed in spreadsheet
        recordFathomMeetingProcessed(
          meetingId,
          payload.meeting_title,
          payload.meeting_date,
          result.client_name,
          result.draft_id
        );

        newMeetingsCount++;

        logProcessing('FATHOM_POLL', result.client_name, `Successfully processed: ${payload.meeting_title}`, 'success');
      } catch (error) {
        logProcessing('FATHOM_POLL', null, `Failed to process meeting ${meetingId}: ${error.message}`, 'error');
      }
    }

    if (newMeetingsCount > 0) {
      logProcessing('FATHOM_POLL', null, `Processed ${newMeetingsCount} new meetings (skipped ${skippedCount} already processed, ${missingIdCount} missing IDs)`, 'success');
    } else {
      const reason = missingIdCount > 0
        ? `${missingIdCount} missing IDs, ${skippedCount} already processed`
        : `${skippedCount} already processed`;
      logProcessing('FATHOM_POLL', null, `No new meetings to process (${reason})`, 'info');
    }

  } catch (error) {
    logProcessing('FATHOM_POLL', null, `Polling failed: ${error.message}`, 'error');
  }
}
