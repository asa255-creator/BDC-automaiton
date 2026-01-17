/**
 * MeetingAutomation.gs - Fathom Webhook Handling and Post-Send Processing
 *
 * This module handles the meeting automation workflow:
 * 1. Receives Fathom webhooks when meetings end
 * 2. Creates draft meeting summary emails
 * 3. Monitors for sent meeting summaries
 * 4. Creates Todoist tasks from action items
 * 5. Appends meeting notes to client Google Docs
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
 * Creates a Gmail draft with the meeting summary.
 * Works with or without a matched client.
 *
 * @param {Object} payload - The Fathom webhook payload
 * @param {Object|null} client - The matched client object (or null if no match)
 * @returns {string} The draft ID
 */
function createMeetingSummaryDraft(payload, client) {
  const meetingDate = formatDateShort(new Date(payload.meeting_date));

  // Build subject and greeting based on whether we have a client
  const clientName = client ? client.client_name : '[ADD CLIENT NAME]';
  const subject = `Team ${clientName} - Here are the notes from the meeting "${payload.meeting_title}" ${meetingDate}`;

  // Build email body
  let body = `<p>Team ${clientName} -</p>`;
  body += `<p>Here are the notes from the meeting "${payload.meeting_title}" ${meetingDate}.</p>`;

  // Add Fathom link if available
  if (payload.fathom_url) {
    body += `<p><a href="${payload.fathom_url}">View full meeting recording</a></p>`;
  }

  body += `<hr/>`;

  // Add summary
  body += `<h3>Summary</h3>`;
  body += `<p>${payload.summary || 'No summary provided.'}</p>`;

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

  // Closing
  body += `<hr/>`;
  body += `<p>Did I miss anything?</p>`;
  const userName = PropertiesService.getScriptProperties().getProperty('USER_NAME') || 'Team';
  body += `<p>Thanks,<br/>${userName}</p>`;

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

  // Determine recipient
  let toAddress = '';
  if (client) {
    const recipients = parseCommaSeparatedList(client.contact_emails);
    toAddress = recipients.length > 0 ? recipients[0] : '';
  }
  // If no client or no contact email, leave To: empty so user must fill it in

  // Create draft
  const draft = GmailApp.createDraft(toAddress, subject, '', {
    htmlBody: body
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
// SENT EMAIL MONITORING
// ============================================================================

/**
 * Monitors for sent meeting summary emails and processes them.
 * Called by the 10-minute trigger.
 */
function monitorSentMeetingSummaries() {
  Logger.log('Checking for sent meeting summaries...');

  // Search for recently sent meeting summary emails
  const query = 'from:me subject:"Meeting Summary" newer_than:1h';
  const threads = GmailApp.search(query, 0, 20);

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const message of messages) {
      // Only process sent messages
      if (message.getFrom().indexOf(getCurrentUserEmail()) === -1) {
        continue;
      }

      // Check if already processed
      if (isMessageProcessed(message.getId())) {
        continue;
      }

      // Process the sent summary
      processSentMeetingSummary(message);
    }
  }
}

/**
 * Processes a sent meeting summary email.
 *
 * @param {GmailMessage} message - The sent Gmail message
 */
function processSentMeetingSummary(message) {
  Logger.log(`Processing sent meeting summary: ${message.getSubject()}`);

  // Extract metadata from email body
  const body = message.getBody();
  const clientId = extractMetadata(body, 'CLIENT_ID');
  const actionItemsJson = extractMetadata(body, 'ACTION_ITEMS');

  if (!clientId) {
    Logger.log('No client ID found in email metadata');
    return;
  }

  const client = getClientById(clientId);
  if (!client) {
    Logger.log(`Client not found: ${clientId}`);
    return;
  }

  // Parse action items
  let actionItems = [];
  try {
    actionItems = actionItemsJson ? JSON.parse(actionItemsJson) : [];
  } catch (e) {
    Logger.log('Failed to parse action items JSON');
  }

  // Create Todoist tasks
  if (actionItems.length > 0 && client.todoist_project_id) {
    createTodoistTasks(actionItems, client);
  }

  // Append meeting notes to client's Google Doc
  if (client.google_doc_url) {
    appendMeetingNotesToDoc(message, client);
  }

  // Apply sub-label to the sent email
  applyMeetingSummaryLabel(message, client);

  // Mark as processed
  markMessageProcessed(message.getId());

  logProcessing(
    'SUMMARY_SENT',
    clientId,
    `Processed sent summary: ${message.getSubject()}`,
    'success'
  );
}

/**
 * Extracts metadata from HTML comment tags in email body.
 *
 * @param {string} body - The email body HTML
 * @param {string} key - The metadata key to extract
 * @returns {string|null} The extracted value or null
 */
function extractMetadata(body, key) {
  const regex = new RegExp(`<!--${key}:(.+?)-->`);
  const match = body.match(regex);
  return match ? match[1] : null;
}

/**
 * Checks if a message has already been processed.
 *
 * @param {string} messageId - The Gmail message ID
 * @returns {boolean} True if already processed
 */
function isMessageProcessed(messageId) {
  const cache = CacheService.getScriptCache();
  return cache.get(`processed_${messageId}`) !== null;
}

/**
 * Marks a message as processed.
 *
 * @param {string} messageId - The Gmail message ID
 */
function markMessageProcessed(messageId) {
  const cache = CacheService.getScriptCache();
  // Cache for 7 days (604800 seconds)
  cache.put(`processed_${messageId}`, 'true', 604800);
}

// ============================================================================
// TODOIST INTEGRATION
// ============================================================================

/**
 * Creates Todoist tasks for action items.
 *
 * @param {Object[]} actionItems - Array of action item objects
 * @param {Object} client - The client object
 */
function createTodoistTasks(actionItems, client) {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    Logger.log('Todoist API token not configured');
    return;
  }

  const projectId = client.todoist_project_id;

  for (const item of actionItems) {
    try {
      createTodoistTask(apiToken, projectId, item, client.client_name);
    } catch (error) {
      Logger.log(`Failed to create Todoist task: ${error.message}`);
      logProcessing(
        'TODOIST_ERROR',
        client.client_name,
        `Failed to create task: ${item.description}`,
        'error'
      );
    }
  }
}

/**
 * Creates a single Todoist task.
 *
 * @param {string} apiToken - Todoist API token
 * @param {string} projectId - Todoist project ID
 * @param {Object} item - Action item object
 * @param {string} clientName - Client name for task content
 */
function createTodoistTask(apiToken, projectId, item, clientName) {
  const url = 'https://api.todoist.com/rest/v2/tasks';

  const taskContent = `[${clientName}] ${item.description}`;

  const payload = {
    content: taskContent,
    project_id: projectId
  };

  // Add due date if provided
  if (item.due_date) {
    payload.due_string = item.due_date;
  }

  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(`Todoist API error: ${responseCode}`);
  }

  Logger.log(`Created Todoist task: ${taskContent}`);
}

/**
 * Fetches tasks from Todoist for a specific project.
 *
 * @param {string} projectId - Todoist project ID
 * @returns {Object[]} Array of task objects
 */
function fetchTodoistTasks(projectId) {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    Logger.log('Todoist API token not configured');
    return [];
  }

  const url = `https://api.todoist.com/rest/v2/tasks?project_id=${projectId}`;

  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${apiToken}`
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`Todoist API error: ${responseCode}`);
      return [];
    }

    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log(`Failed to fetch Todoist tasks: ${error.message}`);
    return [];
  }
}

/**
 * Fetches tasks due today or overdue for a project.
 *
 * @param {string} projectId - Todoist project ID
 * @returns {Object[]} Array of task objects due today or overdue
 */
function fetchTodoistTasksDueToday(projectId) {
  const tasks = fetchTodoistTasks(projectId);
  const today = new Date();
  today.setHours(23, 59, 59, 999);

  return tasks.filter(task => {
    if (!task.due) return false;

    const dueDate = new Date(task.due.date);
    return dueDate <= today;
  });
}

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
    Logger.log('No Google Doc URL configured for client');
    return;
  }

  try {
    const docId = extractDocIdFromUrl(client.google_doc_url);
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Get meeting details from message
    const subject = message.getSubject();
    const date = formatDate(message.getDate());

    // Add section header
    body.appendParagraph(`Meeting Notes - ${date}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Convert email HTML to plain text and append
    const emailBody = message.getPlainBody();
    body.appendParagraph(emailBody);

    // Add separator
    body.appendParagraph('---');
    body.appendParagraph('');

    doc.saveAndClose();

    Logger.log(`Appended meeting notes to doc: ${client.google_doc_url}`);
  } catch (error) {
    Logger.log(`Failed to append to Google Doc: ${error.message}`);
    logProcessing(
      'DOC_APPEND_ERROR',
      client.client_name,
      `Failed to append meeting notes: ${error.message}`,
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
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) {
    return match[1];
  }
  // Assume it's already a doc ID if not a URL
  return url;
}

// ============================================================================
// LABEL APPLICATION
// ============================================================================

/**
 * Applies the Meeting Summaries sub-label to a sent message.
 *
 * @param {GmailMessage} message - The Gmail message
 * @param {Object} client - The client object
 */
function applyMeetingSummaryLabel(message, client) {
  const labelName = `Client: ${client.client_name}/Meeting Summaries`;

  try {
    let label = GmailApp.getUserLabelByName(labelName);

    if (!label) {
      // Create the label if it doesn't exist
      label = GmailApp.createLabel(labelName);
    }

    // Apply label to the thread
    const thread = message.getThread();
    thread.addLabel(label);

    Logger.log(`Applied label: ${labelName}`);
  } catch (error) {
    Logger.log(`Failed to apply label: ${error.message}`);
  }
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
  // include_transcript and include_summary require first-party API keys
  const url = 'https://api.fathom.ai/external/v1/meetings?include_transcript=true&include_summary=true';

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
 * Converts Fathom API meeting data to webhook payload format.
 * This normalizes the API response to match the expected webhook structure.
 *
 * @param {Object} fathomMeeting - The meeting data from Fathom API
 * @returns {Object} Normalized payload matching webhook format
 */
function convertFathomMeetingToPayload(fathomMeeting) {
  // Map Fathom API response to webhook payload format
  // Fathom API returns: title, created_at, transcript, summary, action_items, calendar_invitees, recorded_by

  // Extract transcript text - may be string or object with text property
  let transcriptText = '';
  if (typeof fathomMeeting.transcript === 'string') {
    transcriptText = fathomMeeting.transcript;
  } else if (fathomMeeting.transcript && fathomMeeting.transcript.text) {
    transcriptText = fathomMeeting.transcript.text;
  }

  // Extract summary text - may be string or object
  let summaryText = '';
  if (typeof fathomMeeting.summary === 'string') {
    summaryText = fathomMeeting.summary;
  } else if (fathomMeeting.summary && fathomMeeting.summary.text) {
    summaryText = fathomMeeting.summary.text;
  } else if (fathomMeeting.summary && fathomMeeting.summary.content) {
    summaryText = fathomMeeting.summary.content;
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
    action_items: fathomMeeting.action_items || fathomMeeting.tasks || [],
    participants: participants,
    fathom_url: fathomMeeting.url || null
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
        `Client: ${result.client || 'Not matched'}\n` +
        `Draft created: ${result.draftCreated ? 'Yes' : 'No'}`,
        ui.ButtonSet.OK
      );
    }

  } catch (error) {
    ui.alert('Error', `Failed to load meeting: ${error.message}`, ui.ButtonSet.OK);
    Logger.log(`loadLatestFathomMeeting error: ${error.message}`);
  }
}
