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
  const props = PropertiesService.getScriptProperties();
  const meetingDate = formatDateShort(new Date(payload.meeting_date));

  // Build subject and greeting based on whether we have a client
  const clientName = client ? client.client_name : '[ADD CLIENT NAME]';

  // Get customizable subject template from settings
  const subjectTemplate = props.getProperty('MEETING_SUBJECT_TEMPLATE')
    || 'Team {client_name} - Meeting notes from "{meeting_title}" {date}';
  const subject = subjectTemplate
    .replace('{client_name}', clientName)
    .replace('{meeting_title}', payload.meeting_title)
    .replace('{date}', meetingDate);

  // Build email body with greeting that matches subject style
  let body = `<p>Team ${clientName} -</p>`;
  body += `<p>Here are the notes from the meeting "${payload.meeting_title}" ${meetingDate}.</p>`;
  body += `<hr/>`;

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
 * Monitors for sent meeting summary emails by checking for new conversations
 * in client Meeting Summaries labels.
 * Called by the 10-minute trigger.
 */
function monitorSentMeetingSummaries() {
  logProcessing('SENT_MONITOR', null, 'Checking for new meeting summaries in labeled folders...', 'info');

  // Get all clients with setup_complete
  const allClients = getClientRegistry();
  const clients = allClients.filter(client => client.setup_complete === true);

  if (clients.length === 0) {
    logProcessing('SENT_MONITOR', null, 'No clients with setup_complete found', 'warning');
    return;
  }

  let totalProcessed = 0;

  // Check each client's Meeting Summaries label for new conversations
  for (const client of clients) {
    const labelName = `Client: ${client.client_name}/Meeting Summaries`;

    try {
      const label = GmailApp.getUserLabelByName(labelName);
      if (!label) {
        continue; // Label doesn't exist yet
      }

      // Get threads with this label from the last hour
      const threads = label.getThreads(0, 20);

      for (const thread of threads) {
        const messages = thread.getMessages();
        if (messages.length === 0) continue;

        // Only process the FIRST message in the thread (not replies)
        const firstMessage = messages[0];

        // Check if already processed
        if (isMessageProcessed(firstMessage.getId())) {
          continue;
        }

        // Verify it's from me (I sent it)
        const myEmail = getCurrentUserEmail();
        if (!firstMessage.getFrom().toLowerCase().includes(myEmail.toLowerCase())) {
          continue;
        }

        // Check if this is a new thread (sent within last hour)
        const sentTime = firstMessage.getDate();
        const oneHourAgo = new Date(Date.now() - 60 * 60 * 1000);
        if (sentTime < oneHourAgo) {
          // Mark old messages as processed to skip them in future
          markMessageProcessed(firstMessage.getId());
          continue;
        }

        // Process the sent summary
        try {
          logProcessing('SENT_MONITOR', client.client_name, `Found new summary: ${firstMessage.getSubject()}`, 'info');
          processSentMeetingSummary(firstMessage, client);
          totalProcessed++;
        } catch (error) {
          logProcessing('SENT_MONITOR', client.client_name, `Error processing: ${error.message}`, 'error');
        }
      }

    } catch (error) {
      logProcessing('SENT_MONITOR', client.client_name, `Error checking label: ${error.message}`, 'error');
    }
  }

  logProcessing('SENT_MONITOR', null, `Processed ${totalProcessed} new meeting summaries`, 'success');
}

/**
 * Processes a sent meeting summary email.
 * Extracts action items from the email body (not metadata) since user may have edited.
 *
 * @param {GmailMessage} message - The sent Gmail message
 * @param {Object} client - The client object (already identified from label)
 */
function processSentMeetingSummary(message, client) {
  const subject = message.getSubject();
  logProcessing('SUMMARY_PROCESS', client.client_name, `Processing: ${subject}`, 'info');

  // Extract action items from the email body using AI
  // This is critical because user may have edited action items before sending
  const emailBody = message.getPlainBody();
  const actionItems = extractActionItemsWithAI(emailBody, client);

  // Create Todoist tasks if we have action items and a project
  if (actionItems.length > 0 && client.todoist_project_id) {
    logProcessing('SUMMARY_PROCESS', client.client_name, `Found ${actionItems.length} action items`, 'info');
    createTodoistTasksWithAssignees(actionItems, client);
  } else if (actionItems.length === 0) {
    logProcessing('SUMMARY_PROCESS', client.client_name, 'No action items found in email', 'info');
  }

  // Append meeting notes to client's Google Doc with proper separators
  if (client.google_doc_url) {
    appendMeetingNotesToDoc(message, client);
  }

  // Mark as processed
  markMessageProcessed(message.getId());

  logProcessing('SUMMARY_PROCESS', client.client_name, `Completed processing: ${subject}`, 'success');
}

/**
 * Identifies a client from a comma-separated list of email addresses.
 *
 * @param {string} emailAddresses - Comma-separated email addresses
 * @returns {Object|null} The matched client or null
 */
function identifyClientFromEmailAddresses(emailAddresses) {
  if (!emailAddresses) return null;

  // Parse email addresses (can be "Name <email>" format)
  const emails = emailAddresses.split(',').map(addr => {
    const match = addr.match(/<([^>]+)>/);
    return match ? match[1].trim().toLowerCase() : addr.trim().toLowerCase();
  }).filter(e => e);

  // Try to match against client registry
  const clients = getClientRegistry();

  for (const client of clients) {
    // Check contact emails
    const contactEmails = parseCommaSeparatedList(client.contact_emails)
      .map(e => e.toLowerCase());

    for (const email of emails) {
      if (contactEmails.includes(email)) {
        return client;
      }
    }

    // Check email domains
    const domains = parseCommaSeparatedList(client.email_domains)
      .map(d => d.toLowerCase());

    for (const email of emails) {
      const emailDomain = email.split('@')[1];
      if (emailDomain && domains.includes(emailDomain)) {
        return client;
      }
    }
  }

  return null;
}

/**
 * Extracts action items from email body using Claude AI.
 * This parses the actual sent email content, respecting any edits the user made.
 *
 * @param {string} emailBody - The plain text email body
 * @param {Object} client - The client object
 * @returns {Object[]} Array of structured action items
 */
function extractActionItemsWithAI(emailBody, client) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    logProcessing('AI_EXTRACT', client.client_name, 'Claude API key not configured - skipping AI extraction', 'warning');
    // Fallback: try to extract manually
    return extractActionItemsManually(emailBody);
  }

  // Fetch project collaborators for assignee matching
  let collaborators = [];
  if (client.todoist_project_id) {
    collaborators = fetchProjectCollaborators(client.todoist_project_id);
    logProcessing('AI_EXTRACT', client.client_name, `Found ${collaborators.length} project collaborators`, 'info');
  }

  // Build collaborators JSON for the prompt
  const collaboratorsJson = JSON.stringify(collaborators.map(c => ({
    id: c.id,
    name: c.full_name || c.name,
    email: c.email
  })));

  const today = new Date().toISOString().split('T')[0];

  const prompt = `You are a specialized data processing tool designed to extract action items from meeting summary emails.

Here is the meeting summary email:
---
${emailBody}
---

Here are the project collaborators who can be assigned tasks:
${collaboratorsJson}

Today's date is: ${today}

### Your Task:
1. Find all action items mentioned in the email (usually in a numbered list or "Action Items" section)
2. For each action item, extract:
   - title: A concise title (max 100 chars)
   - description: The full action item text
   - assignee_id: Match the assignee name to a collaborator ID, or null if no match
   - assignee_name: The name mentioned in the action item, or null
   - due_date: In YYYY-MM-DD format. Use context clues like "next Monday", "by Friday". If no date specified, set to one week from today.

### Output Format:
Return ONLY valid JSON (no markdown, no explanation):
{
  "tasks": [
    {
      "title": "...",
      "description": "...",
      "assignee_id": "...",
      "assignee_name": "...",
      "due_date": "YYYY-MM-DD"
    }
  ]
}

If no action items found, return: {"tasks": []}`;

  try {
    const url = 'https://api.anthropic.com/v1/messages';

    // Use dynamic model (prefer sonnet for better extraction quality)
    const models = fetchAvailableModelsFromAPI(false);
    const sonnet = models.find(m => m.id.includes('sonnet'));
    const model = sonnet ? sonnet.id : models[0]?.id || FALLBACK_MODELS[0].id;

    const payload = {
      model: model,
      max_tokens: 2000,
      messages: [{ role: 'user', content: prompt }]
    };

    const options = {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      logProcessing('AI_EXTRACT', client.client_name, `Claude API error: ${responseCode}`, 'error');
      return extractActionItemsManually(emailBody);
    }

    const result = JSON.parse(response.getContentText());

    if (result.content && result.content.length > 0) {
      const aiResponse = result.content[0].text;

      // Parse the JSON response
      try {
        const parsed = JSON.parse(aiResponse);
        logProcessing('AI_EXTRACT', client.client_name, `AI extracted ${parsed.tasks?.length || 0} action items`, 'success');
        return parsed.tasks || [];
      } catch (parseError) {
        logProcessing('AI_EXTRACT', client.client_name, `Failed to parse AI response: ${parseError.message}`, 'error');
        return extractActionItemsManually(emailBody);
      }
    }

    return [];

  } catch (error) {
    logProcessing('AI_EXTRACT', client.client_name, `AI extraction failed: ${error.message}`, 'error');
    return extractActionItemsManually(emailBody);
  }
}

/**
 * Fallback: Extract action items manually from email body without AI.
 *
 * @param {string} emailBody - The plain text email body
 * @returns {Object[]} Array of action items
 */
function extractActionItemsManually(emailBody) {
  const actionItems = [];

  // Look for numbered items after "Action Items" header
  const actionSection = emailBody.match(/Action Items[\s\S]*?(?=\n\n|\n---|\n#|$)/i);
  if (actionSection) {
    const items = actionSection[0].match(/\d+\.\s+(.+)/g);
    if (items) {
      for (const item of items) {
        const text = item.replace(/^\d+\.\s+/, '').trim();
        actionItems.push({
          title: text.substring(0, 100),
          description: text,
          assignee_id: null,
          assignee_name: null,
          due_date: getOneWeekFromNow()
        });
      }
    }
  }

  return actionItems;
}

/**
 * Gets date one week from now in YYYY-MM-DD format.
 *
 * @returns {string} Date string
 */
function getOneWeekFromNow() {
  const date = new Date();
  date.setDate(date.getDate() + 7);
  return date.toISOString().split('T')[0];
}

/**
 * Fetches collaborators for a Todoist project.
 *
 * @param {string} projectId - The Todoist project ID
 * @returns {Object[]} Array of collaborator objects
 */
function fetchProjectCollaborators(projectId) {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    return [];
  }

  try {
    const url = `https://api.todoist.com/rest/v2/projects/${projectId}/collaborators`;

    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiToken}`
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      logProcessing('TODOIST', null, `Failed to fetch collaborators: ${responseCode}`, 'error');
      return [];
    }

    return JSON.parse(response.getContentText());

  } catch (error) {
    logProcessing('TODOIST', null, `Error fetching collaborators: ${error.message}`, 'error');
    return [];
  }
}

/**
 * Creates Todoist tasks with assignee matching.
 *
 * @param {Object[]} actionItems - Array of action items from AI extraction
 * @param {Object} client - The client object
 */
function createTodoistTasksWithAssignees(actionItems, client) {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    logProcessing('TODOIST', client.client_name, 'Todoist API token not configured', 'error');
    return;
  }

  const projectId = client.todoist_project_id;
  let createdCount = 0;

  for (const item of actionItems) {
    try {
      const url = 'https://api.todoist.com/rest/v2/tasks';

      const payload = {
        content: item.title || item.description.substring(0, 100),
        description: item.description,
        project_id: projectId
      };

      // Add assignee if we have one
      if (item.assignee_id) {
        payload.assignee_id = item.assignee_id;
      }

      // Add due date
      if (item.due_date) {
        payload.due_date = item.due_date;
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

      if (responseCode === 200) {
        createdCount++;
        const assigneeInfo = item.assignee_name ? ` (assigned to ${item.assignee_name})` : '';
        logProcessing('TODOIST', client.client_name, `Created task: ${item.title}${assigneeInfo}`, 'success');
      } else {
        logProcessing('TODOIST', client.client_name, `Failed to create task: ${responseCode}`, 'error');
      }

    } catch (error) {
      logProcessing('TODOIST', client.client_name, `Error creating task: ${error.message}`, 'error');
    }
  }

  logProcessing('TODOIST', client.client_name, `Created ${createdCount}/${actionItems.length} tasks`, 'success');
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
    logProcessing('DOC_APPEND', client.client_name, 'No Google Doc URL configured', 'warning');
    return;
  }

  try {
    const docId = extractDocIdFromUrl(client.google_doc_url);
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Get meeting details from message
    const subject = message.getSubject();
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
  // Strip markdown links and formatting for cleaner email
  let summaryText = '';
  if (fathomMeeting.default_summary && fathomMeeting.default_summary.markdown_formatted) {
    summaryText = stripMarkdownLinks(fathomMeeting.default_summary.markdown_formatted);
  } else if (typeof fathomMeeting.summary === 'string') {
    summaryText = stripMarkdownLinks(fathomMeeting.summary);
  } else if (fathomMeeting.summary && fathomMeeting.summary.markdown_formatted) {
    summaryText = stripMarkdownLinks(fathomMeeting.summary.markdown_formatted);
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
