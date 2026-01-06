/**
 * AgendaGeneration.gs - Agenda Creation Logic
 *
 * This module handles automated meeting agenda generation:
 * - Checks for upcoming meetings without agendas
 * - Gathers context from Todoist, Gmail, and previous meeting notes
 * - Generates agendas using Claude API
 * - Sends agenda emails and appends to client Google Docs
 */

// ============================================================================
// MAIN AGENDA GENERATION
// ============================================================================

/**
 * Main function to generate agendas for today's upcoming meetings.
 * Called hourly during business hours (8 AM - 6 PM).
 */
function generateAgendas() {
  Logger.log('Starting agenda generation...');

  // Get today's remaining calendar events
  const events = getTodaysRemainingEvents();

  if (events.length === 0) {
    Logger.log('No remaining events for today');
    return;
  }

  Logger.log(`Found ${events.length} remaining events`);

  for (const event of events) {
    processEventForAgenda(event);
  }

  Logger.log('Agenda generation completed');
}

/**
 * Processes a single calendar event for potential agenda generation.
 *
 * @param {CalendarEvent} event - The calendar event
 */
function processEventForAgenda(event) {
  const eventId = event.getId();
  const eventTitle = event.getTitle();

  Logger.log(`Processing event: ${eventTitle}`);

  // Identify client from event guests
  const client = identifyClientFromCalendarEvent(event);

  if (!client) {
    // Log to unmatched and skip
    const guestEmails = event.getGuestList().map(g => g.getEmail());
    logUnmatched(
      'meeting',
      `Calendar Event: ${eventTitle} at ${event.getStartTime()}`,
      guestEmails
    );
    Logger.log(`No client match for event: ${eventTitle}`);
    return;
  }

  // Ensure running meeting notes doc is attached to the event
  if (client.google_doc_url) {
    attachDocToEventIfMissing(event, client);
  }

  // Check if agenda already generated
  if (isAgendaGenerated(eventId)) {
    Logger.log(`Agenda already generated for: ${eventTitle}`);
    return;
  }

  // Generate agenda for this event
  try {
    generateAgendaForEvent(event, client);
  } catch (error) {
    Logger.log(`Failed to generate agenda for ${eventTitle}: ${error.message}`);
    logProcessing(
      'AGENDA_ERROR',
      client.client_id,
      `Failed to generate agenda: ${error.message}`,
      'error'
    );
  }
}

/**
 * Generates an agenda for a specific event and client.
 *
 * @param {CalendarEvent} event - The calendar event
 * @param {Object} client - The matched client
 */
function generateAgendaForEvent(event, client) {
  Logger.log(`Generating agenda for: ${event.getTitle()} (${client.client_name})`);

  // Gather context
  const context = gatherAgendaContext(event, client);

  // Generate agenda with Claude
  const agendaContent = generateAgendaWithClaude(event, client, context);

  if (!agendaContent) {
    Logger.log('Failed to generate agenda content from Claude');
    return;
  }

  // Send agenda email
  sendAgendaEmail(event, client, agendaContent);

  // Append to client's Google Doc
  appendAgendaToDoc(event, client, agendaContent);

  // Record the generation
  recordGeneratedAgenda(event, client);

  logProcessing(
    'AGENDA_GENERATED',
    client.client_id,
    `Generated agenda for: ${event.getTitle()}`,
    'success'
  );
}

// ============================================================================
// CONTEXT GATHERING
// ============================================================================

/**
 * Gathers context for agenda generation.
 *
 * @param {CalendarEvent} event - The calendar event
 * @param {Object} client - The client object
 * @returns {Object} Context object with tasks, emails, and meeting history
 */
function gatherAgendaContext(event, client) {
  const context = {
    todoistTasks: [],
    recentEmails: [],
    previousMeetingNotes: null,
    unmatchedActionItems: []
  };

  // Fetch Todoist tasks due today or overdue
  if (client.todoist_project_id) {
    context.todoistTasks = fetchTodoistTasksDueToday(client.todoist_project_id);
    Logger.log(`Found ${context.todoistTasks.length} Todoist tasks`);
  }

  // Fetch recent emails
  context.recentEmails = fetchRecentClientEmails(client);
  Logger.log(`Found ${context.recentEmails.length} recent emails`);

  // Fetch previous meeting notes and identify unmatched action items
  if (client.google_doc_url) {
    const meetingHistory = fetchPreviousMeetingNotes(client);
    context.previousMeetingNotes = meetingHistory.notes;
    context.unmatchedActionItems = findUnmatchedActionItems(
      meetingHistory.actionItems,
      context.todoistTasks
    );
    Logger.log(`Found ${context.unmatchedActionItems.length} unmatched action items`);
  }

  return context;
}

/**
 * Fetches recent emails from/to the client.
 *
 * @param {Object} client - The client object
 * @returns {Object[]} Array of email summary objects
 */
function fetchRecentClientEmails(client) {
  const domains = parseCommaSeparatedList(client.email_domains);
  const contacts = parseCommaSeparatedList(client.contact_emails);

  if (domains.length === 0 && contacts.length === 0) {
    return [];
  }

  // Build search query
  const fromParts = [];
  const toParts = [];

  for (const domain of domains) {
    fromParts.push(`from:*@${domain}`);
    toParts.push(`to:*@${domain}`);
  }

  for (const contact of contacts) {
    fromParts.push(`from:${contact}`);
    toParts.push(`to:${contact}`);
  }

  const query = `(${fromParts.join(' OR ')} OR ${toParts.join(' OR ')}) newer_than:7d`;

  try {
    const threads = GmailApp.search(query, 0, 20);
    const emails = [];

    for (const thread of threads) {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];

      emails.push({
        subject: thread.getFirstMessageSubject(),
        sender: lastMessage.getFrom(),
        date: lastMessage.getDate(),
        snippet: lastMessage.getPlainBody().substring(0, 500)
      });
    }

    return emails;
  } catch (error) {
    Logger.log(`Failed to fetch client emails: ${error.message}`);
    return [];
  }
}

/**
 * Fetches previous meeting notes from client's Google Doc.
 *
 * @param {Object} client - The client object
 * @returns {Object} Object with notes text and extracted action items
 */
function fetchPreviousMeetingNotes(client) {
  const result = {
    notes: null,
    actionItems: []
  };

  if (!client.google_doc_url) {
    return result;
  }

  try {
    const docId = extractDocIdFromUrl(client.google_doc_url);
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const text = body.getText();

    // Find the most recent "Meeting Notes" section
    const regex = /Meeting Notes - (\d{4}-\d{2}-\d{2}|\w+ \d{1,2}, \d{4})\n([\s\S]*?)(?=Meeting Notes -|Meeting Agenda -|$)/gi;
    const matches = [...text.matchAll(regex)];

    if (matches.length > 0) {
      // Get the most recent match
      const lastMatch = matches[matches.length - 1];
      result.notes = lastMatch[2].trim();

      // Extract action items from the notes
      result.actionItems = extractActionItemsFromText(result.notes);
    }

    return result;
  } catch (error) {
    Logger.log(`Failed to fetch previous meeting notes: ${error.message}`);
    return result;
  }
}

/**
 * Extracts action items from meeting notes text.
 *
 * @param {string} text - The meeting notes text
 * @returns {string[]} Array of action item strings
 */
function extractActionItemsFromText(text) {
  const actionItems = [];

  // Look for numbered items after "Action Items" header
  const actionSection = text.match(/Action Items[\s\S]*?(?=\n\n|\n---|\n#|$)/i);
  if (actionSection) {
    const items = actionSection[0].match(/\d+\.\s+(.+)/g);
    if (items) {
      for (const item of items) {
        actionItems.push(item.replace(/^\d+\.\s+/, '').trim());
      }
    }
  }

  // Also look for bullet points
  const bulletItems = text.match(/[-*]\s+(?:TODO|Action|Follow.up):\s*(.+)/gi);
  if (bulletItems) {
    for (const item of bulletItems) {
      actionItems.push(item.replace(/^[-*]\s+(?:TODO|Action|Follow.up):\s*/i, '').trim());
    }
  }

  return actionItems;
}

/**
 * Finds action items from previous meetings that aren't in Todoist.
 *
 * @param {string[]} actionItems - Action items from meeting notes
 * @param {Object[]} todoistTasks - Current Todoist tasks
 * @returns {string[]} Action items not found in Todoist
 */
function findUnmatchedActionItems(actionItems, todoistTasks) {
  if (actionItems.length === 0) {
    return [];
  }

  const taskContents = todoistTasks.map(t => t.content.toLowerCase());

  return actionItems.filter(item => {
    const itemLower = item.toLowerCase();
    // Check if any Todoist task contains this action item (fuzzy match)
    return !taskContents.some(taskContent =>
      taskContent.includes(itemLower) ||
      itemLower.includes(taskContent) ||
      similarityScore(itemLower, taskContent) > 0.7
    );
  });
}

/**
 * Calculates a simple similarity score between two strings.
 *
 * @param {string} str1 - First string
 * @param {string} str2 - Second string
 * @returns {number} Similarity score between 0 and 1
 */
function similarityScore(str1, str2) {
  const words1 = str1.toLowerCase().split(/\s+/);
  const words2 = str2.toLowerCase().split(/\s+/);

  const commonWords = words1.filter(w => words2.includes(w));
  return commonWords.length / Math.max(words1.length, words2.length);
}

// ============================================================================
// CLAUDE API INTEGRATION
// ============================================================================

/**
 * Generates an agenda using Claude API.
 *
 * @param {CalendarEvent} event - The calendar event
 * @param {Object} client - The client object
 * @param {Object} context - The gathered context
 * @returns {string|null} The generated agenda content or null
 */
function generateAgendaWithClaude(event, client, context) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    Logger.log('Claude API key not configured');
    return null;
  }

  const prompt = buildAgendaPrompt(event, client, context);

  try {
    const url = 'https://api.anthropic.com/v1/messages';

    const payload = {
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      messages: [
        {
          role: 'user',
          content: prompt
        }
      ]
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
      Logger.log(`Claude API error: ${responseCode} - ${response.getContentText()}`);
      return null;
    }

    const result = JSON.parse(response.getContentText());

    if (result.content && result.content.length > 0) {
      return result.content[0].text;
    }

    return null;

  } catch (error) {
    Logger.log(`Failed to call Claude API: ${error.message}`);
    return null;
  }
}

/**
 * Builds the prompt for Claude agenda generation.
 *
 * @param {CalendarEvent} event - The calendar event
 * @param {Object} client - The client object
 * @param {Object} context - The gathered context
 * @returns {string} The formatted prompt
 */
function buildAgendaPrompt(event, client, context) {
  const eventDate = formatDateTime(event.getStartTime());

  let prompt = `Generate a concise meeting agenda for the following meeting. Include time allocations for each agenda item.

Meeting Details:
- Title: ${event.getTitle()}
- Client: ${client.client_name}
- Date/Time: ${eventDate}
- Duration: ${getEventDurationMinutes(event)} minutes

`;

  // Add Todoist tasks
  if (context.todoistTasks.length > 0) {
    prompt += `Outstanding Tasks (Due Today or Overdue):\n`;
    for (const task of context.todoistTasks) {
      prompt += `- ${task.content}`;
      if (task.due) {
        prompt += ` (Due: ${task.due.date})`;
      }
      prompt += `\n`;
    }
    prompt += `\n`;
  }

  // Add recent email activity
  if (context.recentEmails.length > 0) {
    prompt += `Recent Email Activity (Last 7 Days):\n`;
    for (const email of context.recentEmails.slice(0, 5)) {
      prompt += `- ${email.subject} (${formatDate(email.date)})\n`;
    }
    prompt += `\n`;
  }

  // Add previous meeting notes
  if (context.previousMeetingNotes) {
    prompt += `Previous Meeting Notes Summary:\n${context.previousMeetingNotes.substring(0, 500)}\n\n`;
  }

  // Add unmatched action items
  if (context.unmatchedActionItems.length > 0) {
    prompt += `Action Items from Last Meeting Not Yet in Task List:\n`;
    for (const item of context.unmatchedActionItems) {
      prompt += `- ${item}\n`;
    }
    prompt += `\n`;
  }

  prompt += `Please generate a structured agenda with:
1. Clear agenda items with time allocations
2. Priority items based on outstanding tasks and action items
3. Any topics suggested by recent email activity
4. Time for open discussion

Format the agenda professionally and keep it concise.`;

  return prompt;
}

/**
 * Gets the duration of an event in minutes.
 *
 * @param {CalendarEvent} event - The calendar event
 * @returns {number} Duration in minutes
 */
function getEventDurationMinutes(event) {
  const start = event.getStartTime().getTime();
  const end = event.getEndTime().getTime();
  return Math.round((end - start) / (1000 * 60));
}

// ============================================================================
// AGENDA OUTPUT
// ============================================================================

/**
 * Sends the agenda email to the user.
 *
 * @param {CalendarEvent} event - The calendar event
 * @param {Object} client - The client object
 * @param {string} agendaContent - The generated agenda
 */
function sendAgendaEmail(event, client, agendaContent) {
  const userEmail = getCurrentUserEmail();
  const subject = `Agenda: ${event.getTitle()}`;
  const eventDateTime = formatDateTime(event.getStartTime());

  let body = `<h2>Meeting Agenda</h2>`;
  body += `<p><strong>Meeting:</strong> ${event.getTitle()}</p>`;
  body += `<p><strong>Client:</strong> ${client.client_name}</p>`;
  body += `<p><strong>Date/Time:</strong> ${eventDateTime}</p>`;
  body += `<hr/>`;
  body += `<div style="white-space: pre-wrap;">${agendaContent}</div>`;
  body += `<hr/>`;
  body += `<p><em>Please review and let me know if there are any additional topics you'd like to discuss.</em></p>`;

  GmailApp.sendEmail(userEmail, subject, '', {
    htmlBody: body
  });

  Logger.log(`Sent agenda email for: ${event.getTitle()}`);
}

/**
 * Appends the agenda to the client's Google Doc.
 *
 * @param {CalendarEvent} event - The calendar event
 * @param {Object} client - The client object
 * @param {string} agendaContent - The generated agenda
 */
function appendAgendaToDoc(event, client, agendaContent) {
  if (!client.google_doc_url) {
    Logger.log('No Google Doc URL configured for client');
    return;
  }

  try {
    const docId = extractDocIdFromUrl(client.google_doc_url);
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    const formattedDate = formatDate(event.getStartTime());

    // Add section header
    body.appendParagraph(`Meeting Agenda - ${formattedDate}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Add agenda content
    body.appendParagraph(agendaContent);

    // Add separator
    body.appendParagraph('');

    doc.saveAndClose();

    Logger.log(`Appended agenda to doc for: ${client.client_name}`);
  } catch (error) {
    Logger.log(`Failed to append agenda to doc: ${error.message}`);
  }
}

// ============================================================================
// TRACKING
// ============================================================================

/**
 * Checks if an agenda has already been generated for an event.
 *
 * @param {string} eventId - The calendar event ID
 * @returns {boolean} True if agenda already generated
 */
function isAgendaGenerated(eventId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.GENERATED_AGENDAS);

  if (!sheet) {
    return false;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === eventId) {
      return true;
    }
  }

  return false;
}

/**
 * Records a generated agenda in the tracking sheet.
 *
 * @param {CalendarEvent} event - The calendar event
 * @param {Object} client - The client object
 */
function recordGeneratedAgenda(event, client) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.GENERATED_AGENDAS);

  if (!sheet) {
    Logger.log('Generated_Agendas sheet not found');
    return;
  }

  sheet.appendRow([
    event.getId(),
    event.getTitle(),
    client.client_id,
    new Date().toISOString()
  ]);

  Logger.log(`Recorded generated agenda for: ${event.getTitle()}`);
}

// ============================================================================
// CALENDAR HELPERS
// ============================================================================

/**
 * Gets today's remaining calendar events.
 *
 * @returns {CalendarEvent[]} Array of calendar events
 */
function getTodaysRemainingEvents() {
  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const endOfDay = new Date(now);
  endOfDay.setHours(23, 59, 59, 999);

  const events = calendar.getEvents(now, endOfDay);

  // Filter out all-day events and events without guests
  return events.filter(event => {
    // Skip all-day events
    if (event.isAllDayEvent()) {
      return false;
    }

    // Only include events with external guests
    const guests = event.getGuestList();
    return guests.length > 0;
  });
}

/**
 * Gets events for the upcoming week.
 *
 * @returns {CalendarEvent[]} Array of calendar events
 */
function getWeekEvents() {
  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const endOfWeek = new Date(now);
  endOfWeek.setDate(endOfWeek.getDate() + 7);

  const events = calendar.getEvents(now, endOfWeek);

  return events.filter(event => !event.isAllDayEvent());
}

// ============================================================================
// CALENDAR ATTACHMENT
// ============================================================================

/**
 * Attaches the client's running meeting notes doc to a calendar event if not already attached.
 * Requires the Advanced Calendar Service to be enabled.
 *
 * @param {CalendarEvent} event - The calendar event
 * @param {Object} client - The client object with google_doc_url
 */
function attachDocToEventIfMissing(event, client) {
  if (!client.google_doc_url) {
    return;
  }

  try {
    const docId = extractDocIdFromUrl(client.google_doc_url);
    const eventId = event.getId();
    const calendarId = 'primary';

    // Get existing attachments using Advanced Calendar Service
    const calendarEvent = Calendar.Events.get(calendarId, eventId);
    const existingAttachments = calendarEvent.attachments || [];

    // Check if doc is already attached
    const docUrl = `https://docs.google.com/document/d/${docId}`;
    const isAlreadyAttached = existingAttachments.some(att =>
      att.fileUrl && att.fileUrl.includes(docId)
    );

    if (isAlreadyAttached) {
      Logger.log(`Doc already attached to event: ${event.getTitle()}`);
      return;
    }

    // Get doc details for attachment
    const doc = DocumentApp.openById(docId);
    const docName = doc.getName();

    // Add the attachment
    const newAttachment = {
      fileUrl: docUrl,
      title: docName,
      mimeType: 'application/vnd.google-apps.document'
    };

    // Update event with new attachment
    const updatedAttachments = [...existingAttachments, newAttachment];

    Calendar.Events.patch(
      { attachments: updatedAttachments },
      calendarId,
      eventId,
      { supportsAttachments: true }
    );

    Logger.log(`Attached "${docName}" to event: ${event.getTitle()}`);

    logProcessing(
      'DOC_ATTACHED',
      client.client_id,
      `Attached meeting notes to: ${event.getTitle()}`,
      'success'
    );

  } catch (error) {
    Logger.log(`Failed to attach doc to event: ${error.message}`);
    // Don't log as error - Advanced Calendar Service may not be enabled
    // This is a nice-to-have feature
  }
}

/**
 * Manually attaches running meeting notes to all upcoming client meetings.
 * Useful for initial setup or bulk attachment.
 */
function attachDocsToAllUpcomingMeetings() {
  Logger.log('Attaching docs to all upcoming client meetings...');

  const events = getWeekEvents();
  let attachedCount = 0;

  for (const event of events) {
    const client = identifyClientFromCalendarEvent(event);

    if (client && client.google_doc_url) {
      attachDocToEventIfMissing(event, client);
      attachedCount++;
    }
  }

  Logger.log(`Processed ${attachedCount} events for doc attachment`);
}
