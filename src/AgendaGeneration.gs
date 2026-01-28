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
  logProcessing('AGENDA_GEN', null, 'Starting agenda generation scan', 'info');

  // Get today's remaining calendar events
  const events = getTodaysRemainingEvents();

  if (events.length === 0) {
    logProcessing('AGENDA_GEN', null, 'No remaining events for today', 'info');
    return;
  }

  logProcessing('AGENDA_GEN', null, `Found ${events.length} remaining events to process`, 'info');

  for (const event of events) {
    processEventForAgenda(event);
  }

  logProcessing('AGENDA_GEN', null, 'Agenda generation scan completed', 'success');
}

/**
 * Processes a single calendar event for potential agenda generation.
 *
 * @param {CalendarEvent} event - The calendar event
 */
function processEventForAgenda(event) {
  const eventId = event.getId();
  const eventTitle = event.getTitle();

  logProcessing('AGENDA_GEN', null, `Processing event: ${eventTitle}`, 'info');

  // Identify client from event guests
  const client = identifyClientFromCalendarEvent(event);

  if (!client) {
    // Log to unmatched and skip
    const guestEmails = event.getGuestList(true).map(g => g.getEmail());
    logUnmatched(
      'meeting',
      `Calendar Event: ${eventTitle} at ${event.getStartTime()}`,
      guestEmails
    );
    logProcessing('AGENDA_GEN', null, `No client match for event: ${eventTitle}`, 'info');
    return;
  }

  logProcessing('AGENDA_GEN', client.client_name, `Matched event "${eventTitle}" to client`, 'info');

  // Ensure running meeting notes doc is attached to the event
  if (client.google_doc_url) {
    try {
      attachDocToEventIfMissing(event, client);
    } catch (error) {
      logProcessing('AGENDA_GEN', client.client_name, `Failed to attach doc to event: ${error.message}`, 'warning');
    }
  }

  // Check if agenda already generated
  if (isAgendaGenerated(eventId)) {
    logProcessing('AGENDA_GEN', client.client_name, `Agenda already generated for: ${eventTitle}`, 'info');
    return;
  }

  // Generate agenda for this event
  try {
    generateAgendaForEvent(event, client);
  } catch (error) {
    const errorMsg = `Failed to generate agenda for ${eventTitle}: ${error.message}\nStack: ${error.stack}`;
    logProcessing('AGENDA_ERROR', client.client_name, errorMsg, 'error');
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

  // DIAGNOSTIC: Start trace
  const traceId = startAgendaTrace(event, client);

  // Gather context
  if (traceId) logAgendaStep(traceId, event, client, 4, 'GATHER_CONTEXT_START', 'started', 'Starting context gathering');
  const context = gatherAgendaContext(event, client, traceId);
  if (traceId) logAgendaStep(traceId, event, client, 4, 'GATHER_CONTEXT_START', 'success', 'Context gathering complete');

  // Generate agenda with Claude
  const agendaContent = generateAgendaWithClaude(event, client, context, traceId);

  if (!agendaContent) {
    Logger.log('Failed to generate agenda content from Claude');
    if (traceId) logAgendaStep(traceId, event, client, 8, 'AGENDA_FAILED', 'failed', 'Failed to generate agenda content');
    return;
  }

  // Send agenda email (will apply label from Client Registry)
  if (traceId) logAgendaStep(traceId, event, client, 9, 'SEND_EMAIL', 'started', 'Sending agenda email');
  sendAgendaEmail(event, client, agendaContent, traceId);
  if (traceId) logAgendaStep(traceId, event, client, 9, 'SEND_EMAIL', 'success', 'Email sent successfully');

  // Append to client's Google Doc
  if (traceId) logAgendaStep(traceId, event, client, 10, 'APPEND_TO_DOC', 'started', 'Appending agenda to Google Doc');
  appendAgendaToDoc(event, client, agendaContent, traceId);
  if (traceId) logAgendaStep(traceId, event, client, 10, 'APPEND_TO_DOC', 'success', 'Appended to doc successfully');

  // Record the generation
  recordGeneratedAgenda(event, client);

  if (traceId) logAgendaStep(traceId, event, client, 11, 'COMPLETE', 'success', 'Agenda generation completed successfully');

  logProcessing(
    'AGENDA_GENERATED',
    client.client_name,
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
  const contacts = parseCommaSeparatedList(client.contact_emails);

  if (contacts.length === 0) {
    return [];
  }

  // Build search query using contact emails
  const fromParts = [];
  const toParts = [];

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
 * Fetches ONLY the most recent meeting notes and agenda from client's Google Doc.
 * This minimizes token usage by not pulling the entire document history.
 *
 * @param {Object} client - The client object
 * @returns {Object} Object with lastNotes, lastAgenda, and extracted actionItems
 */
function fetchPreviousMeetingNotes(client) {
  const result = {
    notes: null,
    agenda: null,
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

    // Find the LAST meeting notes section using the structured delimiters
    // Pattern: ═══...═══ MEETING NOTES - [date] ───...─── [content] ───...─── END OF MEETING NOTES
    const notesRegex = /═{20,}\s*\n\s*MEETING NOTES - ([^\n]+)\s*\n\s*─{20,}\s*\n([\s\S]*?)\n\s*─{20,}\s*\n\s*END OF MEETING NOTES/gi;
    const notesMatches = [...text.matchAll(notesRegex)];

    if (notesMatches.length > 0) {
      // Get the LAST (most recent) meeting notes only
      const lastNotesMatch = notesMatches[notesMatches.length - 1];
      result.notes = lastNotesMatch[2].trim();

      // Extract action items from the notes
      result.actionItems = extractActionItemsFromText(result.notes);

      logProcessing('AGENDA_CONTEXT', client.client_name, `Found last meeting notes from ${lastNotesMatch[1]}`, 'info');
    } else {
      // Fallback: try old format without delimiters
      const oldRegex = /Meeting Notes - (\d{4}-\d{2}-\d{2}|\w+ \d{1,2}, \d{4})\n([\s\S]*?)(?=Meeting Notes -|Meeting Agenda -|═{20,}|$)/gi;
      const oldMatches = [...text.matchAll(oldRegex)];

      if (oldMatches.length > 0) {
        const lastMatch = oldMatches[oldMatches.length - 1];
        result.notes = lastMatch[2].trim();
        result.actionItems = extractActionItemsFromText(result.notes);

        logProcessing('AGENDA_CONTEXT', client.client_name, `Found last meeting notes (old format) from ${lastMatch[1]}`, 'info');
      }
    }

    // Find the LAST agenda section
    // Pattern: ═══...═══ MEETING AGENDA - [date] ───...─── [content] ───...─── END OF MEETING AGENDA
    const agendaRegex = /═{20,}\s*\n\s*MEETING AGENDA - ([^\n]+)\s*\n\s*─{20,}\s*\n([\s\S]*?)\n\s*─{20,}\s*\n\s*END OF MEETING AGENDA/gi;
    const agendaMatches = [...text.matchAll(agendaRegex)];

    if (agendaMatches.length > 0) {
      // Get the LAST (most recent) agenda only
      const lastAgendaMatch = agendaMatches[agendaMatches.length - 1];
      result.agenda = lastAgendaMatch[2].trim();

      logProcessing('AGENDA_CONTEXT', client.client_name, `Found last agenda from ${lastAgendaMatch[1]}`, 'info');
    } else {
      // Fallback: try old format
      const oldAgendaRegex = /Meeting Agenda - (\d{4}-\d{2}-\d{2}|\w+ \d{1,2}, \d{4})\n([\s\S]*?)(?=Meeting Notes -|Meeting Agenda -|═{20,}|$)/gi;
      const oldAgendaMatches = [...text.matchAll(oldAgendaRegex)];

      if (oldAgendaMatches.length > 0) {
        const lastMatch = oldAgendaMatches[oldAgendaMatches.length - 1];
        result.agenda = lastMatch[2].trim();

        logProcessing('AGENDA_CONTEXT', client.client_name, `Found last agenda (old format) from ${lastMatch[1]}`, 'info');
      }
    }

    return result;
  } catch (error) {
    logProcessing('AGENDA_CONTEXT', client.client_name, `Failed to fetch previous meeting notes: ${error.message}`, 'error');
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
 * @param {string} traceId - Optional trace ID for diagnostic logging
 * @returns {string|null} The generated agenda content or null
 */
function generateAgendaWithClaude(event, client, context, traceId) {
  const startTime = new Date().getTime();

  // DIAGNOSTIC: Log step - Check API key
  if (traceId) logAgendaStep(traceId, event, client, 5, 'CHECK_API_KEY', 'started', 'Verifying Claude API key');
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    if (traceId) logAgendaStep(traceId, event, client, 5, 'CHECK_API_KEY', 'failed', 'Claude API key not configured');
    logProcessing('AGENDA_ERROR', client.client_name, 'Claude API key not configured - cannot generate agenda', 'error');
    return null;
  }

  if (traceId) logAgendaStep(traceId, event, client, 5, 'CHECK_API_KEY', 'success', 'API key found');

  // DIAGNOSTIC: Log step - Build prompt
  if (traceId) logAgendaStep(traceId, event, client, 6, 'BUILD_PROMPT', 'started', 'Building Claude prompt with context');

  const prompt = buildAgendaPrompt(event, client, context);

  if (traceId) logAgendaStep(traceId, event, client, 6, 'BUILD_PROMPT', 'success', `Prompt built (${prompt.length} chars)`);

  try {
    const url = 'https://api.anthropic.com/v1/messages';

    // Use model preference from Prompts sheet (allows user to choose haiku vs sonnet)
    const model = getModelForPrompt('AGENDA_CLAUDE_PROMPT');
    logProcessing('AGENDA_GEN', client.client_name, `Using Claude model: ${model}`, 'info');

    const payload = {
      model: model,
      max_tokens: 1000,
      messages: [
        {
          role: 'user',
          content: prompt
        }
      ]
    };

    const headers = {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json'
    };

    const options = {
      method: 'POST',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    // DIAGNOSTIC: Log API request
    const requestId = logAPIRequest(
      'Claude API',
      url,
      'POST',
      headers,
      payload,
      {
        clientId: client.client_name,
        eventId: event.getId(),
        flow: 'Agenda Generation'
      }
    );

    if (traceId) logAgendaStep(traceId, event, client, 7, 'CALL_CLAUDE_API', 'started', `Calling Claude API (request ID: ${requestId || 'N/A'})`);

    // Make API call and measure duration
    const apiStartTime = new Date().getTime();
    const response = UrlFetchApp.fetch(url, options);
    const apiDuration = new Date().getTime() - apiStartTime;

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    // DIAGNOSTIC: Log API response
    let parseSuccess = false;
    let result = null;
    let extractedData = null;
    let apiError = null;

    try {
      if (responseCode === 200) {
        result = JSON.parse(responseText);
        parseSuccess = true;
        extractedData = {
          model: result.model || model,
          stop_reason: result.stop_reason,
          content_length: result.content && result.content[0] ? result.content[0].text.length : 0
        };
      } else {
        apiError = `HTTP ${responseCode}`;
      }
    } catch (e) {
      apiError = `Parse error: ${e.message}`;
    }

    logAPIResponse(
      requestId,
      'Claude API',
      responseCode,
      { 'content-type': response.getHeaders()['Content-Type'] || '' },
      responseText,
      parseSuccess,
      extractedData,
      apiError,
      apiDuration
    );

    if (responseCode !== 200) {
      const errorDetail = `Claude API returned ${responseCode}: ${responseText}`;
      if (traceId) logAgendaStep(traceId, event, client, 7, 'CALL_CLAUDE_API', 'failed', `API error: ${responseCode}`, null, apiDuration);
      logProcessing('AGENDA_ERROR', client.client_name, errorDetail, 'error');
      return null;
    }

    if (!result) {
      result = JSON.parse(responseText);
    }

    if (result.content && result.content.length > 0) {
      if (traceId) logAgendaStep(traceId, event, client, 7, 'CALL_CLAUDE_API', 'success', `API call successful (${apiDuration}ms)`, `Response: ${extractedData ? extractedData.content_length : 0} chars`, apiDuration);
      logProcessing('AGENDA_GEN', client.client_name, 'Successfully generated agenda with Claude', 'success');
      let content = result.content[0].text;

      if (!content || content.trim().length === 0) {
        logProcessing('AGENDA_ERROR', client.client_name, 'Claude returned empty text content', 'error');
        return null;
      }

      // Strip markdown code fences if present (Claude sometimes wraps HTML in ```html ... ```)
      content = content.replace(/^```html\s*/i, '').replace(/\s*```$/, '');
      content = content.replace(/^```\s*/i, '').replace(/\s*```$/, '');
      content = content.trim();

      // Extract body content from full HTML document if present
      // Claude sometimes returns full HTML with <!DOCTYPE>, <head>, <style>, etc.
      const bodyMatch = content.match(/<body[^>]*>([\s\S]*)<\/body>/i);
      if (bodyMatch && bodyMatch[1].trim().length > 0) {
        content = bodyMatch[1].trim();
        logProcessing('AGENDA_GEN', client.client_name, 'Extracted body content from full HTML document', 'info');
      }

      // If content is still empty or too short after extraction, log error
      if (!content || content.trim().length < 10) {
        logProcessing('AGENDA_ERROR', client.client_name, `Generated agenda is too short or empty: "${content}"`, 'error');
        return null;
      }

      logProcessing('AGENDA_GEN', client.client_name, `Generated agenda content (${content.length} chars)`, 'success');
      return content;
    }

    logProcessing('AGENDA_ERROR', client.client_name, 'Claude returned empty content array', 'error');
    return null;

  } catch (error) {
    const errorDetail = `Failed to call Claude API: ${error.message}\nStack: ${error.stack}`;
    logProcessing('AGENDA_ERROR', client.client_name, errorDetail, 'error');
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

  // Build context sections
  let todoistSection = '';
  if (context.todoistTasks.length > 0) {
    todoistSection = `Outstanding Tasks (Due Today or Overdue):\n`;
    for (const task of context.todoistTasks) {
      todoistSection += `- ${task.content}`;
      if (task.due) {
        todoistSection += ` (Due: ${task.due.date})`;
      }
      todoistSection += `\n`;
    }
  }

  let emailsSection = '';
  if (context.recentEmails.length > 0) {
    emailsSection = `Recent Email Activity (Last 7 Days):\n`;
    for (const email of context.recentEmails.slice(0, 5)) {
      emailsSection += `- ${email.subject} (${formatDate(email.date)})\n`;
    }
  }

  let notesSection = '';
  if (context.previousMeetingNotes) {
    notesSection = `Previous Meeting Notes Summary:\n${context.previousMeetingNotes.substring(0, 500)}`;
  }

  let actionItemsSection = '';
  if (context.unmatchedActionItems.length > 0) {
    actionItemsSection = `Action Items from Last Meeting Not Yet in Task List:\n`;
    for (const item of context.unmatchedActionItems) {
      actionItemsSection += `- ${item}\n`;
    }
  }

  // Get prompt template from sheet
  const template = getPrompt('AGENDA_CLAUDE_PROMPT');

  // Apply variables to template
  return applyTemplate(template, {
    event_title: event.getTitle(),
    client_name: client.client_name,
    date_time: eventDate,
    duration: getEventDurationMinutes(event).toString(),
    todoist_section: todoistSection,
    emails_section: emailsSection,
    notes_section: notesSection,
    action_items_section: actionItemsSection
  });
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
  const eventDateTime = formatDateTime(event.getStartTime());
  const eventDate = formatDateShort(event.getStartTime());

  // Get customizable subject template from settings
  const props = PropertiesService.getScriptProperties();
  const subjectTemplate = props.getProperty('AGENDA_SUBJECT_TEMPLATE')
    || 'Agenda: {client_name} - {meeting_title} ({date})';
  const subject = subjectTemplate
    .replace('{client_name}', client.client_name)
    .replace('{meeting_title}', event.getTitle())
    .replace('{date}', eventDate);

  // Get email template from sheet
  const template = getPrompt('AGENDA_EMAIL_TEMPLATE');

  let body = applyTemplate(template, {
    event_title: event.getTitle(),
    client_name: client.client_name,
    date_time: eventDateTime,
    agenda_content: agendaContent
  });

  // Ensure proper UTF-8 encoding by adding meta tag if not present
  if (!body.match(/<meta[^>]+charset/i)) {
    // If body doesn't have HTML structure, wrap it
    if (!body.match(/<html/i)) {
      body = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
${body}
</body>
</html>`;
    } else {
      // Insert meta tag in existing HTML
      body = body.replace(/<head[^>]*>/i, match => {
        return match + '\n<meta charset="UTF-8">\n<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">';
      });
    }
  }

  GmailApp.sendEmail(userEmail, subject, '', {
    htmlBody: body,
    charset: 'UTF-8'
  });

  Logger.log(`Sent agenda email for: ${event.getTitle()}`);

  // Apply label from Client Registry immediately after sending
  try {
    const labelName = client.meeting_agendas_label;
    if (labelName) {
      // Search for the just-sent email
      const query = `from:me to:me subject:"${subject}" newer_than:5m`;
      Utilities.sleep(2000); // Wait 2 seconds for email to appear in Gmail
      const threads = GmailApp.search(query, 0, 1);

      if (threads.length > 0) {
        const label = createLabelIfNotExists(labelName);
        threads[0].addLabel(label);
        logProcessing('AGENDA_EMAIL', client.client_name, `Applied label: ${labelName}`, 'success');
      } else {
        logProcessing('AGENDA_EMAIL', client.client_name, 'Could not find sent email to label', 'warning');
      }
    }
  } catch (error) {
    logProcessing('AGENDA_EMAIL', client.client_name, `Failed to apply label: ${error.message}`, 'warning');
  }
}

/**
 * Converts HTML content to plain text suitable for Google Docs.
 * Strips HTML tags and converts common entities to readable text.
 *
 * @param {string} html - The HTML content
 * @returns {string} Plain text version
 */
function htmlToPlainText(html) {
  if (!html) return '';

  let text = html;

  // First, aggressively remove everything in <head>, <style>, and <script> tags
  // Use global flag and handle multiline content
  text = text.replace(/<head[\s\S]*?<\/head>/gi, '');
  text = text.replace(/<style[\s\S]*?<\/style>/gi, '');
  text = text.replace(/<script[\s\S]*?<\/script>/gi, '');

  // Strip DOCTYPE and root HTML tags
  text = text.replace(/<!DOCTYPE[^>]*>/gi, '');
  text = text.replace(/<html[^>]*>/gi, '');
  text = text.replace(/<\/html>/gi, '');
  text = text.replace(/<body[^>]*>/gi, '');
  text = text.replace(/<\/body>/gi, '');

  // Extract just body content if body tags still exist (redundant safety check)
  const bodyMatch = text.match(/<body[^>]*>([\s\S]*)<\/body>/i);
  if (bodyMatch) {
    text = bodyMatch[1];
  }

  // Convert block elements to newlines BEFORE stripping tags
  text = text.replace(/<\/p>/gi, '\n\n');
  text = text.replace(/<p[^>]*>/gi, '');
  text = text.replace(/<br\s*\/?>/gi, '\n');
  text = text.replace(/<\/div>/gi, '\n');
  text = text.replace(/<div[^>]*>/gi, '');
  text = text.replace(/<\/h[1-6]>/gi, '\n\n');
  text = text.replace(/<h[1-6][^>]*>/gi, '');

  // Convert list items
  text = text.replace(/<\/li>/gi, '\n');
  text = text.replace(/<li[^>]*>/gi, '  • ');
  text = text.replace(/<\/ul>/gi, '\n');
  text = text.replace(/<ul[^>]*>/gi, '');
  text = text.replace(/<\/ol>/gi, '\n');
  text = text.replace(/<ol[^>]*>/gi, '');

  // Strip ALL remaining HTML tags - including partial/malformed tags
  // First pass: Remove complete tags
  text = text.replace(/<[^>]+>/g, '');

  // Second pass: Remove any remaining partial tags (handles malformed HTML like "</h" or "<div")
  // This catches incomplete opening tags
  text = text.replace(/<[^>]*$/gm, '');
  // This catches incomplete closing tags or any remaining angle brackets with text
  text = text.replace(/<\/?[a-zA-Z][^>]*/g, '');
  // Clean up any stray angle brackets
  text = text.replace(/[<>]/g, '');

  // Decode ALL HTML entities - including numeric ones
  // First, decode numeric entities (&#123; or &#x1F4; format)
  text = text.replace(/&#(\d+);/g, (match, dec) => {
    try {
      return String.fromCharCode(parseInt(dec, 10));
    } catch (e) {
      return match;
    }
  });
  text = text.replace(/&#x([0-9A-Fa-f]+);/g, (match, hex) => {
    try {
      return String.fromCharCode(parseInt(hex, 16));
    } catch (e) {
      return match;
    }
  });

  // Decode common named HTML entities
  text = text.replace(/&nbsp;/gi, ' ');
  text = text.replace(/&amp;/gi, '&');
  text = text.replace(/&lt;/gi, '<');
  text = text.replace(/&gt;/gi, '>');
  text = text.replace(/&quot;/gi, '"');
  text = text.replace(/&#39;/gi, "'");
  text = text.replace(/&apos;/gi, "'");
  text = text.replace(/&mdash;/gi, '—');
  text = text.replace(/&ndash;/gi, '–');
  text = text.replace(/&bull;/gi, '•');
  text = text.replace(/&hellip;/gi, '…');
  text = text.replace(/&ldquo;/gi, '"');
  text = text.replace(/&rdquo;/gi, '"');
  text = text.replace(/&lsquo;/gi, '\u2018');
  text = text.replace(/&rsquo;/gi, '\u2019');
  text = text.replace(/&euro;/gi, '€');
  text = text.replace(/&pound;/gi, '£');
  text = text.replace(/&yen;/gi, '¥');
  text = text.replace(/&cent;/gi, '¢');
  text = text.replace(/&copy;/gi, '©');
  text = text.replace(/&reg;/gi, '®');
  text = text.replace(/&trade;/gi, '™');

  // Clean up excessive whitespace
  text = text.replace(/\n{3,}/g, '\n\n'); // Max 2 consecutive newlines
  text = text.replace(/[ \t]+/g, ' '); // Multiple spaces to single space
  text = text.replace(/^\s+/gm, ''); // Remove leading whitespace from each line
  text = text.trim();

  return text;
}

/**
 * Appends the agenda to the client's Google Doc with structured delimiters.
 * Uses the same format as meeting notes for consistent parsing.
 *
 * @param {CalendarEvent} event - The calendar event
 * @param {Object} client - The client object
 * @param {string} agendaContent - The generated agenda (HTML format)
 */
function appendAgendaToDoc(event, client, agendaContent, traceId) {
  if (!client.google_doc_url) {
    logProcessing('AGENDA_DOC', client.client_name, 'No Google Doc URL configured', 'warning');
    return;
  }

  try {
    const docId = extractDocIdFromUrl(client.google_doc_url);

    // DIAGNOSTIC: Get doc length before
    let beforeLength = 0;
    try {
      const doc = DocumentApp.openById(docId);
      beforeLength = doc.getBody().getText().length;
    } catch (e) {
      // If we can't get before length, just continue
    }

    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    const formattedDate = formatDate(event.getStartTime());

    // Convert HTML to plain text for the doc
    const plainTextContent = htmlToPlainText(agendaContent);

    // Add blank line before for separation
    body.appendParagraph('');

    // Add start delimiter (parseable marker)
    body.appendParagraph('═══════════════════════════════════════════════════════════');

    // Add section header with date
    body.appendParagraph(`MEETING AGENDA - ${formattedDate}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Add content delimiter
    body.appendParagraph('───────────────────────────────────────────────────────────');

    // Add agenda content as plain text
    body.appendParagraph(plainTextContent);

    // Add end delimiter
    body.appendParagraph('───────────────────────────────────────────────────────────');
    body.appendParagraph(`END OF MEETING AGENDA - ${formattedDate}`);
    body.appendParagraph('═══════════════════════════════════════════════════════════');

    // Add blank line after for separation
    body.appendParagraph('');

    doc.saveAndClose();

    // DIAGNOSTIC: Get doc length after and verify
    let afterLength = 0;
    let verificationStatus = 'skipped';
    try {
      const verifyDoc = DocumentApp.openById(docId);
      afterLength = verifyDoc.getBody().getText().length;
      verificationStatus = afterLength > beforeLength ? 'verified' : 'failed';
    } catch (e) {
      verificationStatus = 'failed';
    }

    // DIAGNOSTIC: Log doc append
    logDocAppend(
      client,
      docId,
      client.google_doc_url,
      'Agenda',
      plainTextContent,
      true,
      null,
      beforeLength,
      afterLength,
      verificationStatus
    );

    logProcessing('AGENDA_DOC', client.client_name, `Appended agenda for: ${event.getTitle()}`, 'success');
  } catch (error) {
    // DIAGNOSTIC: Log failed append
    logDocAppend(
      client,
      extractDocIdFromUrl(client.google_doc_url || ''),
      client.google_doc_url || '',
      'Agenda',
      agendaContent,
      false,
      error.message,
      0,
      0,
      'failed'
    );

    logProcessing('AGENDA_DOC', client.client_name, `Failed to append agenda: ${error.message}`, 'error');
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
    client.client_name,
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

  // Filter out all-day events only
  return events.filter(event => {
    return !event.isAllDayEvent();
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
      client.client_name,
      `Attached meeting notes to: ${event.getTitle()}`,
      'success'
    );

  } catch (error) {
    Logger.log(`Failed to attach doc to event: ${error.message}`);
    // Don't log as error - Advanced Calendar Service may not be enabled
    // This is a nice-to-have feature
  }
}
