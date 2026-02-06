/**
 * SentSummaryService.gs - Processing sent meeting summaries.
 *
 * Handles monitoring of sent summaries and extraction of action items.
 */

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
