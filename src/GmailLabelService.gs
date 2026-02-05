/**
 * GmailLabelService.gs - Gmail label helpers.
 */

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
