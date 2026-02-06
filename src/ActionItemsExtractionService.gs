/**
 * ActionItemsExtractionService.gs - Extract action items from summaries.
 */

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

  const prompt = buildActionItemsExtractionPrompt(emailBody, collaborators);

  try {
    const url = 'https://api.anthropic.com/v1/messages';

    // Use dynamic model (prefer sonnet for better extraction quality)
    const models = fetchAvailableModelsFromAPI(false);
    const sonnet = models.find(m => m.id.includes('sonnet'));
    const model = sonnet ? sonnet.id : models[0]?.id || FALLBACK_MODELS[0].id;

    // Get max_tokens from settings (default 2000)
    const maxTokens = parseInt(PropertiesService.getScriptProperties().getProperty('CLAUDE_SUMMARY_MAX_TOKENS') || '2000');

    const payload = {
      model: model,
      max_tokens: maxTokens,
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

    // Parse response with explicit UTF-8 handling to preserve emojis
    const responseBytes = response.getContent();
    const responseText = Utilities.newBlob(responseBytes).getDataAsString('UTF-8');
    const result = JSON.parse(responseText);

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
