/**
 * PromptManager.gs - Prompt Template Management
 *
 * This module manages customizable prompts and email templates:
 * - Stores prompts in a hidden "Prompts" sheet
 * - Provides UI for editing prompts
 * - Supports variable substitution in templates
 */

// ============================================================================
// PROMPT KEYS AND DEFAULTS
// ============================================================================

/**
 * Default prompts used by the system.
 * These are stored in the hidden Prompts sheet and can be customized.
 */
const DEFAULT_PROMPTS = {
  // Claude prompt for meeting agenda generation
  AGENDA_CLAUDE_PROMPT: `Generate a concise meeting agenda for the following meeting. Include time allocations for each agenda item.

Meeting Details:
- Title: {event_title}
- Client: {client_name}
- Date/Time: {date_time}
- Duration: {duration} minutes

{todoist_section}
{emails_section}
{notes_section}
{action_items_section}

Please generate a structured agenda with:
1. Clear agenda items with time allocations
2. Priority items based on outstanding tasks and action items
3. Any topics suggested by recent email activity
4. Time for open discussion

Format the agenda professionally and keep it concise.`,

  // Email template for meeting agenda
  AGENDA_EMAIL_TEMPLATE: `<h2>Meeting Agenda</h2>
<p><strong>Meeting:</strong> {event_title}</p>
<p><strong>Client:</strong> {client_name}</p>
<p><strong>Date/Time:</strong> {date_time}</p>
<hr/>
<div style="white-space: pre-wrap;">{agenda_content}</div>`,

  // Daily outlook email header
  DAILY_OUTLOOK_HEADER: `<h1>Daily Outlook - {date}</h1>`,

  // Daily outlook alerts section template
  DAILY_OUTLOOK_ALERTS: `<div style="background-color: #fff3cd; padding: 15px; margin-bottom: 20px; border-radius: 5px;">
<h2 style="margin-top: 0; color: #856404;">Alerts</h2>
{alerts_content}
</div>`,

  // Weekly outlook email header
  WEEKLY_OUTLOOK_HEADER: `<h1>Weekly Outlook - Week of {date}</h1>`,

  // Weekly outlook summary section
  WEEKLY_OUTLOOK_SUMMARY: `<div style="background-color: #e9ecef; padding: 15px; margin-bottom: 20px; border-radius: 5px;">
<h2 style="margin-top: 0;">Week Summary</h2>
<p><strong>Total Meetings:</strong> {meeting_count}</p>
<p><strong>Total Tasks:</strong> {task_count}</p>
<p><strong>Clients with Activity:</strong> {client_count}</p>
</div>`
};

/**
 * Prompt metadata for the editor UI
 */
const PROMPT_METADATA = {
  AGENDA_CLAUDE_PROMPT: {
    label: 'Meeting Agenda AI Prompt',
    description: 'The prompt sent to Claude to generate meeting agendas.',
    variables: ['{event_title}', '{client_name}', '{date_time}', '{duration}', '{todoist_section}', '{emails_section}', '{notes_section}', '{action_items_section}']
  },
  AGENDA_EMAIL_TEMPLATE: {
    label: 'Agenda Email Template',
    description: 'HTML template for agenda emails sent to you.',
    variables: ['{event_title}', '{client_name}', '{date_time}', '{agenda_content}']
  },
  DAILY_OUTLOOK_HEADER: {
    label: 'Daily Outlook Header',
    description: 'Header for daily outlook emails.',
    variables: ['{date}']
  },
  DAILY_OUTLOOK_ALERTS: {
    label: 'Daily Outlook Alerts Section',
    description: 'Template for the alerts section in daily outlook.',
    variables: ['{alerts_content}']
  },
  WEEKLY_OUTLOOK_HEADER: {
    label: 'Weekly Outlook Header',
    description: 'Header for weekly outlook emails.',
    variables: ['{date}']
  },
  WEEKLY_OUTLOOK_SUMMARY: {
    label: 'Weekly Outlook Summary',
    description: 'Summary section for weekly outlook.',
    variables: ['{meeting_count}', '{task_count}', '{client_count}']
  }
};

// ============================================================================
// PROMPT SHEET MANAGEMENT
// ============================================================================

/**
 * Creates the hidden Prompts sheet if it doesn't exist.
 * Called during setup.
 *
 * @param {Spreadsheet} ss - The spreadsheet object
 */
function createPromptsSheet(ss) {
  const sheetName = 'Prompts';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);

    // Set up headers
    sheet.getRange(1, 1, 1, 2).setValues([['prompt_key', 'prompt_value']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Add default prompts
    const promptKeys = Object.keys(DEFAULT_PROMPTS);
    const rows = promptKeys.map(key => [key, DEFAULT_PROMPTS[key]]);

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }

    // Set column widths
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 800);

    // Hide the sheet
    sheet.hideSheet();

    Logger.log('Created hidden Prompts sheet with defaults');
  }
}

/**
 * Gets a prompt value from the Prompts sheet.
 *
 * @param {string} promptKey - The prompt key
 * @returns {string} The prompt value or default if not found
 */
function getPrompt(promptKey) {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      return DEFAULT_PROMPTS[promptKey] || '';
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('Prompts');

    if (!sheet) {
      return DEFAULT_PROMPTS[promptKey] || '';
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === promptKey) {
        return data[i][1] || DEFAULT_PROMPTS[promptKey] || '';
      }
    }

    return DEFAULT_PROMPTS[promptKey] || '';

  } catch (e) {
    Logger.log(`Error getting prompt ${promptKey}: ${e.message}`);
    return DEFAULT_PROMPTS[promptKey] || '';
  }
}

/**
 * Sets a prompt value in the Prompts sheet.
 *
 * @param {string} promptKey - The prompt key
 * @param {string} promptValue - The new prompt value
 */
function setPrompt(promptKey, promptValue) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID not set');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName('Prompts');

  if (!sheet) {
    createPromptsSheet(ss);
    sheet = ss.getSheetByName('Prompts');
  }

  const data = sheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === promptKey) {
      sheet.getRange(i + 1, 2).setValue(promptValue);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([promptKey, promptValue]);
  }

  Logger.log(`Updated prompt: ${promptKey}`);
}

/**
 * Gets all prompts for the editor.
 *
 * @returns {Object[]} Array of prompt objects with key, value, label, description, variables
 */
function getAllPromptsForEditor() {
  const prompts = [];

  for (const key of Object.keys(DEFAULT_PROMPTS)) {
    const value = getPrompt(key);
    const metadata = PROMPT_METADATA[key] || {};

    prompts.push({
      key: key,
      value: value,
      label: metadata.label || key,
      description: metadata.description || '',
      variables: metadata.variables || [],
      isDefault: value === DEFAULT_PROMPTS[key]
    });
  }

  return prompts;
}

/**
 * Saves multiple prompts from the editor.
 *
 * @param {Object[]} prompts - Array of {key, value} objects
 * @returns {Object} Result with success status
 */
function savePromptsFromEditor(prompts) {
  try {
    for (const prompt of prompts) {
      if (prompt.key && prompt.value !== undefined) {
        setPrompt(prompt.key, prompt.value);
      }
    }

    return { success: true, message: `Saved ${prompts.length} prompts` };
  } catch (e) {
    Logger.log(`Error saving prompts: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Resets a prompt to its default value.
 *
 * @param {string} promptKey - The prompt key to reset
 * @returns {string} The default value
 */
function resetPromptToDefault(promptKey) {
  const defaultValue = DEFAULT_PROMPTS[promptKey];
  if (defaultValue) {
    setPrompt(promptKey, defaultValue);
  }
  return defaultValue || '';
}

/**
 * Resets all prompts to default values.
 */
function resetAllPromptsToDefault() {
  for (const key of Object.keys(DEFAULT_PROMPTS)) {
    setPrompt(key, DEFAULT_PROMPTS[key]);
  }
  Logger.log('Reset all prompts to defaults');
}

// ============================================================================
// PROMPT EDITOR UI
// ============================================================================

/**
 * Shows the prompts editor dialog.
 */
function showPromptsEditor() {
  const html = HtmlService.createHtmlOutputFromFile('PromptsEditor')
    .setWidth(700)
    .setHeight(600)
    .setTitle('Adjust Prompts');

  SpreadsheetApp.getUi().showModalDialog(html, 'Adjust Prompts');
}

// ============================================================================
// TEMPLATE HELPERS
// ============================================================================

/**
 * Replaces variables in a template string.
 *
 * @param {string} template - The template string with {variable} placeholders
 * @param {Object} variables - Object with variable names as keys
 * @returns {string} The template with variables replaced
 */
function applyTemplate(template, variables) {
  if (!template) return '';

  let result = template;

  for (const [key, value] of Object.entries(variables)) {
    const placeholder = `{${key}}`;
    result = result.split(placeholder).join(value || '');
  }

  return result;
}

// ============================================================================
// JSON FORMATS DOCUMENTATION
// ============================================================================

/**
 * API JSON format examples for reference.
 * Stored in hidden JSON_Formats sheet for documentation.
 */
const API_JSON_FORMATS = {
  TODOIST_CREATE_TASK: {
    service: 'Todoist',
    endpoint: 'POST https://api.todoist.com/rest/v2/tasks',
    auth: 'Bearer {TODOIST_API_TOKEN}',
    request: {
      content: '[ClientName] Task description',
      project_id: '1234567890',
      due_string: '2025-01-20'
    },
    response: {
      id: '8765432109',
      content: '[ClientName] Task description',
      project_id: '1234567890',
      due: {
        date: '2025-01-20',
        string: 'Jan 20'
      }
    }
  },

  TODOIST_GET_PROJECTS: {
    service: 'Todoist',
    endpoint: 'GET https://api.todoist.com/rest/v2/projects',
    auth: 'Bearer {TODOIST_API_TOKEN}',
    request: null,
    response: [
      {
        id: '1234567890',
        name: 'Client Project Name',
        is_inbox_project: false
      }
    ]
  },

  TODOIST_CREATE_PROJECT: {
    service: 'Todoist',
    endpoint: 'POST https://api.todoist.com/rest/v2/projects',
    auth: 'Bearer {TODOIST_API_TOKEN}',
    request: {
      name: 'New Project Name'
    },
    response: {
      id: '1234567890',
      name: 'New Project Name'
    }
  },

  CLAUDE_MESSAGES: {
    service: 'Claude (Anthropic)',
    endpoint: 'POST https://api.anthropic.com/v1/messages',
    auth: 'x-api-key: {CLAUDE_API_KEY}, anthropic-version: 2023-06-01',
    request: {
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      messages: [
        {
          role: 'user',
          content: 'Your prompt here'
        }
      ]
    },
    response: {
      id: 'msg_abc123',
      type: 'message',
      role: 'assistant',
      content: [
        {
          type: 'text',
          text: 'The generated response'
        }
      ],
      model: 'claude-sonnet-4-20250514',
      usage: {
        input_tokens: 100,
        output_tokens: 200
      }
    }
  },

  FATHOM_WEBHOOK: {
    service: 'Fathom (Incoming Webhook)',
    endpoint: 'POST {YOUR_WEB_APP_URL}',
    auth: 'None (webhook receives data)',
    request: {
      meeting_title: 'Weekly Sync with Client',
      meeting_date: '2025-01-15T10:00:00Z',
      transcript: 'Full meeting transcript...',
      summary: 'Meeting summary generated by Fathom',
      action_items: [
        {
          description: 'Follow up on proposal',
          assignee: 'John',
          due_date: '2025-01-20'
        }
      ],
      participants: [
        {
          name: 'John Doe',
          email: 'john@example.com'
        },
        {
          name: 'Jane Smith',
          email: 'jane@client.com'
        }
      ]
    },
    response: {
      status: 'success',
      client_name: 'Client Name',
      draft_id: 'draft_abc123'
    }
  },

  GMAIL_API_CREATE_FILTER: {
    service: 'Gmail API (Advanced Service)',
    endpoint: 'Gmail.Users.Settings.Filters.create()',
    auth: 'OAuth (automatic via Apps Script)',
    request: {
      criteria: {
        query: 'from:client@example.com'
      },
      action: {
        addLabelIds: ['Label_123']
      }
    },
    response: {
      id: 'filter_abc123',
      criteria: {
        query: 'from:client@example.com'
      },
      action: {
        addLabelIds: ['Label_123']
      }
    }
  },

  CALENDAR_API_PATCH_EVENT: {
    service: 'Calendar API (Advanced Service)',
    endpoint: 'Calendar.Events.patch()',
    auth: 'OAuth (automatic via Apps Script)',
    request: {
      attachments: [
        {
          fileUrl: 'https://docs.google.com/document/d/abc123',
          title: 'Meeting Notes - Client',
          mimeType: 'application/vnd.google-apps.document'
        }
      ]
    },
    response: {
      id: 'event_abc123',
      summary: 'Meeting with Client',
      attachments: [
        {
          fileUrl: 'https://docs.google.com/document/d/abc123',
          title: 'Meeting Notes - Client'
        }
      ]
    }
  }
};

/**
 * Creates the hidden JSON_Formats sheet with API documentation.
 *
 * @param {Spreadsheet} ss - The spreadsheet object
 */
function createJsonFormatsSheet(ss) {
  const sheetName = 'JSON_Formats';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);

    // Set up headers
    sheet.getRange(1, 1, 1, 5).setValues([['api_name', 'service', 'endpoint', 'request_json', 'response_json']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Add API format examples
    const rows = [];
    for (const [key, format] of Object.entries(API_JSON_FORMATS)) {
      rows.push([
        key,
        format.service,
        `${format.endpoint}\nAuth: ${format.auth}`,
        format.request ? JSON.stringify(format.request, null, 2) : 'N/A',
        JSON.stringify(format.response, null, 2)
      ]);
    }

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 5).setValues(rows);
    }

    // Set column widths
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 400);
    sheet.setColumnWidth(5, 400);

    // Wrap text in JSON columns
    sheet.getRange(2, 4, rows.length, 2).setWrap(true);

    // Hide the sheet
    sheet.hideSheet();

    Logger.log('Created hidden JSON_Formats sheet with API documentation');
  }
}
