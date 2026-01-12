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
