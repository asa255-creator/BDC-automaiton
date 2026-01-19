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
  // ============================================================================
  // WEEKLY BRIEFING - Comprehensive AI-generated strategic briefing
  // ============================================================================
  WEEKLY_BRIEFING_CLAUDE_PROMPT: `You are an executive assistant creating a strategic weekly briefing for a Blue Dot co-founder.
This briefing is a personal decision-making tool focused on client attention priorities, task ordering, fire identification, gap anticipation, time allocation, and meeting conflict resolution.

Important output note:
The final output must be the raw HTML code for the email body, returned as plain text.
Do not render or display the HTML — just return the HTML source code itself, exactly as it should appear inside an email.
The HTML must include proper headings, paragraphs, and bullet lists so the email displays well when sent.
No external CSS or links should be included. Inline structure only.

Here is today's date: {current_date}

Here are all recent emails labeled "Action Today":
{action_today_emails}

Here are all recent emails labeled "Action This Week":
{action_this_week_emails}

Here are the undone tasks by project (from Todoist):
{todoist_tasks}

Here's a quick summary of the upcoming week's calendar:
{calendar_summary}

🎯 Executive Summary (2–3 sentences)
Identify the highest-stakes deliverable or deadline this week (Critical Issue).
List two or three actions that require immediate attention (Priority Focus).
Note any client or strategic decisions required (Key Decision Points).

Time Audit and Current Schedule Status
• Today's Schedule – key meetings and time blocks
• Schedule Efficiency Flags – overdue items, bottlenecks, or conflicts
• Meeting Conflicts – identify and resolve scheduling issues

Project Status Dashboard
Use the following format for each active project derived from tasks and emails:
🔴 [Project Name] – Blocked / Critical / Overdue
🟡 [Project Name] – Attention Needed / Dependencies
🟢 [Project Name] – On Track / Strong Momentum

For each project include:
• Status – concise summary of current state
• Recent Activity – key developments from recent emails or updates
• Critical Deliverables – must be completed this week

Priority Actions
Break down urgent tasks into time-boxed actions:
• [Project] Immediate Action (time estimate)
• [Project] Coordination (time estimate)
• [Project] Preparation (time estimate)
Use short bullet points with clear next steps.

Upcoming Critical Events
List the week's schedule and mark:
• Meeting conflicts
• Critical client deliverables
• Prep time needed
• Travel or location requirements

Communication Priorities
• Critical follow-ups from recent emails
• Recent communication gaps (over 48 hours without response)
• Client relationship temperature – concerns or opportunities

Revenue and Relationship Risks
• Immediate Risks – what could go wrong this week
• Revenue Opportunities – new work or expansions
• Relationship Maintenance – clients needing attention

Clarifying Questions for Execution
End with three or four strategic questions about:
• Resource allocation
• Priority conflicts
• Information gaps
• Backup plans

🎯 Bottom Line
Summarize the week's strategic focus, execution priorities, and success metrics in one concise paragraph.

Prioritization Framework
When organizing information, prioritize by:
• Client Impact – revenue risk or opportunity
• Timeline Urgency – hard deadlines versus flexible timing
• Dependency Chains – what blocks other work
• Strategic Value – long-term relationship or business impact

Tone and Style
• Decisive and strategic, not descriptive
• Use emojis (🎯 🔴 🟡 🟢) as visual status markers
• Highlight all client names in bold using HTML tags in the output
• Use paragraph and list structure for clarity
• Target approximately 1,500 to 2,000 words

Automation Helpers and Output Instructions
Use the email summaries at the top to locate follow-ups and themes.
Use the task list to populate the project dashboard and time-boxed actions.
Use the calendar and conflict summaries to evaluate workload and overlaps.

Apply time-sensitivity flags automatically:
🔴 Items overdue by more than three days or awaiting reply over twenty-four hours
🟡 Items with dependencies or idle for five days
🟢 Items showing recent progress

Highlight projects with no activity in the past week.
Note any deadlines within the next five business days.
Exclude meeting note summaries and external document references.
List call-time sessions if they appear, but assume attendance is not required.

At the end, return one plain-text block containing the full HTML code that will produce a polished, structured email body.
The output must start with an <h1> heading for the 🎯 Executive Summary section and include all subsequent sections as valid HTML elements.
Return only the HTML source code, not rendered output or explanations.`,

  // ============================================================================
  // DAILY BRIEFING - AI-generated daily focus
  // ============================================================================
  DAILY_BRIEFING_CLAUDE_PROMPT: `You are an executive assistant creating a strategic daily briefing.
This briefing is a personal decision-making tool focused on today's priorities, urgent tasks, and schedule optimization.

Important output note:
The final output must be the raw HTML code for the email body, returned as plain text.
Do not render or display the HTML — just return the HTML source code itself.
The HTML must include proper headings, paragraphs, and bullet lists.
No external CSS or links. Inline structure only.

Here is today's date: {current_date}

Today's calendar events:
{todays_calendar}

Tasks due today or overdue (from Todoist):
{todays_tasks}

Recent urgent emails:
{urgent_emails}

Generate a daily briefing with these sections:

🎯 Today's Focus (2-3 sentences)
The single most important outcome needed today.

📅 Schedule Overview
• List today's meetings with times
• Flag any conflicts or tight transitions
• Note prep time needed

✅ Priority Tasks (time-boxed)
• [Client/Project] Task description (estimated time)
• Maximum 5-7 items, ordered by urgency

⚠️ Alerts
• Overdue items (🔴)
• Items needing attention (🟡)
• Awaiting responses

📧 Communication Queue
• Must-send emails today
• Follow-ups needed

🎯 End of Day Success
Define what "done" looks like for today in one sentence.

Tone: Decisive, action-oriented, concise.
Use emojis as visual markers.
Bold client names.
Target 500-800 words.

Return only the HTML source code.`,

  // ============================================================================
  // MEETING AGENDA - AI-generated meeting preparation
  // ============================================================================
  AGENDA_CLAUDE_PROMPT: `You are an executive assistant preparing a meeting agenda.

Meeting Details:
- Title: {event_title}
- Client: {client_name}
- Date/Time: {date_time}
- Duration: {duration} minutes

Outstanding Todoist tasks for this client:
{todoist_section}

Recent email activity with this client:
{emails_section}

Previous meeting notes:
{notes_section}

Action items from last meeting:
{action_items_section}

Generate a strategic meeting agenda with:

1. 🎯 Meeting Objective (1 sentence)
   What must be accomplished in this meeting?

2. 📋 Agenda Items (with time allocations totaling {duration} minutes)
   • Item name (X min) - brief description
   • Prioritize items that address outstanding tasks and action items
   • Include time for client questions/discussion

3. ⚠️ Items Requiring Attention
   • Overdue tasks to address
   • Unanswered emails to mention
   • Previous action items not completed

4. 📝 Preparation Notes
   • Key points to remember
   • Questions to ask
   • Decisions needed

5. ✅ Desired Outcomes
   • What should be agreed upon
   • Next steps to confirm
   • Follow-up commitments

Format as clean HTML with inline styles.
Use bullet points and clear headings.
Bold the client name throughout.
Keep it actionable and focused.

Return only the HTML source code.`,

  // Email wrapper template for agenda (the AI content goes inside)
  AGENDA_EMAIL_TEMPLATE: `<div style="font-family: Arial, sans-serif; max-width: 800px;">
{agenda_content}
</div>`,

  // ============================================================================
  // LEGACY TEMPLATES (for non-AI formatted sections if needed)
  // ============================================================================
  DAILY_OUTLOOK_HEADER: `<h1 style="color: #1a73e8;">📅 Daily Outlook - {date}</h1>`,

  DAILY_OUTLOOK_ALERTS: `<div style="background-color: #fff3cd; padding: 15px; margin-bottom: 20px; border-radius: 5px; border-left: 4px solid #ffc107;">
<h2 style="margin-top: 0; color: #856404;">⚠️ Alerts</h2>
{alerts_content}
</div>`,

  WEEKLY_OUTLOOK_HEADER: `<h1 style="color: #1a73e8;">📊 Weekly Outlook - Week of {date}</h1>`,

  WEEKLY_OUTLOOK_SUMMARY: `<div style="background-color: #e3f2fd; padding: 15px; margin-bottom: 20px; border-radius: 5px; border-left: 4px solid #1a73e8;">
<h2 style="margin-top: 0; color: #1565c0;">📈 Week Summary</h2>
<p><strong>Total Meetings:</strong> {meeting_count}</p>
<p><strong>Total Tasks:</strong> {task_count}</p>
<p><strong>Clients with Activity:</strong> {client_count}</p>
</div>`
};

/**
 * Prompt metadata for the editor UI
 */
const PROMPT_METADATA = {
  WEEKLY_BRIEFING_CLAUDE_PROMPT: {
    label: '📊 Weekly Briefing AI Prompt',
    description: 'Comprehensive Claude prompt for strategic weekly briefings. This is the main prompt that generates your weekly executive summary with project dashboards, priority actions, and risk analysis.',
    variables: ['{current_date}', '{action_today_emails}', '{action_this_week_emails}', '{todoist_tasks}', '{calendar_summary}']
  },
  DAILY_BRIEFING_CLAUDE_PROMPT: {
    label: '📅 Daily Briefing AI Prompt',
    description: 'Claude prompt for daily focus briefings. Generates your daily priorities, schedule overview, and task list.',
    variables: ['{current_date}', '{todays_calendar}', '{todays_tasks}', '{urgent_emails}']
  },
  AGENDA_CLAUDE_PROMPT: {
    label: '📋 Meeting Agenda AI Prompt',
    description: 'Claude prompt for meeting preparation. Generates strategic agendas with time allocations, preparation notes, and desired outcomes.',
    variables: ['{event_title}', '{client_name}', '{date_time}', '{duration}', '{todoist_section}', '{emails_section}', '{notes_section}', '{action_items_section}']
  },
  AGENDA_EMAIL_TEMPLATE: {
    label: 'Agenda Email Wrapper',
    description: 'HTML wrapper template for agenda emails. The AI-generated content goes inside {agenda_content}.',
    variables: ['{agenda_content}']
  },
  DAILY_OUTLOOK_HEADER: {
    label: 'Daily Header (Legacy)',
    description: 'Header template for non-AI daily outlook. Used if AI generation is disabled.',
    variables: ['{date}']
  },
  DAILY_OUTLOOK_ALERTS: {
    label: 'Daily Alerts Section (Legacy)',
    description: 'Alerts template for non-AI daily outlook.',
    variables: ['{alerts_content}']
  },
  WEEKLY_OUTLOOK_HEADER: {
    label: 'Weekly Header (Legacy)',
    description: 'Header template for non-AI weekly outlook. Used if AI generation is disabled.',
    variables: ['{date}']
  },
  WEEKLY_OUTLOOK_SUMMARY: {
    label: 'Weekly Summary Section (Legacy)',
    description: 'Summary template for non-AI weekly outlook.',
    variables: ['{meeting_count}', '{task_count}', '{client_count}']
  }
};

// ============================================================================
// PROMPT SHEET MANAGEMENT
// ============================================================================

/**
 * Fallback models if API fetch fails.
 * These are only used when the Claude API is unreachable.
 */
const FALLBACK_MODELS = [
  { id: 'claude-sonnet-4-20250514', name: 'Claude Sonnet 4' },
  { id: 'claude-3-haiku-20240307', name: 'Claude 3 Haiku' }
];

/**
 * Cache duration for models list (24 hours in milliseconds)
 */
const MODELS_CACHE_DURATION_MS = 24 * 60 * 60 * 1000;

/**
 * Fetches available Claude models from the Anthropic API.
 * Caches results for 24 hours to minimize API calls.
 *
 * @param {boolean} forceRefresh - If true, bypasses cache and fetches fresh data
 * @returns {Object[]} Array of model objects with id and name
 */
function fetchAvailableModelsFromAPI(forceRefresh) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    Logger.log('Claude API key not configured - using fallback models');
    return FALLBACK_MODELS;
  }

  // Check cache first (unless force refresh)
  if (!forceRefresh) {
    const cachedModels = props.getProperty('CACHED_MODELS');
    const cacheTimestamp = props.getProperty('CACHED_MODELS_TIMESTAMP');

    if (cachedModels && cacheTimestamp) {
      const cacheAge = Date.now() - parseInt(cacheTimestamp, 10);
      if (cacheAge < MODELS_CACHE_DURATION_MS) {
        try {
          return JSON.parse(cachedModels);
        } catch (e) {
          Logger.log('Failed to parse cached models, fetching fresh');
        }
      }
    }
  }

  // Fetch from API
  try {
    const url = 'https://api.anthropic.com/v1/models';

    const options = {
      method: 'GET',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`Failed to fetch models: ${responseCode} - ${response.getContentText()}`);
      return FALLBACK_MODELS;
    }

    const result = JSON.parse(response.getContentText());

    // Extract and format models
    const models = (result.data || [])
      .filter(m => m.type === 'model')
      .map(m => ({
        id: m.id,
        name: m.display_name || formatModelName(m.id)
      }))
      .sort((a, b) => a.name.localeCompare(b.name));

    if (models.length === 0) {
      Logger.log('No models returned from API - using fallback');
      return FALLBACK_MODELS;
    }

    // Cache the results
    props.setProperty('CACHED_MODELS', JSON.stringify(models));
    props.setProperty('CACHED_MODELS_TIMESTAMP', Date.now().toString());

    Logger.log(`Fetched ${models.length} models from Claude API`);
    return models;

  } catch (e) {
    Logger.log(`Error fetching models: ${e.message}`);
    return FALLBACK_MODELS;
  }
}

/**
 * Formats a model ID into a readable name.
 * e.g., "claude-3-haiku-20240307" -> "Claude 3 Haiku"
 *
 * @param {string} modelId - The model ID
 * @returns {string} Formatted display name
 */
function formatModelName(modelId) {
  return modelId
    .replace(/^claude-/, 'Claude ')
    .replace(/-(\d{8})$/, '') // Remove date suffix
    .replace(/-/g, ' ')
    .replace(/\b\w/g, c => c.toUpperCase());
}

/**
 * Gets the model ID to use for a prompt.
 * Looks up the stored preference, or returns a sensible default.
 *
 * @param {string} promptKey - The prompt key
 * @returns {string} The actual model ID to use
 */
function getModelForPrompt(promptKey) {
  const storedModel = getPromptModel(promptKey);

  // If it's already a full model ID, return it
  if (storedModel && storedModel.startsWith('claude-')) {
    return storedModel;
  }

  // Legacy: convert tier names to model IDs
  const models = fetchAvailableModelsFromAPI(false);

  if (storedModel === 'haiku') {
    const haiku = models.find(m => m.id.includes('haiku'));
    return haiku ? haiku.id : models[0]?.id || FALLBACK_MODELS[0].id;
  }

  if (storedModel === 'sonnet') {
    const sonnet = models.find(m => m.id.includes('sonnet'));
    return sonnet ? sonnet.id : models[0]?.id || FALLBACK_MODELS[0].id;
  }

  // Default to first available model
  return models[0]?.id || FALLBACK_MODELS[0].id;
}

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

    // Set up headers (now includes model_preference column)
    sheet.getRange(1, 1, 1, 3).setValues([['prompt_key', 'prompt_value', 'model_preference']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Add default prompts with haiku as default model
    const promptKeys = Object.keys(DEFAULT_PROMPTS);
    const rows = promptKeys.map(key => [key, DEFAULT_PROMPTS[key], 'haiku']);

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }

    // Set column widths
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 800);
    sheet.setColumnWidth(3, 100);

    // Hide the sheet
    sheet.hideSheet();

    Logger.log('Created hidden Prompts sheet with defaults');
  } else {
    // Check if model_preference column exists, add if not
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('model_preference')) {
      const lastCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lastCol).setValue('model_preference');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      // Set default model for existing rows
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const defaults = Array(lastRow - 1).fill(['haiku']);
        sheet.getRange(2, lastCol, lastRow - 1, 1).setValues(defaults);
      }
      sheet.setColumnWidth(lastCol, 100);
      Logger.log('Added model_preference column to existing Prompts sheet');
    }
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
 * Gets the model preference for a prompt.
 *
 * @param {string} promptKey - The prompt key
 * @returns {string} The model tier ('haiku' or 'sonnet'), defaults to 'haiku'
 */
function getPromptModel(promptKey) {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) return 'haiku';

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('Prompts');
    if (!sheet) return 'haiku';

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const modelCol = headers.indexOf('model_preference');

    if (modelCol === -1) return 'haiku';

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === promptKey) {
        return data[i][modelCol] || 'haiku';
      }
    }

    return 'haiku';
  } catch (e) {
    Logger.log(`Error getting model for ${promptKey}: ${e.message}`);
    return 'haiku';
  }
}

/**
 * Maps model tier to actual model ID.
 *
 * @param {string} tier - The model tier ('haiku', 'sonnet', 'opus')
 * @returns {string} The actual Claude model ID
 */
function getModelIdForTier(tier) {
  // Map tier names to actual model IDs
  const tierMap = {
    'haiku': 'claude-3-haiku-20240307',
    'sonnet': 'claude-sonnet-4-20250514',
    'opus': 'claude-opus-4-20241120'
  };

  return tierMap[tier] || tierMap['haiku']; // Default to haiku if unknown tier
}

/**
 * Gets the full model ID for a prompt (resolves tier to actual model ID).
 *
 * @param {string} promptKey - The prompt key
 * @returns {string} The actual model ID to use
 */
function getModelForPrompt(promptKey) {
  const tier = getPromptModel(promptKey);
  return getModelIdForTier(tier);
}

/**
 * Sets a prompt value in the Prompts sheet.
 *
 * @param {string} promptKey - The prompt key
 * @param {string} promptValue - The new prompt value
 * @param {string} modelPreference - Optional model preference ('haiku' or 'sonnet')
 */
function setPrompt(promptKey, promptValue, modelPreference) {
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
  const headers = data[0];
  const modelCol = headers.indexOf('model_preference');
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === promptKey) {
      sheet.getRange(i + 1, 2).setValue(promptValue);
      // Update model preference if provided and column exists
      if (modelPreference && modelCol !== -1) {
        sheet.getRange(i + 1, modelCol + 1).setValue(modelPreference);
      }
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([promptKey, promptValue, modelPreference || 'haiku']);
  }

  Logger.log(`Updated prompt: ${promptKey} (model: ${modelPreference || 'unchanged'})`);
}

/**
 * Gets all prompts for the editor.
 *
 * @returns {Object[]} Array of prompt objects with key, value, label, description, variables, model
 */
function getAllPromptsForEditor() {
  const prompts = [];

  for (const key of Object.keys(DEFAULT_PROMPTS)) {
    const value = getPrompt(key);
    const model = getPromptModel(key);
    const metadata = PROMPT_METADATA[key] || {};

    prompts.push({
      key: key,
      value: value,
      model: model,
      label: metadata.label || key,
      description: metadata.description || '',
      variables: metadata.variables || [],
      isDefault: value === DEFAULT_PROMPTS[key]
    });
  }

  return prompts;
}

/**
 * Gets available models for the editor dropdown.
 * Fetches dynamically from Claude API with caching.
 *
 * @returns {Object[]} Array of {value, label} for dropdown
 */
function getAvailableModels() {
  const models = fetchAvailableModelsFromAPI(false);

  return models.map(m => ({
    value: m.id,
    label: m.name
  }));
}

/**
 * Refreshes the models cache and returns updated list.
 * Called from editor UI when user clicks "Refresh Models".
 *
 * @returns {Object} Result with models array or error
 */
function refreshModelsCache() {
  try {
    const models = fetchAvailableModelsFromAPI(true);

    return {
      success: true,
      models: models.map(m => ({ value: m.id, label: m.name })),
      message: `Found ${models.length} available models`
    };
  } catch (e) {
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Saves multiple prompts from the editor.
 *
 * @param {Object[]} prompts - Array of {key, value, model} objects
 * @returns {Object} Result with success status
 */
function savePromptsFromEditor(prompts) {
  try {
    for (const prompt of prompts) {
      if (prompt.key && prompt.value !== undefined) {
        setPrompt(prompt.key, prompt.value, prompt.model);
      }
    }

    return { success: true, message: `Saved ${prompts.length} prompts` };
  } catch (e) {
    Logger.log(`Error saving prompts: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Compresses a prompt using Claude AI while preserving all requirements.
 * Uses Sonnet for better comprehension (Haiku is too aggressive).
 *
 * @param {string} promptText - The original prompt text to compress
 * @param {string} promptKey - Optional prompt key to look up available variables
 * @returns {Object} Result with compressed text or error
 */
function compressPromptWithAI(promptText, promptKey) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    return { success: false, error: 'Claude API key not configured. Add it in Settings.' };
  }

  // Get available variables for this prompt type
  let variablesSection = '';
  if (promptKey && PROMPT_METADATA[promptKey] && PROMPT_METADATA[promptKey].variables) {
    const variables = PROMPT_METADATA[promptKey].variables;
    variablesSection = `
CRITICAL - These variables MUST appear in your output exactly as shown:
${variables.join(', ')}
`;
  }

  const compressionPrompt = `You are a CONSERVATIVE prompt editor. Your goal is to make MINOR optimizations to reduce token count while keeping the prompt's meaning, structure, and instructions fully intact.

THIS IS NOT AGGRESSIVE COMPRESSION. You are making small edits, not rewriting.

WHAT YOU MUST KEEP (do not remove or summarize):
- ALL instructions (every single instruction must remain)
- ALL section headers and labels
- ALL formatting requirements (HTML, timezone, styling rules)
- ALL variable placeholders like {variable_name} or {{variable_name}} exactly as written
- ALL context descriptions (what data is being provided)
- ALL output format specifications
- ALL constraints and rules
- ALL examples if present
${variablesSection}
WHAT YOU CAN DO (minor optimizations only):
- Remove filler words ("please", "kindly", "make sure to")
- Shorten wordy phrases ("in order to" → "to", "make sure that" → "ensure")
- Remove redundant repetition of the same instruction
- Combine adjacent sentences that say the same thing

WHAT YOU CANNOT DO:
- Remove any instruction, even if it seems obvious
- Summarize multiple instructions into one
- Remove context about what the variables contain
- Change the structure or organization
- Remove formatting/output requirements
- Strip out section labels or headers

The compressed prompt MUST produce IDENTICAL output when used. If in doubt, keep it.

Original prompt:
---
${promptText}
---

Return the optimized prompt. Keep it nearly identical to the original - only make minor word-level improvements.`;

  try {
    const url = 'https://api.anthropic.com/v1/messages';

    // Use Sonnet for better comprehension - Haiku is too aggressive
    const models = fetchAvailableModelsFromAPI(false);
    const sonnet = models.find(m => m.id.includes('sonnet'));
    const model = sonnet ? sonnet.id : models[0]?.id || FALLBACK_MODELS[0].id;

    const payload = {
      model: model,
      max_tokens: 4096,
      messages: [
        { role: 'user', content: compressionPrompt }
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
      const errorText = response.getContentText();
      Logger.log(`Claude API error: ${responseCode} - ${errorText}`);
      return { success: false, error: `API error (${responseCode}): ${errorText}` };
    }

    const result = JSON.parse(response.getContentText());
    const compressedText = result.content[0].text;

    // Calculate token savings estimate (rough: 4 chars per token)
    const originalTokens = Math.ceil(promptText.length / 4);
    const compressedTokens = Math.ceil(compressedText.length / 4);
    const savings = Math.round((1 - compressedTokens / originalTokens) * 100);

    return {
      success: true,
      compressed: compressedText,
      originalLength: promptText.length,
      compressedLength: compressedText.length,
      estimatedSavings: `~${savings}% fewer tokens`
    };

  } catch (e) {
    Logger.log(`Error compressing prompt: ${e.message}`);
    return { success: false, error: e.message };
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
