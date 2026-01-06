/**
 * ClientOnboarding.gs - Client Setup Automation
 *
 * This module handles automatic resource creation when new clients are added:
 * - Creates Google Doc for meeting notes (with naming schema)
 * - Creates Todoist project for the client
 * - Updates Client_Registry with the created resource URLs/IDs
 * - Can be triggered manually or on a schedule
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const ONBOARDING_CONFIG = {
  // Google Doc naming schema: "Client Notes - [Client Name]"
  DOC_NAME_TEMPLATE: 'Client Notes - {client_name}',

  // Folder ID for storing client docs (optional - leave empty for root)
  // Set this in Script Properties as CLIENT_DOCS_FOLDER_ID
  DOCS_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('CLIENT_DOCS_FOLDER_ID') || '',

  // Initial doc content template
  DOC_TEMPLATE: `
# {client_name} - Meeting Notes

This document contains running meeting notes and agendas for {client_name}.

---

## Table of Contents
- Meeting notes are appended below in reverse chronological order
- Each entry includes: Meeting Agenda, Meeting Notes, and Action Items

---

`
};

// ============================================================================
// MAIN ONBOARDING FUNCTIONS
// ============================================================================

/**
 * Processes all clients in the registry and creates missing resources.
 * Can be run manually or on a daily schedule.
 */
function processNewClients() {
  Logger.log('Checking for clients needing onboarding...');

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CLIENT_REGISTRY);

  if (!sheet) {
    Logger.log('Client_Registry sheet not found');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find column indices
  const colIndex = {
    client_id: headers.indexOf('client_id'),
    client_name: headers.indexOf('client_name'),
    google_doc_url: headers.indexOf('google_doc_url'),
    todoist_project_id: headers.indexOf('todoist_project_id')
  };

  let clientsProcessed = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const clientId = row[colIndex.client_id];
    const clientName = row[colIndex.client_name];
    const googleDocUrl = row[colIndex.google_doc_url];
    const todoistProjectId = row[colIndex.todoist_project_id];

    if (!clientId || !clientName) {
      continue; // Skip incomplete rows
    }

    let updated = false;

    // Create Google Doc if missing
    if (!googleDocUrl) {
      const docUrl = createClientDoc(clientName);
      if (docUrl) {
        sheet.getRange(i + 1, colIndex.google_doc_url + 1).setValue(docUrl);
        updated = true;
        Logger.log(`Created Google Doc for ${clientName}`);
      }
    }

    // Create Todoist project if missing
    if (!todoistProjectId) {
      const projectId = createTodoistProject(clientName);
      if (projectId) {
        sheet.getRange(i + 1, colIndex.todoist_project_id + 1).setValue(projectId);
        updated = true;
        Logger.log(`Created Todoist project for ${clientName}`);
      }
    }

    if (updated) {
      clientsProcessed++;
      logProcessing('CLIENT_ONBOARD', clientId, `Onboarded client: ${clientName}`, 'success');
    }
  }

  Logger.log(`Onboarding complete. Processed ${clientsProcessed} clients.`);
}

/**
 * Onboards a single client by name.
 * Useful for manual onboarding or API calls.
 *
 * @param {string} clientName - The client name to onboard
 * @returns {Object} Result with created resource info
 */
function onboardClient(clientName) {
  const result = {
    success: false,
    clientName: clientName,
    googleDocUrl: null,
    todoistProjectId: null,
    errors: []
  };

  // Create Google Doc
  try {
    result.googleDocUrl = createClientDoc(clientName);
  } catch (error) {
    result.errors.push(`Failed to create Google Doc: ${error.message}`);
  }

  // Create Todoist project
  try {
    result.todoistProjectId = createTodoistProject(clientName);
  } catch (error) {
    result.errors.push(`Failed to create Todoist project: ${error.message}`);
  }

  result.success = result.errors.length === 0;

  return result;
}

// ============================================================================
// GOOGLE DOC CREATION
// ============================================================================

/**
 * Creates a Google Doc for a client using the naming schema.
 *
 * @param {string} clientName - The client name
 * @returns {string|null} The URL of the created doc or null
 */
function createClientDoc(clientName) {
  try {
    // Generate doc name from template
    const docName = ONBOARDING_CONFIG.DOC_NAME_TEMPLATE.replace('{client_name}', clientName);

    // Create the document
    const doc = DocumentApp.create(docName);
    const docId = doc.getId();

    // Add initial content
    const body = doc.getBody();
    const initialContent = ONBOARDING_CONFIG.DOC_TEMPLATE
      .replace(/{client_name}/g, clientName);

    body.setText(initialContent);

    // Apply basic formatting
    const paragraphs = body.getParagraphs();
    if (paragraphs.length > 0) {
      paragraphs[0].setHeading(DocumentApp.ParagraphHeading.HEADING1);
    }

    doc.saveAndClose();

    // Move to folder if specified
    if (ONBOARDING_CONFIG.DOCS_FOLDER_ID) {
      moveDocToFolder(docId, ONBOARDING_CONFIG.DOCS_FOLDER_ID);
    }

    const docUrl = doc.getUrl();
    Logger.log(`Created doc: ${docUrl}`);

    return docUrl;

  } catch (error) {
    Logger.log(`Failed to create Google Doc for ${clientName}: ${error.message}`);
    logProcessing('DOC_CREATE_ERROR', null, `Failed to create doc for ${clientName}: ${error.message}`, 'error');
    return null;
  }
}

/**
 * Moves a document to a specific folder.
 *
 * @param {string} docId - The document ID
 * @param {string} folderId - The destination folder ID
 */
function moveDocToFolder(docId, folderId) {
  try {
    const file = DriveApp.getFileById(docId);
    const folder = DriveApp.getFolderById(folderId);

    file.moveTo(folder);
    Logger.log(`Moved doc ${docId} to folder ${folderId}`);
  } catch (error) {
    Logger.log(`Failed to move doc to folder: ${error.message}`);
  }
}

/**
 * Creates a folder for client documents if it doesn't exist.
 *
 * @param {string} folderName - Name of the folder
 * @returns {string} The folder ID
 */
function createClientDocsFolder(folderName = 'Client Meeting Notes') {
  // Check if folder already exists
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    const folder = folders.next();
    return folder.getId();
  }

  // Create new folder
  const folder = DriveApp.createFolder(folderName);
  const folderId = folder.getId();

  // Store in script properties
  PropertiesService.getScriptProperties().setProperty('CLIENT_DOCS_FOLDER_ID', folderId);

  Logger.log(`Created client docs folder: ${folderId}`);
  return folderId;
}

// ============================================================================
// TODOIST PROJECT CREATION
// ============================================================================

/**
 * Creates a Todoist project for a client.
 *
 * @param {string} clientName - The client name
 * @returns {string|null} The project ID or null
 */
function createTodoistProject(clientName) {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    Logger.log('Todoist API token not configured');
    return null;
  }

  try {
    const url = 'https://api.todoist.com/rest/v2/projects';

    const payload = {
      name: clientName
    };

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
      Logger.log(`Todoist API error: ${responseCode} - ${response.getContentText()}`);
      return null;
    }

    const project = JSON.parse(response.getContentText());
    Logger.log(`Created Todoist project: ${project.id} for ${clientName}`);

    return project.id;

  } catch (error) {
    Logger.log(`Failed to create Todoist project for ${clientName}: ${error.message}`);
    logProcessing('TODOIST_CREATE_ERROR', null, `Failed to create project for ${clientName}: ${error.message}`, 'error');
    return null;
  }
}

/**
 * Gets all Todoist projects.
 *
 * @returns {Object[]} Array of project objects
 */
function getTodoistProjects() {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    return [];
  }

  try {
    const url = 'https://api.todoist.com/rest/v2/projects';

    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiToken}`
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      return [];
    }

    return JSON.parse(response.getContentText());

  } catch (error) {
    Logger.log(`Failed to fetch Todoist projects: ${error.message}`);
    return [];
  }
}

/**
 * Finds a Todoist project by name.
 *
 * @param {string} projectName - The project name to find
 * @returns {Object|null} The project object or null
 */
function findTodoistProject(projectName) {
  const projects = getTodoistProjects();
  return projects.find(p => p.name.toLowerCase() === projectName.toLowerCase()) || null;
}

// ============================================================================
// SPREADSHEET WATCH (onChange trigger)
// ============================================================================

/**
 * Trigger function that runs when the spreadsheet is edited.
 * Checks for new clients and onboards them automatically.
 *
 * To enable: Create an "onChange" trigger for this function.
 *
 * @param {Object} e - The event object
 */
function onSpreadsheetChange(e) {
  // Only process edit events on Client_Registry
  if (e.changeType !== 'EDIT') {
    return;
  }

  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEETS.CLIENT_REGISTRY) {
    return;
  }

  // Process new clients
  processNewClients();
}

/**
 * Sets up the onChange trigger for automatic client onboarding.
 */
function setupOnboardingTrigger() {
  // Remove existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onSpreadsheetChange') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create new trigger
  ScriptApp.newTrigger('onSpreadsheetChange')
    .forSpreadsheet(CONFIG.SPREADSHEET_ID)
    .onChange()
    .create();

  Logger.log('Onboarding trigger set up successfully');
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Validates that all required resources exist for a client.
 *
 * @param {Object} client - The client object
 * @returns {Object} Validation result
 */
function validateClientResources(client) {
  const result = {
    isValid: true,
    missingResources: []
  };

  // Check Google Doc
  if (!client.google_doc_url) {
    result.missingResources.push('Google Doc');
    result.isValid = false;
  } else {
    // Verify doc is accessible
    try {
      const docId = extractDocIdFromUrl(client.google_doc_url);
      DocumentApp.openById(docId);
    } catch (e) {
      result.missingResources.push('Google Doc (inaccessible)');
      result.isValid = false;
    }
  }

  // Check Todoist project
  if (!client.todoist_project_id) {
    result.missingResources.push('Todoist Project');
    result.isValid = false;
  }

  return result;
}

/**
 * Generates a report of all clients and their resource status.
 *
 * @returns {Object[]} Array of client status objects
 */
function getClientResourceReport() {
  const clients = getClientRegistry();
  const report = [];

  for (const client of clients) {
    const validation = validateClientResources(client);

    report.push({
      client_id: client.client_id,
      client_name: client.client_name,
      has_google_doc: !!client.google_doc_url,
      has_todoist_project: !!client.todoist_project_id,
      is_fully_configured: validation.isValid,
      missing_resources: validation.missingResources
    });
  }

  return report;
}
