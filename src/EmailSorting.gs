/**
 * EmailSorting.gs - Gmail Label and Filter Management
 *
 * This module manages Gmail labels and filters for client email organization:
 * - Creates and maintains client labels
 * - Generates filters based on Client_Registry
 * - Handles meeting summary and agenda sub-labels
 * - Manages internal briefing labels
 */

// ============================================================================
// MAIN SYNC FUNCTION
// ============================================================================

/**
 * Synchronizes all Gmail labels and filters based on Client_Registry.
 * Called daily at 6:00 AM by trigger.
 * Only processes clients where setup_complete is TRUE.
 */
function syncLabelsAndFilters() {
  Logger.log('Starting label and filter synchronization...');

  // Get all clients
  const allClients = getClientRegistry();

  // Filter to only clients with setup_complete = true
  const clients = allClients.filter(client => client.setup_complete === true);

  if (clients.length === 0) {
    Logger.log('No clients with setup_complete found');
    return;
  }

  Logger.log(`Processing ${clients.length} clients with setup_complete=true (${allClients.length - clients.length} skipped)`);

  // Create/update labels for each client
  for (const client of clients) {
    syncClientLabels(client);
  }

  // Create briefing labels
  createBriefingLabels();

  // Sync filters (only for setup_complete clients)
  syncFilters(clients);

  Logger.log('Label and filter synchronization completed.');
}

// ============================================================================
// LABEL MANAGEMENT
// ============================================================================

/**
 * Creates or verifies labels and filters for a specific client.
 *
 * Labels created:
 * - Client: [client_name]
 * - Client: [client_name]/Meeting Summaries
 * - Client: [client_name]/Meeting Agendas
 *
 * Filters created (if Gmail API enabled):
 * - Emails from client contacts -> Client: [client_name]
 * - Sent meeting summaries -> Client: [client_name]/Meeting Summaries
 * - Self-sent agendas -> Client: [client_name]/Meeting Agendas
 *
 * @param {Object} client - The client object
 */
function syncClientLabels(client) {
  const baseLabelName = `Client: ${client.client_name}`;

  // Create base client label
  createLabelIfNotExists(baseLabelName);

  // Create sub-labels
  createLabelIfNotExists(`${baseLabelName}/Meeting Summaries`);
  createLabelIfNotExists(`${baseLabelName}/Meeting Agendas`);

  Logger.log(`Synced labels for client: ${client.client_name}`);

  // Create filters (requires Gmail API Advanced Service)
  const contacts = parseCommaSeparatedList(client.contact_emails);

  if (contacts.length > 0) {
    // Filter for incoming emails from client contacts
    const fromCriteria = buildFromCriteria(contacts);
    if (fromCriteria) {
      createGmailApiFilter(fromCriteria, baseLabelName);
    }

    // Filter for sent meeting summaries to client (uses subject pattern from settings)
    const toCriteria = buildToCriteria(contacts);
    if (toCriteria) {
      const subjectPattern = getSubjectFilterPattern();
      const summaryCriteria = `from:me subject:"${subjectPattern}" ${toCriteria}`;
      createGmailApiFilter(summaryCriteria, `${baseLabelName}/Meeting Summaries`);
    }
  }

  // Filter for self-sent agendas (these are sent to yourself)
  const agendaCriteria = `from:me to:me subject:"Agenda:"`;
  createGmailApiFilter(agendaCriteria, `${baseLabelName}/Meeting Agendas`);

  Logger.log(`Synced filters for client: ${client.client_name}`);
}

/**
 * Creates the briefing labels for daily and weekly outlooks.
 */
function createBriefingLabels() {
  createLabelIfNotExists('Brief: Daily');
  createLabelIfNotExists('Brief: Weekly');
  Logger.log('Briefing labels created/verified');
}

/**
 * Creates a Gmail label if it doesn't already exist.
 *
 * @param {string} labelName - The name of the label to create
 * @returns {GmailLabel} The label object
 */
function createLabelIfNotExists(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log(`Created label: ${labelName}`);
  }

  return label;
}

/**
 * Gets all existing client labels.
 *
 * @returns {GmailLabel[]} Array of client labels
 */
function getClientLabels() {
  const allLabels = GmailApp.getUserLabels();
  return allLabels.filter(label => label.getName().startsWith('Client:'));
}

/**
 * Removes orphaned client labels (labels for clients no longer in registry).
 *
 * @param {Object[]} clients - Array of current client objects
 */
function removeOrphanedLabels(clients) {
  const existingLabels = getClientLabels();
  const clientNames = clients.map(c => c.client_name);

  for (const label of existingLabels) {
    const labelName = label.getName();
    // Extract client name from label
    const match = labelName.match(/^Client: ([^/]+)/);
    if (match) {
      const labelClientName = match[1];
      if (!clientNames.includes(labelClientName)) {
        // This is an orphaned label - log but don't delete automatically
        Logger.log(`Orphaned label found: ${labelName}`);
        logProcessing(
          'ORPHAN_LABEL',
          null,
          `Orphaned label found: ${labelName}`,
          'warning'
        );
      }
    }
  }
}

// ============================================================================
// FILTER MANAGEMENT
// ============================================================================

/**
 * Synchronizes Gmail filters for all clients.
 * Uses Gmail API for advanced filter operations.
 *
 * @param {Object[]} clients - Array of client objects
 */
function syncFilters(clients) {
  Logger.log('Syncing Gmail filters...');

  // Note: Gmail Apps Script doesn't have direct filter management
  // We'll use Gmail advanced service (must be enabled) or document the filters

  for (const client of clients) {
    createClientFilters(client);
  }

  // Create global internal filters
  createGlobalInternalFilters();

  Logger.log('Filter sync completed');
}

/**
 * Creates filters for a specific client.
 *
 * Filters created:
 * 1. Client Filter - Labels incoming messages from client
 * 2. Meeting Summary Sub Filter - Labels meeting summary emails
 * 3. Meeting Agenda Sub Filter - Labels agenda emails
 *
 * @param {Object} client - The client object
 */
function createClientFilters(client) {
  const contacts = parseCommaSeparatedList(client.contact_emails);

  if (contacts.length === 0) {
    Logger.log(`No contact emails for client: ${client.client_name}`);
    return;
  }

  // Build filter criteria
  const fromCriteria = buildFromCriteria(contacts);
  const toCriteria = buildToCriteria(contacts);

  // Log filter specifications (actual creation requires Gmail API)
  logFilterSpec('CLIENT_FILTER', client.client_name, {
    criteria: fromCriteria,
    action: `Apply label: Client: ${client.client_name}`
  });

  // Use dynamic subject pattern from settings
  const subjectPattern = getSubjectFilterPattern();
  logFilterSpec('SUMMARY_FILTER', client.client_name, {
    criteria: `from:me subject:'${subjectPattern}' ${toCriteria}`,
    action: `Apply label: Client: ${client.client_name}/Meeting Summaries`
  });

  logFilterSpec('AGENDA_FILTER', client.client_name, {
    criteria: `from:me to:me subject:'Agenda: ${client.client_name}'`,
    action: `Apply label: Client: ${client.client_name}/Meeting Agendas`
  });
}

/**
 * Builds the 'from' criteria for Gmail filter.
 *
 * @param {string[]} contacts - Array of contact emails
 * @returns {string} Gmail search query for from criteria
 */
function buildFromCriteria(contacts) {
  const parts = [];

  for (const contact of contacts) {
    parts.push(`from:${contact}`);
  }

  if (parts.length === 0) {
    return '';
  }

  if (parts.length === 1) {
    return parts[0];
  }

  return `{${parts.join(' ')}}`;
}

/**
 * Builds the 'to' criteria for Gmail filter.
 *
 * @param {string[]} contacts - Array of contact emails
 * @returns {string} Gmail search query for to criteria
 */
function buildToCriteria(contacts) {
  const parts = [];

  for (const contact of contacts) {
    parts.push(`to:${contact}`);
  }

  if (parts.length === 0) {
    return '';
  }

  if (parts.length === 1) {
    return parts[0];
  }

  return `{${parts.join(' ')}}`;
}

/**
 * Creates global internal filters for briefing emails.
 */
function createGlobalInternalFilters() {
  // Daily Outlook Filter
  logFilterSpec('DAILY_OUTLOOK_FILTER', null, {
    criteria: "from:me to:me subject:'Daily Outlook'",
    action: "Apply label: Brief: Daily"
  });

  // Weekly Outlook Filter
  logFilterSpec('WEEKLY_OUTLOOK_FILTER', null, {
    criteria: "from:me to:me subject:'Weekly Outlook'",
    action: "Apply label: Brief: Weekly"
  });
}

/**
 * Logs a filter specification for documentation and manual creation.
 *
 * @param {string} filterType - Type of filter
 * @param {string|null} clientName - Client name if applicable
 * @param {Object} spec - Filter specification
 */
function logFilterSpec(filterType, clientName, spec) {
  const details = `Filter: ${filterType}` +
    (clientName ? ` for ${clientName}` : '') +
    ` | Criteria: ${spec.criteria} | Action: ${spec.action}`;

  Logger.log(details);
}

// ============================================================================
// GMAIL API FILTER MANAGEMENT (Requires Gmail Advanced Service)
// ============================================================================

/**
 * Creates a Gmail filter using the Gmail API.
 * Requires the Gmail Advanced Service to be enabled.
 *
 * @param {string} criteria - The filter criteria (Gmail search query)
 * @param {string} labelName - The label to apply
 * @returns {Object|null} The created filter or null if failed
 */
function createGmailApiFilter(criteria, labelName) {
  try {
    // Check if Gmail API is available
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      Logger.log('Gmail Advanced Service not enabled - skipping filter creation');
      return null;
    }

    // Get or create the label
    const label = createLabelIfNotExists(labelName);
    const labelId = getLabelId(labelName);

    if (!labelId) {
      Logger.log(`Could not get label ID for: ${labelName}`);
      return null;
    }

    // Check if filter already exists
    const existingFilters = listGmailFilters();
    for (const filter of existingFilters) {
      if (filter.criteria && filter.criteria.query === criteria) {
        Logger.log(`Filter already exists for criteria: ${criteria}`);
        return filter;
      }
    }

    // Create the filter
    const filter = Gmail.Users.Settings.Filters.create({
      criteria: {
        query: criteria
      },
      action: {
        addLabelIds: [labelId]
      }
    }, 'me');

    Logger.log(`Created Gmail filter: ${criteria} -> ${labelName}`);
    return filter;

  } catch (error) {
    Logger.log(`Failed to create Gmail filter: ${error.message}`);
    return null;
  }
}

/**
 * Gets the Gmail label ID for a label name.
 *
 * @param {string} labelName - The label name
 * @returns {string|null} The label ID or null
 */
function getLabelId(labelName) {
  try {
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      return null;
    }

    const response = Gmail.Users.Labels.list('me');
    const labels = response.labels || [];

    for (const label of labels) {
      if (label.name === labelName) {
        return label.id;
      }
    }
    return null;
  } catch (error) {
    Logger.log(`Failed to get label ID: ${error.message}`);
    return null;
  }
}

/**
 * Lists all existing Gmail filters.
 *
 * @returns {Object[]} Array of filter objects
 */
function listGmailFilters() {
  try {
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      return [];
    }

    const response = Gmail.Users.Settings.Filters.list('me');
    return response.filter || [];
  } catch (error) {
    Logger.log(`Failed to list filters: ${error.message}`);
    return [];
  }
}

/**
 * Deletes a Gmail filter by ID.
 *
 * @param {string} filterId - The filter ID to delete
 */
function deleteGmailFilter(filterId) {
  // Note: This requires Gmail Advanced Service
  /*
  try {
    Gmail.Users.Settings.Filters.remove('me', filterId);
    Logger.log(`Deleted filter: ${filterId}`);
  } catch (error) {
    Logger.log(`Failed to delete filter: ${error.message}`);
  }
  */
}

// ============================================================================
// MANUAL LABEL APPLICATION
// ============================================================================

/**
 * Manually applies labels to messages matching client criteria.
 * Useful for retroactively labeling existing emails.
 *
 * @param {Object} client - The client object
 * @param {number} maxResults - Maximum number of messages to process
 */
function retroactivelyLabelClientEmails(client, maxResults = 100) {
  const contacts = parseCommaSeparatedList(client.contact_emails);

  if (contacts.length === 0) {
    return;
  }

  const query = buildFromCriteria(contacts);
  const labelName = `Client: ${client.client_name}`;
  const label = createLabelIfNotExists(labelName);

  try {
    const threads = GmailApp.search(query, 0, maxResults);

    for (const thread of threads) {
      if (!thread.getLabels().some(l => l.getName() === labelName)) {
        thread.addLabel(label);
      }
    }

    Logger.log(`Retroactively labeled ${threads.length} threads for ${client.client_name}`);
  } catch (error) {
    Logger.log(`Failed to retroactively label: ${error.message}`);
  }
}

/**
 * Retroactively labels all client emails.
 * Should be run once after initial setup.
 */
function retroactivelyLabelAllClients() {
  const clients = getClientRegistry();

  for (const client of clients) {
    retroactivelyLabelClientEmails(client);
  }

  Logger.log('Completed retroactive labeling for all clients');
}

// ============================================================================
// FILTER UPDATE FUNCTIONS
// ============================================================================

/**
 * Extracts a stable subject pattern from the email subject template.
 * Used to create Gmail filters that match the template pattern.
 *
 * For example: "Team {client_name} - Meeting notes from "{meeting_title}" {date}"
 * becomes: "Meeting notes from"
 *
 * @returns {string} The stable portion of the subject for filter matching
 */
function getSubjectFilterPattern() {
  const props = PropertiesService.getScriptProperties();
  const template = props.getProperty('MEETING_SUBJECT_TEMPLATE')
    || 'Team {client_name} - Meeting notes from "{meeting_title}" {date}';

  // Extract text between placeholders - look for the most unique static part
  // Remove all placeholders to see what's left
  let pattern = template
    .replace(/{client_name}/g, '')
    .replace(/{meeting_title}/g, '')
    .replace(/{date}/g, '')
    .replace(/"/g, '') // Remove quotes
    .trim();

  // Find the longest static portion (usually "Meeting notes from" or similar)
  const parts = pattern.split(/\s+-\s+/).filter(p => p.trim().length > 5);

  if (parts.length > 0) {
    // Return the longest meaningful part
    return parts.reduce((a, b) => a.length > b.length ? a : b).trim();
  }

  // Fallback - if template is too dynamic, use a generic pattern
  return 'Meeting notes';
}

/**
 * Updates Gmail filters for meeting summaries when the subject template changes.
 * Deletes old filters and creates new ones based on the current template.
 */
function updateMeetingSummaryFilters() {
  Logger.log('Updating meeting summary filters...');

  const subjectPattern = getSubjectFilterPattern();
  Logger.log(`Using subject pattern for filters: "${subjectPattern}"`);

  // Get all clients with setup_complete
  const allClients = getClientRegistry();
  const clients = allClients.filter(client => client.setup_complete === true);

  if (clients.length === 0) {
    Logger.log('No clients with setup_complete found');
    return;
  }

  // Check if Gmail API is available
  if (typeof Gmail === 'undefined' || !Gmail.Users) {
    Logger.log('Gmail Advanced Service not enabled - logging filter specs only');
    // Just log what filters should be created
    for (const client of clients) {
      const contacts = parseCommaSeparatedList(client.contact_emails);
      if (contacts.length > 0) {
        const toCriteria = buildToCriteria(contacts);
        logFilterSpec('SUMMARY_FILTER_UPDATE', client.client_name, {
          criteria: `from:me subject:"${subjectPattern}" ${toCriteria}`,
          action: `Apply label: Client: ${client.client_name}/Meeting Summaries`
        });
      }
    }
    return;
  }

  // Delete old meeting summary filters and create new ones
  try {
    const existingFilters = listGmailFilters();

    // Find and delete old meeting summary filters
    for (const filter of existingFilters) {
      if (filter.criteria && filter.criteria.query) {
        const query = filter.criteria.query;
        // Match old summary filter patterns
        if (query.includes('from:me') &&
            (query.includes('subject:"Meeting Summary"') ||
             query.includes("subject:'Meeting Summary'") ||
             query.includes('Meeting notes'))) {
          try {
            Gmail.Users.Settings.Filters.remove('me', filter.id);
            Logger.log(`Deleted old filter: ${query}`);
          } catch (e) {
            Logger.log(`Failed to delete filter ${filter.id}: ${e.message}`);
          }
        }
      }
    }

    // Create new filters for each client
    for (const client of clients) {
      const contacts = parseCommaSeparatedList(client.contact_emails);
      if (contacts.length > 0) {
        const toCriteria = buildToCriteria(contacts);
        const criteria = `from:me subject:"${subjectPattern}" ${toCriteria}`;
        const labelName = `Client: ${client.client_name}/Meeting Summaries`;

        createGmailApiFilter(criteria, labelName);
        Logger.log(`Created new filter for ${client.client_name}: ${criteria}`);
      }
    }

    Logger.log('Meeting summary filters updated successfully');
    logProcessing('FILTER_UPDATE', null, `Updated meeting summary filters with pattern: "${subjectPattern}"`, 'success');

  } catch (error) {
    Logger.log(`Failed to update filters: ${error.message}`);
    logProcessing('FILTER_UPDATE', null, `Failed to update filters: ${error.message}`, 'error');
    throw error;
  }
}
