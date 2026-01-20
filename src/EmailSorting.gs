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
  // Use stored label names from Client_Registry, or default to standard pattern
  const baseLabelName = client.gmail_label || `Client: ${client.client_name}`;
  const summaryLabelName = client.meeting_summaries_label || `${baseLabelName}/Meeting Summaries`;
  const agendaLabelName = client.meeting_agendas_label || `${baseLabelName}/Meeting Agendas`;

  // Create base client label
  createLabelIfNotExists(baseLabelName);

  // Create sub-labels
  createLabelIfNotExists(summaryLabelName);
  createLabelIfNotExists(agendaLabelName);

  Logger.log(`Synced labels for client: ${client.client_name} (base: ${baseLabelName})`);

  // Create filters (requires Gmail API Advanced Service)
  const contacts = parseCommaSeparatedList(client.contact_emails);

  if (contacts.length > 0) {
    // Filter for incoming emails from client contacts
    const fromCriteria = buildFromCriteria(contacts);
    if (fromCriteria) {
      createGmailApiFilter(fromCriteria, baseLabelName);
    }

    // Filter for outgoing emails to client contacts
    const toCriteria = buildToCriteria(contacts);
    if (toCriteria) {
      createGmailApiFilter(toCriteria, baseLabelName);
    }

    // Filter for sent meeting summaries to client (uses client name in subject)
    if (toCriteria) {
      const subjectPattern = getSubjectFilterPatternForClient(client.client_name);
      const summaryCriteria = `from:me subject:"${subjectPattern}" ${toCriteria}`;
      createGmailApiFilter(summaryCriteria, summaryLabelName);
    }
  }

  // Filter for self-sent agendas (uses client name in subject)
  const agendaPattern = getAgendaFilterPatternForClient(client.client_name);
  const agendaCriteria = `from:me to:me subject:"${agendaPattern}"`;
  createGmailApiFilter(agendaCriteria, agendaLabelName);

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

  // Use client-specific subject pattern (includes client name)
  const subjectPattern = getSubjectFilterPatternForClient(client.client_name);
  logFilterSpec('SUMMARY_FILTER', client.client_name, {
    criteria: `from:me subject:'${subjectPattern}' ${toCriteria}`,
    action: `Apply label: Client: ${client.client_name}/Meeting Summaries`
  });

  // Use client-specific agenda pattern (includes client name)
  const agendaPattern = getAgendaFilterPatternForClient(client.client_name);
  logFilterSpec('AGENDA_FILTER', client.client_name, {
    criteria: `from:me to:me subject:'${agendaPattern}'`,
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
 * ONLY creates filters for system-managed labels.
 *
 * @param {string} criteria - The filter criteria (Gmail search query)
 * @param {string} labelName - The label to apply (must be a system-managed label)
 * @returns {Object|null} The created filter or null if failed
 */
function createGmailApiFilter(criteria, labelName) {
  try {
    // Check if Gmail API is available
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      Logger.log('Gmail Advanced Service not enabled - skipping filter creation');
      return null;
    }

    // Safety check: Only create filters for system-managed labels
    const props = PropertiesService.getScriptProperties();
    const dailyLabel = props.getProperty('DAILY_BRIEFING_LABEL') || 'Brief: Daily';
    const weeklyLabel = props.getProperty('WEEKLY_BRIEFING_LABEL') || 'Brief: Weekly';

    const isSystemLabel = labelName.startsWith('Client: ') ||
                         labelName === dailyLabel ||
                         labelName === weeklyLabel ||
                         labelName === 'Brief: Daily' ||
                         labelName === 'Brief: Weekly';

    if (!isSystemLabel) {
      Logger.log(`SAFETY: Refusing to create filter for non-system label: ${labelName}`);
      return null;
    }

    // Get or create the label
    const label = createLabelIfNotExists(labelName);
    const labelId = getLabelId(labelName);

    if (!labelId) {
      Logger.log(`Could not get label ID for: ${labelName}`);
      return null;
    }

    // Check if identical filter already exists (same criteria AND same label)
    const existingFilters = listGmailFilters();
    for (const filter of existingFilters) {
      if (filter.criteria && filter.criteria.query === criteria) {
        // Check if it applies the same label
        const existingLabelIds = filter.action && filter.action.addLabelIds ? filter.action.addLabelIds : [];
        if (existingLabelIds.includes(labelId)) {
          Logger.log(`Filter already exists: ${criteria} -> ${labelName}`);
          return filter;
        }
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
 * SAFETY: Only deletes system-created filters.
 *
 * @param {string} filterId - The filter ID to delete
 * @returns {boolean} True if deleted, false if skipped or failed
 */
function deleteGmailFilter(filterId) {
  try {
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      Logger.log('Gmail Advanced Service not enabled');
      return false;
    }

    // Get the filter details first
    const existingFilters = listGmailFilters();
    const filter = existingFilters.find(f => f.id === filterId);

    if (!filter) {
      Logger.log(`Filter not found: ${filterId}`);
      return false;
    }

    // CRITICAL SAFETY CHECK: Only delete system-created filters
    if (!isSystemCreatedFilter(filter)) {
      Logger.log(`SAFETY: Refusing to delete user-created filter: ${filterId}`);
      logProcessing('FILTER_DELETE', null, `Blocked attempt to delete user filter: ${filterId}`, 'warning');
      return false;
    }

    // Safe to delete - it's a system filter
    Gmail.Users.Settings.Filters.remove('me', filterId);
    Logger.log(`Deleted system filter: ${filterId}`);
    return true;

  } catch (error) {
    Logger.log(`Failed to delete filter: ${error.message}`);
    return false;
  }
}

/**
 * Lists all system-managed filters.
 * Useful for debugging and verification.
 *
 * @returns {Object[]} Array of system-managed filter objects
 */
function listSystemManagedFilters() {
  try {
    const allFilters = listGmailFilters();
    const systemFilters = [];

    for (const filter of allFilters) {
      if (isSystemCreatedFilter(filter)) {
        systemFilters.push({
          id: filter.id,
          criteria: filter.criteria ? filter.criteria.query : 'N/A',
          labels: getLabelNamesFromFilter(filter)
        });
      }
    }

    Logger.log(`Found ${systemFilters.length} system-managed filters out of ${allFilters.length} total`);
    return systemFilters;

  } catch (error) {
    Logger.log(`Error listing system filters: ${error.message}`);
    return [];
  }
}

/**
 * Gets label names from a filter object.
 *
 * @param {Object} filter - The filter object
 * @returns {string[]} Array of label names
 */
function getLabelNamesFromFilter(filter) {
  if (!filter.action || !filter.action.addLabelIds) {
    return [];
  }

  const labelNames = [];
  for (const labelId of filter.action.addLabelIds) {
    const name = getLabelNameById(labelId);
    if (name) {
      labelNames.push(name);
    }
  }

  return labelNames;
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
 * Extracts the subject prefix pattern for a specific client.
 * Used to create Gmail filters that match emails for that client.
 *
 * For template: "Team {client_name} - Meeting notes from "{meeting_title}" {date}"
 * With client "ACME Corp", returns: "Team ACME Corp"
 *
 * @param {string} clientName - The client name to insert
 * @returns {string} The subject pattern for this client
 */
function getSubjectFilterPatternForClient(clientName) {
  const props = PropertiesService.getScriptProperties();
  const template = props.getProperty('MEETING_SUBJECT_TEMPLATE')
    || 'Team {client_name} - Meeting notes from "{meeting_title}" {date}';

  // Find where {client_name} appears and extract text before/after up to the next placeholder
  // For "Team {client_name} - Meeting notes..." we want "Team ClientName"

  // Split template at {client_name}
  const parts = template.split('{client_name}');
  if (parts.length < 2) {
    // No {client_name} placeholder - just use client name directly
    return clientName;
  }

  // Get prefix before {client_name} (e.g., "Team ")
  const prefix = parts[0];

  // Get text after {client_name} up to next placeholder or punctuation
  let suffix = parts[1];
  // Cut off at next placeholder or after first few words
  const nextPlaceholder = suffix.search(/\{[^}]+\}/);
  if (nextPlaceholder > 0) {
    suffix = suffix.substring(0, nextPlaceholder);
  }
  // Also cut at common separators to keep pattern short
  const separatorMatch = suffix.match(/^(\s*[-–—:]\s*)/);
  suffix = separatorMatch ? '' : suffix.split(/[-–—]/)[0];

  const pattern = (prefix + clientName + suffix).trim();
  return pattern;
}

/**
 * Extracts the agenda subject pattern for a specific client.
 * Used to create Gmail filters that match agenda emails for that client.
 *
 * For template: "Agenda: {client_name} - {meeting_title} ({date})"
 * With client "ACME Corp", returns: "Agenda: ACME Corp"
 *
 * @param {string} clientName - The client name to insert
 * @returns {string} The subject pattern for this client's agendas
 */
function getAgendaFilterPatternForClient(clientName) {
  const props = PropertiesService.getScriptProperties();
  const template = props.getProperty('AGENDA_SUBJECT_TEMPLATE')
    || 'Agenda: {client_name} - {meeting_title} ({date})';

  // Split template at {client_name}
  const parts = template.split('{client_name}');
  if (parts.length < 2) {
    // No {client_name} placeholder - just use "Agenda:" prefix
    return 'Agenda:';
  }

  // Get prefix before {client_name} (e.g., "Agenda: ")
  const prefix = parts[0];

  // Get text after {client_name} up to next placeholder
  let suffix = parts[1];
  const nextPlaceholder = suffix.search(/\{[^}]+\}/);
  if (nextPlaceholder > 0) {
    suffix = suffix.substring(0, nextPlaceholder);
  }
  // Cut at common separators
  const separatorMatch = suffix.match(/^(\s*[-–—:]\s*)/);
  suffix = separatorMatch ? '' : suffix.split(/[-–—]/)[0];

  const pattern = (prefix + clientName + suffix).trim();
  return pattern;
}

/**
 * Checks if a filter was created by this system.
 * Only returns true if the filter applies labels managed by this system.
 *
 * System-managed labels:
 * - Client: *
 * - Client: [name]/Meeting Summaries
 * - Client: [name]/Meeting Agendas
 * - Brief: Daily (or custom daily label)
 * - Brief: Weekly (or custom weekly label)
 *
 * @param {Object} filter - The Gmail filter object
 * @returns {boolean} True if this is a system-created filter
 */
function isSystemCreatedFilter(filter) {
  if (!filter.action || !filter.action.addLabelIds) {
    return false;
  }

  try {
    // Get all label IDs and names
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      return false;
    }

    const response = Gmail.Users.Labels.list('me');
    const labels = response.labels || [];

    // Map label IDs to names
    const labelMap = {};
    for (const label of labels) {
      labelMap[label.id] = label.name;
    }

    // Check if ANY of the filter's labels are system-managed
    for (const labelId of filter.action.addLabelIds) {
      const labelName = labelMap[labelId];
      if (!labelName) continue;

      // Get custom briefing labels from settings
      const props = PropertiesService.getScriptProperties();
      const dailyLabel = props.getProperty('DAILY_BRIEFING_LABEL') || 'Brief: Daily';
      const weeklyLabel = props.getProperty('WEEKLY_BRIEFING_LABEL') || 'Brief: Weekly';

      // System-managed label patterns
      if (labelName.startsWith('Client: ') ||
          labelName === dailyLabel ||
          labelName === weeklyLabel ||
          labelName === 'Brief: Daily' ||
          labelName === 'Brief: Weekly') {
        return true;
      }
    }

    return false;
  } catch (error) {
    Logger.log(`Error checking if filter is system-created: ${error.message}`);
    return false;
  }
}

/**
 * Updates Gmail filters for meeting summaries when the subject template changes.
 * ONLY manages filters that were created by this system.
 * Never touches user-created filters.
 */
function updateMeetingSummaryFilters() {
  logProcessing('FILTER_UPDATE', null, 'Starting filter update...', 'info');

  // Get all clients with setup_complete
  const allClients = getClientRegistry();
  const clients = allClients.filter(client => client.setup_complete === true);

  if (clients.length === 0) {
    logProcessing('FILTER_UPDATE', null, 'No clients with setup_complete found', 'warning');
    return;
  }

  logProcessing('FILTER_UPDATE', null, `Processing ${clients.length} clients`, 'info');

  // Check if Gmail API is available
  if (typeof Gmail === 'undefined' || !Gmail.Users) {
    logProcessing('FILTER_UPDATE', null, 'Gmail Advanced Service not enabled - cannot update filters programmatically', 'error');
    return;
  }

  // Delete old meeting summary filters and create new ones
  try {
    const existingFilters = listGmailFilters();
    let deletedCount = 0;
    let createdCount = 0;
    let skippedUserFilters = 0;

    // Find and delete ONLY system-created meeting summary filters
    for (const filter of existingFilters) {
      // CRITICAL SAFETY CHECK: Only process system-created filters
      if (!isSystemCreatedFilter(filter)) {
        skippedUserFilters++;
        continue; // Skip user-created filters
      }

      if (filter.criteria && filter.criteria.query) {
        const query = filter.criteria.query;

        // Only delete system meeting summary filters
        // (applies Client: */Meeting Summaries label)
        const labelIds = filter.action.addLabelIds || [];
        let isMeetingSummaryFilter = false;

        for (const labelId of labelIds) {
          const labelName = getLabelNameById(labelId);
          if (labelName && labelName.includes('/Meeting Summaries')) {
            isMeetingSummaryFilter = true;
            break;
          }
        }

        if (isMeetingSummaryFilter) {
          try {
            Gmail.Users.Settings.Filters.remove('me', filter.id);
            logProcessing('FILTER_UPDATE', null, `Deleted system filter: ${query}`, 'info');
            deletedCount++;
          } catch (e) {
            logProcessing('FILTER_UPDATE', null, `Failed to delete filter: ${e.message}`, 'error');
          }
        }
      }
    }

    // Create new filters for each client using client-specific subject pattern
    for (const client of clients) {
      const contacts = parseCommaSeparatedList(client.contact_emails);
      if (contacts.length > 0) {
        const toCriteria = buildToCriteria(contacts);
        const subjectPattern = getSubjectFilterPatternForClient(client.client_name);
        const criteria = `from:me subject:"${subjectPattern}" ${toCriteria}`;
        const labelName = `Client: ${client.client_name}/Meeting Summaries`;

        const result = createGmailApiFilter(criteria, labelName);
        if (result) {
          logProcessing('FILTER_UPDATE', client.client_name, `Created filter: ${criteria}`, 'success');
          createdCount++;
        } else {
          logProcessing('FILTER_UPDATE', client.client_name, `Filter may already exist: ${criteria}`, 'warning');
        }
      } else {
        logProcessing('FILTER_UPDATE', client.client_name, 'No contact emails - skipping', 'warning');
      }
    }

    logProcessing('FILTER_UPDATE', null, `Filter update complete. Deleted: ${deletedCount}, Created: ${createdCount}, User filters skipped: ${skippedUserFilters}`, 'success');

  } catch (error) {
    logProcessing('FILTER_UPDATE', null, `Failed to update filters: ${error.message}`, 'error');
    throw error;
  }
}

/**
 * Gets the label name by label ID.
 *
 * @param {string} labelId - The label ID
 * @returns {string|null} The label name or null
 */
function getLabelNameById(labelId) {
  try {
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      return null;
    }

    const response = Gmail.Users.Labels.list('me');
    const labels = response.labels || [];

    for (const label of labels) {
      if (label.id === labelId) {
        return label.name;
      }
    }
    return null;
  } catch (error) {
    Logger.log(`Failed to get label name: ${error.message}`);
    return null;
  }
}
