/**
 * FilterTracking.gs - Filter ID Storage and Tracking
 *
 * Stores Gmail filter IDs in Client_Registry to reliably identify system-created filters.
 */

// ============================================================================
// FILTER ID COLUMN MANAGEMENT
// ============================================================================

/**
 * Ensures filter ID columns exist in Client_Registry.
 * Adds them if missing (for backwards compatibility with older versions).
 *
 * @returns {boolean} True if columns were added, false if they already existed
 */
function ensureFilterIdColumnsExist() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID not set');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('Client_Registry');

  if (!sheet) {
    throw new Error('Client_Registry sheet not found');
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Check if filter ID columns already exist
  const hasFromFilter = headers.includes('from_filter_id');
  const hasToFilter = headers.includes('to_filter_id');
  const hasSummaryFilter = headers.includes('summary_filter_id');
  const hasAgendaFilter = headers.includes('agenda_filter_id');

  if (hasFromFilter && hasToFilter && hasSummaryFilter && hasAgendaFilter) {
    Logger.log('Filter ID columns already exist in Client_Registry');
    return false; // Already exist
  }

  // Add missing columns
  const newColumns = [];
  if (!hasFromFilter) newColumns.push('from_filter_id');
  if (!hasToFilter) newColumns.push('to_filter_id');
  if (!hasSummaryFilter) newColumns.push('summary_filter_id');
  if (!hasAgendaFilter) newColumns.push('agenda_filter_id');

  const startCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, startCol, 1, newColumns.length).setValues([newColumns]);
  sheet.getRange(1, startCol, 1, newColumns.length).setFontWeight('bold');

  logProcessing('FILTER_TRACKING', null, `Added ${newColumns.length} filter ID columns to Client_Registry: ${newColumns.join(', ')}`, 'success');
  return true;
}

/**
 * Stores filter IDs for a client in Client_Registry.
 *
 * @param {string} clientName - The client name
 * @param {Object} filterIds - Object with from_filter_id, to_filter_id, summary_filter_id, agenda_filter_id
 */
function storeFilterIds(clientName, filterIds) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    Logger.log('SPREADSHEET_ID not set');
    return;
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('Client_Registry');

  if (!sheet) {
    Logger.log('Client_Registry sheet not found');
    return;
  }

  // Ensure columns exist
  ensureFilterIdColumnsExist();

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const clientNameIdx = headers.indexOf('client_name');
  const fromFilterIdx = headers.indexOf('from_filter_id');
  const toFilterIdx = headers.indexOf('to_filter_id');
  const summaryFilterIdx = headers.indexOf('summary_filter_id');
  const agendaFilterIdx = headers.indexOf('agenda_filter_id');

  if (clientNameIdx === -1 || fromFilterIdx === -1 || toFilterIdx === -1 ||
      summaryFilterIdx === -1 || agendaFilterIdx === -1) {
    Logger.log('Required columns not found in Client_Registry');
    return;
  }

  // Find client row
  for (let i = 1; i < data.length; i++) {
    if (data[i][clientNameIdx] === clientName) {
      const rowNum = i + 1;

      // Update filter IDs
      if (filterIds.from_filter_id) {
        sheet.getRange(rowNum, fromFilterIdx + 1).setValue(filterIds.from_filter_id);
      }
      if (filterIds.to_filter_id) {
        sheet.getRange(rowNum, toFilterIdx + 1).setValue(filterIds.to_filter_id);
      }
      if (filterIds.summary_filter_id) {
        sheet.getRange(rowNum, summaryFilterIdx + 1).setValue(filterIds.summary_filter_id);
      }
      if (filterIds.agenda_filter_id) {
        sheet.getRange(rowNum, agendaFilterIdx + 1).setValue(filterIds.agenda_filter_id);
      }

      Logger.log(`Stored filter IDs for ${clientName}`);
      return;
    }
  }

  Logger.log(`Client not found: ${clientName}`);
}

/**
 * Gets all stored filter IDs from Client_Registry.
 *
 * @returns {Set} Set of all filter IDs stored in the system
 */
function getAllStoredFilterIds() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    return new Set();
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('Client_Registry');

  if (!sheet) {
    return new Set();
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const fromFilterIdx = headers.indexOf('from_filter_id');
  const toFilterIdx = headers.indexOf('to_filter_id');
  const summaryFilterIdx = headers.indexOf('summary_filter_id');
  const agendaFilterIdx = headers.indexOf('agenda_filter_id');

  const filterIds = new Set();

  // Skip if columns don't exist yet
  if (fromFilterIdx === -1) {
    return filterIds;
  }

  // Collect all non-empty filter IDs
  for (let i = 1; i < data.length; i++) {
    if (fromFilterIdx !== -1 && data[i][fromFilterIdx]) {
      filterIds.add(data[i][fromFilterIdx]);
    }
    if (toFilterIdx !== -1 && data[i][toFilterIdx]) {
      filterIds.add(data[i][toFilterIdx]);
    }
    if (summaryFilterIdx !== -1 && data[i][summaryFilterIdx]) {
      filterIds.add(data[i][summaryFilterIdx]);
    }
    if (agendaFilterIdx !== -1 && data[i][agendaFilterIdx]) {
      filterIds.add(data[i][agendaFilterIdx]);
    }
  }

  return filterIds;
}

// ============================================================================
// PATTERN-BASED FILTER IDENTIFICATION
// ============================================================================

/**
 * Checks if a filter matches system-created patterns using criteria matching.
 * Used as fallback when filter ID is not stored.
 *
 * @param {Object} filter - The Gmail filter object
 * @returns {boolean} True if filter matches system patterns
 */
function matchesSystemFilterPattern(filter) {
  if (!filter.criteria || !filter.criteria.query) {
    return false;
  }

  const criteria = filter.criteria.query;

  // Get all client names and their labels from registry
  const clients = getClientRegistry();

  for (const client of clients) {
    // Pattern 1: Agenda filters - from:me to:me subject:"Agenda: [ClientName]"
    const agendaPattern = getAgendaFilterPatternForClient(client.client_name);
    if (criteria.includes('from:me') &&
        criteria.includes('to:me') &&
        criteria.includes(`subject:"${agendaPattern}"`)) {
      return true;
    }

    // Pattern 2: Meeting summary filters - from:me subject:"Team [ClientName]"
    const summaryPattern = getSubjectFilterPatternForClient(client.client_name);
    if (criteria.includes('from:me') &&
        criteria.includes(`subject:"${summaryPattern}"`)) {
      return true;
    }

    // Pattern 3: FROM filters - from:contact@example.com
    const contacts = parseCommaSeparatedList(client.contact_emails);
    for (const contact of contacts) {
      if (criteria.includes(`from:${contact}`)) {
        return true;
      }
    }

    // Pattern 4: TO filters - to:contact@example.com
    for (const contact of contacts) {
      if (criteria.includes(`to:${contact}`)) {
        return true;
      }
    }
  }

  // Pattern 5: Briefing filters
  if (criteria.includes(`subject:'Daily Outlook'`) ||
      criteria.includes(`subject:"Daily Outlook"`)) {
    return true;
  }
  if (criteria.includes(`subject:'Weekly Outlook'`) ||
      criteria.includes(`subject:"Weekly Outlook"`)) {
    return true;
  }

  return false;
}

// ============================================================================
// COMPREHENSIVE BROKEN FILTER DETECTION
// ============================================================================

/**
 * Identifies broken filters using all three methods:
 * 1. Stored filter IDs (most reliable)
 * 2. Criteria pattern matching (fallback)
 * 3. User confirmation (always required)
 *
 * @param {boolean} autoFix - If true, deletes broken system filters after user confirmation
 * @returns {Object} Summary of broken filters found and fixed
 */
function findAndFixBrokenFiltersComprehensive(autoFix = false) {
  Logger.log('=== COMPREHENSIVE BROKEN FILTER CHECK ===\n');

  try {
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      Logger.log('❌ Gmail Advanced Service not enabled');
      return { broken: 0, fixed: 0, skipped: 0 };
    }

    // Get all filters
    const response = Gmail.Users.Settings.Filters.list('me');
    const allFilters = response.filter || [];

    // Get stored filter IDs (Method 1)
    const storedFilterIds = getAllStoredFilterIds();
    Logger.log(`Found ${storedFilterIds.size} stored filter IDs in Client_Registry`);

    const brokenSystemFilters = [];
    const brokenUserFilters = [];

    allFilters.forEach((filter) => {
      // Check if filter has no actions
      const hasAction = filter.action && (
        (filter.action.addLabelIds && filter.action.addLabelIds.length > 0) ||
        (filter.action.removeLabelIds && filter.action.removeLabelIds.length > 0) ||
        filter.action.forward
      );

      if (!hasAction) {
        const criteria = filter.criteria ? filter.criteria.query : 'N/A';
        const filterInfo = {
          id: filter.id,
          criteria: criteria,
          method: null
        };

        // Method 1: Check stored filter IDs
        if (storedFilterIds.has(filter.id)) {
          filterInfo.method = 'Stored Filter ID';
          brokenSystemFilters.push(filterInfo);
          Logger.log(`BROKEN SYSTEM FILTER (Method 1 - Stored ID):`);
          Logger.log(`  Criteria: ${criteria}`);
          Logger.log(`  ID: ${filter.id}`);
          Logger.log('');
        }
        // Method 2: Check criteria patterns
        else if (matchesSystemFilterPattern(filter)) {
          filterInfo.method = 'Criteria Pattern Match';
          brokenSystemFilters.push(filterInfo);
          Logger.log(`BROKEN SYSTEM FILTER (Method 2 - Pattern Match):`);
          Logger.log(`  Criteria: ${criteria}`);
          Logger.log(`  ID: ${filter.id}`);
          Logger.log('');
        }
        // Not system-created
        else {
          brokenUserFilters.push(filterInfo);
          Logger.log(`BROKEN USER FILTER (SKIPPED):`);
          Logger.log(`  Criteria: ${criteria}`);
          Logger.log(`  ID: ${filter.id}`);
          Logger.log('');
        }
      }
    });

    Logger.log(`\n=== SUMMARY ===`);
    Logger.log(`Total filters: ${allFilters.length}`);
    Logger.log(`Broken system filters: ${brokenSystemFilters.length}`);
    Logger.log(`Broken user filters (skipped): ${brokenUserFilters.length}`);

    if (brokenSystemFilters.length > 0) {
      Logger.log('\nBROKEN SYSTEM FILTERS BY DETECTION METHOD:');
      const byStoredId = brokenSystemFilters.filter(f => f.method === 'Stored Filter ID').length;
      const byPattern = brokenSystemFilters.filter(f => f.method === 'Criteria Pattern Match').length;
      Logger.log(`  Stored Filter ID: ${byStoredId}`);
      Logger.log(`  Criteria Pattern: ${byPattern}`);
    }

    if (brokenSystemFilters.length > 0 && autoFix) {
      Logger.log('\nDELETING BROKEN SYSTEM FILTERS...');
      let deletedCount = 0;

      for (const broken of brokenSystemFilters) {
        try {
          Gmail.Users.Settings.Filters.remove('me', broken.id);
          Logger.log(`✅ Deleted (${broken.method}): ${broken.criteria}`);
          deletedCount++;
        } catch (e) {
          Logger.log(`❌ Failed to delete ${broken.id}: ${e.message}`);
        }
      }

      Logger.log(`\nDeleted ${deletedCount} broken system filters`);
      Logger.log(`Skipped ${brokenUserFilters.length} user-created filters`);
      return {
        broken: brokenSystemFilters.length,
        fixed: deletedCount,
        skipped: brokenUserFilters.length
      };
    }

    return {
      broken: brokenSystemFilters.length,
      fixed: 0,
      skipped: brokenUserFilters.length
    };

  } catch (e) {
    Logger.log('❌ Error checking filters: ' + e.message);
    return { broken: 0, fixed: 0, skipped: 0 };
  }
}
