// ============================================================================
// FILTER RECONCILIATION - Links orphaned filters to Client_Registry
// ============================================================================

/**
 * Reconciles Client_Registry filter IDs with actual Gmail filters.
 * For each client, checks if filter IDs are tracked in sheet.
 * If missing, searches Gmail for matching filters and offers to link or create.
 *
 * @returns {Object} Reconciliation results
 */
function reconcileClientFilters() {
  Logger.log('=== FILTER RECONCILIATION STARTED ===\n');

  const results = {
    clients: [],
    totalMissing: 0,
    totalMatched: 0,
    totalUnmatched: 0
  };

  // Check Gmail API
  if (typeof Gmail === 'undefined' || !Gmail.Users) {
    Logger.log('❌ Gmail Advanced Service not enabled');
    results.error = 'Gmail Advanced Service not enabled';
    return results;
  }

  // Get all Gmail filters once
  const allGmailFilters = listGmailFilters();
  Logger.log(`Found ${allGmailFilters.length} total Gmail filters\n`);

  // Get clients
  const allClients = getClientRegistry();
  const clients = allClients.filter(client => client.setup_complete === true);

  if (clients.length === 0) {
    Logger.log('No clients with setup_complete=true');
    return results;
  }

  Logger.log(`Checking ${clients.length} clients...\n`);

  // Check each client
  for (const client of clients) {
    const clientResult = reconcileClientFilterIds(client, allGmailFilters);
    results.clients.push(clientResult);

    results.totalMissing += clientResult.missingCount;
    results.totalMatched += clientResult.matchedCount;
    results.totalUnmatched += clientResult.unmatchedCount;
  }

  Logger.log('\n=== RECONCILIATION SUMMARY ===');
  Logger.log(`Total clients checked: ${results.clients.length}`);
  Logger.log(`Total missing filter IDs: ${results.totalMissing}`);
  Logger.log(`Total matched filters: ${results.totalMatched}`);
  Logger.log(`Total unmatched (need creation): ${results.totalUnmatched}`);

  return results;
}

/**
 * Reconciles filter IDs for a single client.
 * Checks which filter IDs are missing from Client_Registry and searches Gmail for matches.
 *
 * @param {Object} client - Client object from registry
 * @param {Array} allGmailFilters - All Gmail filters
 * @returns {Object} Client reconciliation result
 */
function reconcileClientFilterIds(client, allGmailFilters) {
  Logger.log(`\n--- Client: ${client.client_name} ---`);

  const result = {
    clientName: client.client_name,
    missing: [],
    missingCount: 0,
    matchedCount: 0,
    unmatchedCount: 0
  };

  const contacts = parseCommaSeparatedList(client.contact_emails);
  if (contacts.length === 0) {
    Logger.log('⚠️  No contact emails - skipping');
    return result;
  }

  // Build expected label names
  const baseLabelName = client.gmail_label || `Client: ${client.client_name}`;
  const summaryLabelName = client.meeting_summaries_label || `${baseLabelName}/Meeting Summaries`;
  const agendaLabelName = client.meeting_agendas_label || `${baseLabelName}/Meeting Agendas`;

  // Check each filter type
  const filterTypes = [
    {
      type: 'from_filter_id',
      name: 'Incoming emails filter',
      idField: 'from_filter_id',
      criteria: buildFromCriteria(contacts),
      expectedLabel: baseLabelName
    },
    {
      type: 'to_filter_id',
      name: 'Outgoing emails filter',
      idField: 'to_filter_id',
      criteria: buildToCriteria(contacts),
      expectedLabel: baseLabelName
    },
    {
      type: 'summary_filter_id',
      name: 'Meeting summaries filter',
      idField: 'summary_filter_id',
      criteria: `from:me subject:"${getSubjectFilterPatternForClient(client.client_name)}" ${buildToCriteria(contacts)}`,
      expectedLabel: summaryLabelName
    },
    {
      type: 'agenda_filter_id',
      name: 'Meeting agendas filter',
      idField: 'agenda_filter_id',
      criteria: `from:me to:me subject:"${getAgendaFilterPatternForClient(client.client_name)}"`,
      expectedLabel: agendaLabelName
    }
  ];

  for (const filterType of filterTypes) {
    const currentId = client[filterType.idField];

    if (!currentId || currentId === '') {
      // Missing in Client_Registry - search for match
      Logger.log(`\n❌ MISSING: ${filterType.name}`);
      Logger.log(`   Expected criteria: ${filterType.criteria}`);
      Logger.log(`   Expected label: ${filterType.expectedLabel}`);

      const matches = findMatchingGmailFilters(allGmailFilters, filterType.criteria, filterType.expectedLabel);

      const missingFilter = {
        type: filterType.type,
        name: filterType.name,
        expectedCriteria: filterType.criteria,
        expectedLabel: filterType.expectedLabel,
        matches: matches
      };

      result.missing.push(missingFilter);
      result.missingCount++;

      if (matches.length > 0) {
        Logger.log(`   ✅ FOUND ${matches.length} potential match(es) in Gmail:`);
        matches.forEach((match, idx) => {
          Logger.log(`   \n   Match #${idx + 1}:`);
          Logger.log(`      Filter ID: ${match.id}`);
          Logger.log(`      Criteria: ${match.criteria.query || '(none)'}`);
          Logger.log(`      Labels: ${match.labelNames.join(', ') || '(none)'}`);
          Logger.log(`      Actions: Skip inbox=${match.action.skipInbox || false}, Archive=${!match.action.skipInbox}`);
        });
        result.matchedCount++;
      } else {
        Logger.log(`   ⚠️  NO MATCHES - will need to create new filter`);
        result.unmatchedCount++;
      }
    } else {
      Logger.log(`✅ ${filterType.name}: ${currentId}`);
    }
  }

  return result;
}

/**
 * Finds Gmail filters matching expected criteria and label.
 *
 * @param {Array} allFilters - All Gmail filters
 * @param {string} expectedCriteria - Expected filter criteria
 * @param {string} expectedLabel - Expected label name
 * @returns {Array} Matching filters with full details
 */
function findMatchingGmailFilters(allFilters, expectedCriteria, expectedLabel) {
  const matches = [];
  const expectedLabelId = getLabelId(expectedLabel);

  // Normalize criteria for comparison (remove extra spaces, quotes)
  const normalizedExpected = normalizeCriteria(expectedCriteria);

  for (const filter of allFilters) {
    const filterCriteria = filter.criteria.query || '';
    const normalizedFilter = normalizeCriteria(filterCriteria);

    // Check if criteria matches
    const criteriaMatch = normalizedFilter === normalizedExpected;

    // Check if label matches
    const filterLabels = filter.action.addLabelIds || [];
    const labelMatch = expectedLabelId && filterLabels.includes(expectedLabelId);

    if (criteriaMatch && labelMatch) {
      // Get label names for display
      const labelNames = filterLabels.map(id => getLabelNameById(id)).filter(n => n);

      matches.push({
        id: filter.id,
        criteria: filter.criteria,
        action: filter.action,
        labelIds: filterLabels,
        labelNames: labelNames
      });
    }
  }

  return matches;
}

/**
 * Normalizes filter criteria for comparison.
 * Removes extra whitespace, standardizes quote types.
 *
 * @param {string} criteria - Filter criteria
 * @returns {string} Normalized criteria
 */
function normalizeCriteria(criteria) {
  return criteria
    .toLowerCase()
    .replace(/[""]/g, '"')  // Standardize quotes
    .replace(/\s+/g, ' ')   // Collapse whitespace
    .trim();
}

/**
 * Gets label name by ID.
 *
 * @param {string} labelId - Label ID
 * @returns {string|null} Label name or null
 */
function getLabelNameById(labelId) {
  try {
    const label = Gmail.Users.Labels.get('me', labelId);
    return label.name;
  } catch (e) {
    return null;
  }
}

/**
 * Links a Gmail filter to a client's filter ID column.
 *
 * @param {string} clientName - Client name
 * @param {string} filterIdField - Field name (from_filter_id, to_filter_id, etc)
 * @param {string} filterId - Gmail filter ID
 * @returns {boolean} Success
 */
function linkFilterToClient(clientName, filterIdField, filterId) {
  Logger.log(`Linking filter ${filterId} to ${clientName}.${filterIdField}`);

  const filterIds = {};
  filterIds[filterIdField] = filterId;

  storeFilterIds(clientName, filterIds);
  return true;
}

/**
 * Creates missing filters for a client and updates Client_Registry.
 *
 * @param {string} clientName - Client name
 * @param {Array} missingFilters - Array of missing filter specs
 * @returns {Object} Creation results
 */
function createMissingFilters(clientName, missingFilters) {
  Logger.log(`\nCreating ${missingFilters.length} missing filters for ${clientName}...`);

  const results = {
    created: 0,
    failed: 0,
    filterIds: {}
  };

  for (const missing of missingFilters) {
    try {
      const filter = createGmailApiFilter(missing.expectedCriteria, missing.expectedLabel);

      if (filter && filter.id) {
        Logger.log(`✅ Created ${missing.name}: ${filter.id}`);
        results.filterIds[missing.type] = filter.id;
        results.created++;
      } else {
        Logger.log(`❌ Failed to create ${missing.name}`);
        results.failed++;
      }
    } catch (e) {
      Logger.log(`❌ Error creating ${missing.name}: ${e.message}`);
      results.failed++;
    }
  }

  // Store all created filter IDs
  if (Object.keys(results.filterIds).length > 0) {
    storeFilterIds(clientName, results.filterIds);
  }

  return results;
}
