/**
 * VerifyFilters.gs - Diagnostic Functions for Gmail Filter Verification
 *
 * Run these functions to check if your filters are working properly.
 */

/**
 * Main verification function - run this to check filter status.
 * Outputs results to Logger and returns a summary object.
 */
function verifyFiltersAndLabels() {
  Logger.log('=== GMAIL FILTER & LABEL VERIFICATION ===\n');

  const results = {
    gmailApiEnabled: false,
    labelsCreated: [],
    filtersCreated: [],
    clients: [],
    issues: []
  };

  // 1. Check if Gmail Advanced Service is enabled
  Logger.log('1. Checking Gmail Advanced Service status...');
  try {
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      Logger.log('❌ Gmail Advanced Service is NOT enabled');
      results.gmailApiEnabled = false;
      results.issues.push('Gmail Advanced Service not enabled - filters cannot be created programmatically');
      Logger.log('\nTO FIX:');
      Logger.log('  1. In Apps Script editor, click on + next to Services');
      Logger.log('  2. Find "Gmail API" and select it');
      Logger.log('  3. Click "Add"');
      Logger.log('  4. Re-run syncLabelsAndFilters()');
    } else {
      Logger.log('✅ Gmail Advanced Service is enabled');
      results.gmailApiEnabled = true;
    }
  } catch (e) {
    Logger.log('❌ Error checking Gmail API: ' + e.message);
    results.gmailApiEnabled = false;
  }

  Logger.log('\n2. Checking created labels...');

  // 2. Check if labels exist
  try {
    const allLabels = GmailApp.getUserLabels();
    const clientLabels = allLabels.filter(label => label.getName().startsWith('Client:'));
    const briefLabels = allLabels.filter(label => label.getName().startsWith('Brief:'));

    Logger.log(`Found ${clientLabels.length} client labels:`);
    clientLabels.forEach(label => {
      Logger.log(`  - ${label.getName()}`);
      results.labelsCreated.push(label.getName());
    });

    Logger.log(`\nFound ${briefLabels.length} briefing labels:`);
    briefLabels.forEach(label => {
      Logger.log(`  - ${label.getName()}`);
      results.labelsCreated.push(label.getName());
    });

    if (clientLabels.length === 0) {
      results.issues.push('No client labels found - run syncLabelsAndFilters()');
    }
  } catch (e) {
    Logger.log('❌ Error checking labels: ' + e.message);
  }

  Logger.log('\n3. Checking created filters...');

  // 3. Check if filters exist
  if (results.gmailApiEnabled) {
    try {
      const allFilters = listGmailFilters();
      const systemFilters = [];

      for (const filter of allFilters) {
        if (isSystemCreatedFilter(filter)) {
          const labelNames = getLabelNamesFromFilter(filter);
          const criteria = filter.criteria ? filter.criteria.query : 'N/A';
          systemFilters.push({ criteria, labels: labelNames });

          Logger.log(`  Criteria: ${criteria}`);
          Logger.log(`  Labels: ${labelNames.join(', ')}`);
          Logger.log('');

          results.filtersCreated.push({ criteria, labels: labelNames });
        }
      }

      Logger.log(`Found ${systemFilters.length} system-managed filters out of ${allFilters.length} total`);

      if (systemFilters.length === 0 && clientLabels.length > 0) {
        results.issues.push('Labels exist but no filters created - run syncLabelsAndFilters() again');
      }
    } catch (e) {
      Logger.log('❌ Error checking filters: ' + e.message);
    }
  } else {
    Logger.log('⚠️  Skipping filter check (Gmail API not enabled)');
  }

  Logger.log('\n4. Checking Client Registry...');

  // 4. Check clients
  try {
    const allClients = getClientRegistry();
    const setupCompleteClients = allClients.filter(c => c.setup_complete === true);

    Logger.log(`Total clients: ${allClients.length}`);
    Logger.log(`Clients with setup_complete=true: ${setupCompleteClients.length}`);

    setupCompleteClients.forEach(client => {
      const contacts = parseCommaSeparatedList(client.contact_emails);
      Logger.log(`  - ${client.client_name}: ${contacts.length} contact email(s)`);

      if (contacts.length === 0) {
        results.issues.push(`${client.client_name} has no contact emails - filters cannot be created`);
      }

      results.clients.push({
        name: client.client_name,
        contacts: contacts.length
      });
    });

    if (setupCompleteClients.length === 0) {
      results.issues.push('No clients with setup_complete=true - set clients to setup_complete in Client_Registry');
    }
  } catch (e) {
    Logger.log('❌ Error checking clients: ' + e.message);
  }

  // 5. Summary
  Logger.log('\n=== SUMMARY ===');
  if (results.issues.length === 0) {
    Logger.log('✅ Everything looks good!');
  } else {
    Logger.log('⚠️  Issues found:');
    results.issues.forEach(issue => {
      Logger.log(`  - ${issue}`);
    });
  }

  return results;
}

/**
 * Quick test to see if filters are working.
 * Searches Gmail for emails matching filter criteria.
 */
function testFilterCriteria(clientName) {
  if (!clientName) {
    Logger.log('Usage: testFilterCriteria("Client Name")');
    return;
  }

  Logger.log(`Testing filter criteria for: ${clientName}\n`);

  // Get client
  const clients = getClientRegistry();
  const client = clients.find(c => c.client_name === clientName);

  if (!client) {
    Logger.log('❌ Client not found: ' + clientName);
    return;
  }

  const contacts = parseCommaSeparatedList(client.contact_emails);

  if (contacts.length === 0) {
    Logger.log('❌ No contact emails for this client');
    return;
  }

  Logger.log('Contact emails: ' + contacts.join(', '));

  // Test FROM criteria
  const fromCriteria = buildFromCriteria(contacts);
  Logger.log(`\nFROM criteria: ${fromCriteria}`);

  try {
    const fromThreads = GmailApp.search(fromCriteria, 0, 5);
    Logger.log(`  Found ${fromThreads.length} emails matching FROM criteria`);
    fromThreads.forEach(thread => {
      Logger.log(`    - ${thread.getFirstMessageSubject()}`);
    });
  } catch (e) {
    Logger.log(`  ❌ Error: ${e.message}`);
  }

  // Test TO criteria
  const toCriteria = buildToCriteria(contacts);
  Logger.log(`\nTO criteria: ${toCriteria}`);

  try {
    const toThreads = GmailApp.search(toCriteria, 0, 5);
    Logger.log(`  Found ${toThreads.length} emails matching TO criteria`);
    toThreads.forEach(thread => {
      Logger.log(`    - ${thread.getFirstMessageSubject()}`);
    });
  } catch (e) {
    Logger.log(`  ❌ Error: ${e.message}`);
  }

  // Test meeting summary criteria
  const subjectPattern = getSubjectFilterPatternForClient(clientName);
  const summaryCriteria = `from:me subject:"${subjectPattern}" ${toCriteria}`;
  Logger.log(`\nMEETING SUMMARY criteria: ${summaryCriteria}`);

  try {
    const summaryThreads = GmailApp.search(summaryCriteria, 0, 5);
    Logger.log(`  Found ${summaryThreads.length} emails matching SUMMARY criteria`);
    summaryThreads.forEach(thread => {
      Logger.log(`    - ${thread.getFirstMessageSubject()}`);
    });
  } catch (e) {
    Logger.log(`  ❌ Error: ${e.message}`);
  }

  // Test agenda criteria
  const agendaPattern = getAgendaFilterPatternForClient(clientName);
  const agendaCriteria = `from:me to:me subject:"${agendaPattern}"`;
  Logger.log(`\nAGENDA criteria: ${agendaCriteria}`);

  try {
    const agendaThreads = GmailApp.search(agendaCriteria, 0, 5);
    Logger.log(`  Found ${agendaThreads.length} emails matching AGENDA criteria`);
    agendaThreads.forEach(thread => {
      Logger.log(`    - ${thread.getFirstMessageSubject()}`);
    });
  } catch (e) {
    Logger.log(`  ❌ Error: ${e.message}`);
  }
}

/**
 * Shows all Gmail filters in your account (both user-created and system-created).
 */
function showAllGmailFilters() {
  Logger.log('=== ALL GMAIL FILTERS ===\n');

  try {
    if (typeof Gmail === 'undefined' || !Gmail.Users) {
      Logger.log('❌ Gmail Advanced Service not enabled');
      Logger.log('Cannot list filters programmatically without Gmail API');
      Logger.log('\nYou can manually check filters in Gmail:');
      Logger.log('  1. Open Gmail');
      Logger.log('  2. Click Settings (gear icon) > See all settings');
      Logger.log('  3. Go to "Filters and Blocked Addresses" tab');
      return;
    }

    const response = Gmail.Users.Settings.Filters.list('me');
    const allFilters = response.filter || [];

    Logger.log(`Total filters in account: ${allFilters.length}\n`);

    let systemCount = 0;
    let userCount = 0;

    allFilters.forEach((filter, index) => {
      const isSystem = isSystemCreatedFilter(filter);
      const criteria = filter.criteria ? filter.criteria.query : 'N/A';
      const labelNames = getLabelNamesFromFilter(filter);

      Logger.log(`Filter ${index + 1} ${isSystem ? '(SYSTEM)' : '(USER)'}`);
      Logger.log(`  Criteria: ${criteria}`);
      Logger.log(`  Labels: ${labelNames.join(', ') || 'None'}`);
      Logger.log('');

      if (isSystem) {
        systemCount++;
      } else {
        userCount++;
      }
    });

    Logger.log('=== SUMMARY ===');
    Logger.log(`System-managed filters: ${systemCount}`);
    Logger.log(`User-created filters: ${userCount}`);
    Logger.log(`Total filters: ${allFilters.length}`);

  } catch (e) {
    Logger.log('❌ Error listing filters: ' + e.message);
  }
}
