/**
 * OutlookReports.gs - Daily and Weekly Outlook Logic
 *
 * This module generates daily and weekly outlook reports:
 * - Daily outlook: Runs at 7:00 AM every day
 * - Weekly outlook: Runs at 7:00 AM every Monday
 *
 * Reports include:
 * - Meetings organized by client
 * - Todoist tasks due today/this week
 * - Schedule conflicts
 * - Missing agendas
 * - Overdue prerequisite tasks
 */

// ============================================================================
// DAILY OUTLOOK
// ============================================================================

/**
 * Generates and sends the daily outlook report.
 * Called at 7:00 AM daily.
 * Uses Claude AI to generate strategic briefing from all data.
 */
function generateDailyOutlook() {
  Logger.log('Generating daily outlook...');

  // Run auto-mark-read before generating report (if enabled)
  autoMarkOldEmailsAsRead();

  const today = new Date();
  const reportData = compileDailyData(today);

  // Generate HTML report using AI
  let htmlReport;
  const aiGenerated = generateDailyOutlookWithClaude(reportData, today);

  if (aiGenerated) {
    htmlReport = aiGenerated;
    Logger.log('Daily outlook generated with AI');
  } else {
    // Fallback to template-based if AI fails
    htmlReport = formatDailyOutlookHtml(reportData, today);
    Logger.log('Daily outlook generated with fallback template (AI unavailable)');
  }

  // Send email
  const props = PropertiesService.getScriptProperties();
  const dailyLabel = props.getProperty('DAILY_BRIEFING_LABEL') || 'Brief: Daily';
  const subject = `Daily Outlook - ${formatDate(today)}`;
  sendOutlookEmail(subject, htmlReport, dailyLabel);

  Logger.log('Daily outlook sent');
}

/**
 * Generates daily outlook HTML using Claude AI.
 * Passes ALL data through the AI prompt for strategic analysis.
 *
 * @param {Object} data - The compiled report data
 * @param {Date} date - The report date
 * @returns {string|null} AI-generated HTML or null if failed
 */
function generateDailyOutlookWithClaude(data, date) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    Logger.log('Claude API key not configured - falling back to template');
    return null;
  }

  const prompt = buildDailyOutlookPrompt(data, date);

  try {
    const url = 'https://api.anthropic.com/v1/messages';
    const model = getModelForPrompt('DAILY_BRIEFING_CLAUDE_PROMPT');

    const payload = {
      model: model,
      max_tokens: 4000,
      messages: [
        {
          role: 'user',
          content: prompt
        }
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
      Logger.log(`Claude API error for daily outlook: ${responseCode} - ${response.getContentText()}`);
      return null;
    }

    // Parse response with explicit UTF-8 handling
    const responseText = response.getContentText('UTF-8');
    const result = JSON.parse(responseText);

    if (result.content && result.content.length > 0) {
      let content = result.content[0].text;

      // Ensure content is treated as UTF-8
      // If Claude returns full HTML with charset, preserve it
      // If not, we'll add it in sendOutlookEmail
      return content;
    }

    return null;

  } catch (error) {
    Logger.log(`Failed to generate daily outlook with Claude: ${error.message}`);
    return null;
  }
}

/**
 * Builds the prompt for Claude daily outlook generation.
 * Includes ALL available data: calendar, tasks, unread emails, drafts, conflicts.
 *
 * @param {Object} data - The compiled report data
 * @param {Date} date - The report date
 * @returns {string} The formatted prompt
 */
function buildDailyOutlookPrompt(data, date) {
  // Build calendar section
  let calendarSection = '';
  if (data.meetings.length > 0) {
    calendarSection = 'Today\'s Calendar:\n';
    for (const meeting of data.meetings.sort((a, b) => a.start - b.start)) {
      calendarSection += `- ${formatTime(meeting.start)} - ${formatTime(meeting.end)}: ${meeting.title}`;
      if (meeting.clientName !== 'Unassigned') {
        calendarSection += ` (${meeting.clientName})`;
      }
      if (!meeting.hasAgenda) {
        calendarSection += ' [NO AGENDA]';
      }
      calendarSection += '\n';
    }
  } else {
    calendarSection = 'Today\'s Calendar:\nNo meetings scheduled.\n';
  }

  // Build conflicts section
  let conflictsSection = '';
  if (data.conflicts.length > 0) {
    conflictsSection = '\nSchedule Conflicts:\n';
    for (const conflict of data.conflicts) {
      conflictsSection += `- "${conflict.event1.title}" overlaps with "${conflict.event2.title}"\n`;
    }
  }

  // Build tasks section
  let tasksSection = '';
  if (data.tasks.length > 0) {
    tasksSection = 'Tasks Due Today or Overdue:\n';
    for (const task of data.tasks) {
      const overdueFlag = task.isOverdue ? ' [OVERDUE]' : '';
      tasksSection += `- [${task.client.client_name}] ${task.content}${overdueFlag}\n`;
    }
  } else {
    tasksSection = 'Tasks Due Today or Overdue:\nNo tasks due today.\n';
  }

  // Build unread emails section
  let unreadSection = '';
  if (data.includeUnreadEmails && data.unreadEmails.totalUnread > 0) {
    unreadSection = `\nUnread Emails (${data.unreadEmails.totalUnread} total):\n`;

    if (data.unreadEmails.recentUnread.length > 0) {
      unreadSection += 'Recent (last 24 hours):\n';
      for (const email of data.unreadEmails.recentUnread.slice(0, 10)) {
        unreadSection += `- From: ${email.from}, Subject: ${email.subject}\n`;
      }
    }

    if (data.unreadEmails.olderBacklog.length > 0) {
      unreadSection += `Older backlog (${data.unreadEmails.olderBacklog.length} emails):\n`;
      for (const email of data.unreadEmails.olderBacklog.slice(0, 5)) {
        unreadSection += `- From: ${email.from}, Subject: ${email.subject} (${email.date})\n`;
      }
    }
  }

  // Get the prompt template
  const template = getPrompt('DAILY_BRIEFING_CLAUDE_PROMPT');

  // Apply variables to template
  return applyTemplate(template, {
    current_date: formatDate(date),
    todays_calendar: calendarSection + conflictsSection,
    todays_tasks: tasksSection,
    urgent_emails: unreadSection || 'No unread emails to report.'
  });
}

/**
 * Compiles data for the daily outlook.
 *
 * @param {Date} date - The date for the outlook
 * @returns {Object} Compiled report data
 */
function compileDailyData(date) {
  const data = {
    meetings: [],
    tasks: [],
    clientSummaries: {},
    conflicts: [],
    missingAgendas: [],
    overdueTasks: [],
    unreadEmails: { recentUnread: [], olderBacklog: [], totalUnread: 0 },
    includeUnreadEmails: false
  };

  // Check if unread emails should be included (controlled by settings)
  const props = PropertiesService.getScriptProperties();
  const includeUnreadEmails = props.getProperty('INCLUDE_UNREAD_EMAILS') === 'true';
  data.includeUnreadEmails = includeUnreadEmails;

  if (includeUnreadEmails) {
    // Fetch unread emails from last 1 day for daily report
    data.unreadEmails = fetchUnreadEmails(1);
  }

  // Get today's events
  const calendar = CalendarApp.getDefaultCalendar();
  const startOfDay = new Date(date);
  startOfDay.setHours(0, 0, 0, 0);
  const endOfDay = new Date(date);
  endOfDay.setHours(23, 59, 59, 999);

  const events = calendar.getEvents(startOfDay, endOfDay);

  // Process each event
  for (const event of events) {
    if (event.isAllDayEvent()) continue;

    const eventInfo = {
      title: event.getTitle(),
      start: event.getStartTime(),
      end: event.getEndTime(),
      client: null,
      clientName: 'Unassigned',
      hasAgenda: false
    };

    // Identify client
    const client = identifyClientFromCalendarEvent(event);
    if (client) {
      eventInfo.client = client;
      eventInfo.clientName = client.client_name;

      // Check if agenda exists
      eventInfo.hasAgenda = isAgendaGenerated(event.getId());

      if (!eventInfo.hasAgenda) {
        data.missingAgendas.push(eventInfo);
      }

      // Add to client summary
      if (!data.clientSummaries[client.client_name]) {
        data.clientSummaries[client.client_name] = {
          client: client,
          meetings: [],
          tasks: []
        };
      }
      data.clientSummaries[client.client_name].meetings.push(eventInfo);
    }

    data.meetings.push(eventInfo);
  }

  // Detect conflicts
  data.conflicts = detectScheduleConflicts(data.meetings);

  // Get all clients and fetch their tasks
  const clients = getClientRegistry();
  for (const client of clients) {
    if (client.todoist_project_id) {
      const tasks = fetchTodoistTasksDueToday(client.todoist_project_id);

      for (const task of tasks) {
        const taskInfo = {
          content: task.content,
          due: task.due ? task.due.date : null,
          client: client,
          isOverdue: isTaskOverdue(task, date)
        };

        data.tasks.push(taskInfo);

        if (taskInfo.isOverdue) {
          data.overdueTasks.push(taskInfo);
        }

        // Add to client summary
        if (!data.clientSummaries[client.client_name]) {
          data.clientSummaries[client.client_name] = {
            client: client,
            meetings: [],
            tasks: []
          };
        }
        data.clientSummaries[client.client_name].tasks.push(taskInfo);
      }
    }
  }

  return data;
}

/**
 * Formats the daily outlook as HTML.
 *
 * @param {Object} data - The compiled report data
 * @param {Date} date - The report date
 * @returns {string} HTML formatted report
 */
function formatDailyOutlookHtml(data, date) {
  let html = `<html><body style="font-family: Arial, sans-serif;">`;

  // Get header template from sheet
  const headerTemplate = getPrompt('DAILY_OUTLOOK_HEADER');
  html += applyTemplate(headerTemplate, { date: formatDate(date) });

  // Alerts section
  if (data.conflicts.length > 0 || data.missingAgendas.length > 0 || data.overdueTasks.length > 0) {
    let alertsContent = '';
    const alertsTemplate = getPrompt('DAILY_OUTLOOK_ALERTS');

    if (data.conflicts.length > 0) {
      alertsContent += `<h3>Schedule Conflicts</h3><ul>`;
      for (const conflict of data.conflicts) {
        alertsContent += `<li><strong>${conflict.event1.title}</strong> overlaps with <strong>${conflict.event2.title}</strong></li>`;
      }
      alertsContent += `</ul>`;
    }

    if (data.missingAgendas.length > 0) {
      alertsContent += `<h3>Missing Agendas</h3><ul>`;
      for (const meeting of data.missingAgendas) {
        alertsContent += `<li>${meeting.title} (${formatTime(meeting.start)}) - ${meeting.clientName}</li>`;
      }
      alertsContent += `</ul>`;
    }

    if (data.overdueTasks.length > 0) {
      alertsContent += `<h3>Overdue Tasks</h3><ul>`;
      for (const task of data.overdueTasks) {
        alertsContent += `<li>[${task.client.client_name}] ${task.content}</li>`;
      }
      alertsContent += `</ul>`;
    }

    html += applyTemplate(alertsTemplate, { alerts_content: alertsContent });
  }

  // Today's Schedule
  html += `<h2>Today's Schedule</h2>`;
  if (data.meetings.length === 0) {
    html += `<p>No meetings scheduled for today.</p>`;
  } else {
    html += `<table style="width: 100%; border-collapse: collapse;">`;
    html += `<tr style="background-color: #f8f9fa;">`;
    html += `<th style="padding: 10px; text-align: left; border-bottom: 2px solid #dee2e6;">Time</th>`;
    html += `<th style="padding: 10px; text-align: left; border-bottom: 2px solid #dee2e6;">Meeting</th>`;
    html += `<th style="padding: 10px; text-align: left; border-bottom: 2px solid #dee2e6;">Client</th>`;
    html += `<th style="padding: 10px; text-align: left; border-bottom: 2px solid #dee2e6;">Agenda</th>`;
    html += `</tr>`;

    for (const meeting of data.meetings.sort((a, b) => a.start - b.start)) {
      html += `<tr>`;
      html += `<td style="padding: 10px; border-bottom: 1px solid #dee2e6;">${formatTime(meeting.start)} - ${formatTime(meeting.end)}</td>`;
      html += `<td style="padding: 10px; border-bottom: 1px solid #dee2e6;">${meeting.title}</td>`;
      html += `<td style="padding: 10px; border-bottom: 1px solid #dee2e6;">${meeting.clientName}</td>`;
      html += `<td style="padding: 10px; border-bottom: 1px solid #dee2e6;">${meeting.hasAgenda ? '✓' : '✗'}</td>`;
      html += `</tr>`;
    }

    html += `</table>`;
  }

  // Tasks by Client
  html += `<h2>Tasks by Client</h2>`;
  const clientIds = Object.keys(data.clientSummaries);

  if (clientIds.length === 0) {
    html += `<p>No tasks due today.</p>`;
  } else {
    for (const clientId of clientIds) {
      const summary = data.clientSummaries[clientId];
      if (summary.tasks.length > 0) {
        html += `<h3>${summary.client.client_name}</h3>`;
        html += `<ul>`;
        for (const task of summary.tasks) {
          const overdueStyle = task.isOverdue ? 'color: #dc3545;' : '';
          html += `<li style="${overdueStyle}">${task.content}`;
          if (task.due) {
            html += ` <em>(Due: ${task.due})</em>`;
          }
          html += `</li>`;
        }
        html += `</ul>`;
      }
    }
  }

  // Unread Emails Section (last 24 hours) - only if enabled in settings
  if (data.includeUnreadEmails) {
    html += `<h2>📧 Inbox Status</h2>`;
    if (data.unreadEmails.totalUnread === 0) {
      html += `<p style="color: #28a745;">✓ No unread emails!</p>`;
    } else {
      // Recent unread (last 24 hours)
      html += formatUnreadEmailsHtml(
        data.unreadEmails.recentUnread,
        '📬 Unread Emails - Last 24 Hours',
        '#17a2b8'
      );

      // Older backlog (if any)
      if (data.unreadEmails.olderBacklog.length > 0) {
        html += formatUnreadEmailsHtml(
          data.unreadEmails.olderBacklog,
          '⚠️ Older Unread Emails (Backlog)',
          '#dc3545'
        );
      }
    }
  }

  html += `</body></html>`;
  return html;
}

// ============================================================================
// WEEKLY OUTLOOK
// ============================================================================

/**
 * Generates and sends the weekly outlook report.
 * Called at 7:00 AM every Monday.
 * Uses Claude AI to generate strategic briefing from all data.
 */
function generateWeeklyOutlook() {
  Logger.log('Generating weekly outlook...');

  const today = new Date();
  const reportData = compileWeeklyData(today);

  // Generate HTML report using AI
  let htmlReport;
  const aiGenerated = generateWeeklyOutlookWithClaude(reportData, today);

  if (aiGenerated) {
    htmlReport = aiGenerated;
    Logger.log('Weekly outlook generated with AI');
  } else {
    // Fallback to template-based if AI fails
    htmlReport = formatWeeklyOutlookHtml(reportData, today);
    Logger.log('Weekly outlook generated with fallback template (AI unavailable)');
  }

  // Send email
  const props = PropertiesService.getScriptProperties();
  const weeklyLabel = props.getProperty('WEEKLY_BRIEFING_LABEL') || 'Brief: Weekly';
  const subject = `Weekly Outlook - Week of ${formatDate(today)}`;
  sendOutlookEmail(subject, htmlReport, weeklyLabel);

  Logger.log('Weekly outlook sent');
}

/**
 * Generates weekly outlook HTML using Claude AI.
 * Passes ALL data through the AI prompt for strategic analysis.
 *
 * @param {Object} data - The compiled report data
 * @param {Date} startDate - The start date of the week
 * @returns {string|null} AI-generated HTML or null if failed
 */
function generateWeeklyOutlookWithClaude(data, startDate) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    Logger.log('Claude API key not configured - falling back to template');
    return null;
  }

  const prompt = buildWeeklyOutlookPrompt(data, startDate);

  try {
    const url = 'https://api.anthropic.com/v1/messages';
    const model = getModelForPrompt('WEEKLY_BRIEFING_CLAUDE_PROMPT');

    const payload = {
      model: model,
      max_tokens: 8000,
      messages: [
        {
          role: 'user',
          content: prompt
        }
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
      Logger.log(`Claude API error for weekly outlook: ${responseCode} - ${response.getContentText()}`);
      return null;
    }

    // Parse response with explicit UTF-8 handling
    const responseText = response.getContentText('UTF-8');
    const result = JSON.parse(responseText);

    if (result.content && result.content.length > 0) {
      let content = result.content[0].text;

      // Ensure content is treated as UTF-8
      return content;
    }

    return null;

  } catch (error) {
    Logger.log(`Failed to generate weekly outlook with Claude: ${error.message}`);
    return null;
  }
}

/**
 * Builds the prompt for Claude weekly outlook generation.
 * Includes ALL available data: calendar, tasks, unread emails, drafts, conflicts.
 *
 * @param {Object} data - The compiled report data
 * @param {Date} startDate - The start date of the week
 * @returns {string} The formatted prompt
 */
function buildWeeklyOutlookPrompt(data, startDate) {
  // Build calendar summary for the week
  let calendarSummary = 'Week\'s Calendar Summary:\n';
  const dayKeys = Object.keys(data.dayData).sort();

  for (const dayKey of dayKeys) {
    const dayInfo = data.dayData[dayKey];
    const dayName = dayInfo.date.toLocaleDateString('en-US', { weekday: 'long' });
    calendarSummary += `\n${dayName} (${dayKey}):\n`;

    if (dayInfo.meetings.length === 0) {
      calendarSummary += '  No meetings\n';
    } else {
      for (const meeting of dayInfo.meetings.sort((a, b) => a.start - b.start)) {
        calendarSummary += `  - ${formatTime(meeting.start)}: ${meeting.title}`;
        if (meeting.clientName !== 'Unassigned') {
          calendarSummary += ` (${meeting.clientName})`;
        }
        if (!meeting.hasAgenda) {
          calendarSummary += ' [NO AGENDA]';
        }
        calendarSummary += '\n';
      }
    }
  }

  // Add conflicts
  if (data.conflicts.length > 0) {
    calendarSummary += '\nSchedule Conflicts:\n';
    for (const conflict of data.conflicts) {
      calendarSummary += `- "${conflict.event1.title}" overlaps with "${conflict.event2.title}"\n`;
    }
  }

  // Build tasks section (Todoist)
  let todoistTasks = 'Tasks by Client:\n';
  const clientNames = Object.keys(data.clientSummaries);

  if (clientNames.length > 0) {
    for (const clientName of clientNames) {
      const summary = data.clientSummaries[clientName];
      if (summary.tasks.length > 0) {
        todoistTasks += `\n${clientName}:\n`;
        for (const task of summary.tasks) {
          const overdueFlag = task.isOverdue ? ' [OVERDUE]' : '';
          const dueDate = task.due ? ` (Due: ${task.due})` : '';
          todoistTasks += `  - ${task.content}${dueDate}${overdueFlag}\n`;
        }
      }
    }
  } else {
    todoistTasks = 'Tasks by Client:\nNo tasks scheduled this week.\n';
  }

  // Build action today/this week emails section
  let actionEmails = '';

  // Fetch Action Today emails
  const actionTodayLabel = GmailApp.getUserLabelByName('Action Today');
  if (actionTodayLabel) {
    const actionTodayThreads = actionTodayLabel.getThreads(0, 10);
    if (actionTodayThreads.length > 0) {
      actionEmails += 'Emails labeled "Action Today":\n';
      for (const thread of actionTodayThreads) {
        const firstMessage = thread.getMessages()[0];
        actionEmails += `- From: ${firstMessage.getFrom()}, Subject: ${thread.getFirstMessageSubject()}\n`;
      }
      actionEmails += '\n';
    }
  }

  // Fetch Action This Week emails
  const actionWeekLabel = GmailApp.getUserLabelByName('Action This Week');
  if (actionWeekLabel) {
    const actionWeekThreads = actionWeekLabel.getThreads(0, 10);
    if (actionWeekThreads.length > 0) {
      actionEmails += 'Emails labeled "Action This Week":\n';
      for (const thread of actionWeekThreads) {
        const firstMessage = thread.getMessages()[0];
        actionEmails += `- From: ${firstMessage.getFrom()}, Subject: ${thread.getFirstMessageSubject()}\n`;
      }
      actionEmails += '\n';
    }
  }

  // Build unread emails section
  if (data.includeUnreadEmails && data.unreadEmails.totalUnread > 0) {
    actionEmails += `\nUnread Emails (${data.unreadEmails.totalUnread} total):\n`;

    if (data.unreadEmails.recentUnread.length > 0) {
      actionEmails += 'Recent (last 7 days):\n';
      for (const email of data.unreadEmails.recentUnread.slice(0, 15)) {
        actionEmails += `- From: ${email.from}, Subject: ${email.subject}\n`;
      }
    }

    if (data.unreadEmails.olderBacklog.length > 0) {
      actionEmails += `\nOlder backlog (${data.unreadEmails.olderBacklog.length} emails):\n`;
      for (const email of data.unreadEmails.olderBacklog.slice(0, 10)) {
        actionEmails += `- From: ${email.from}, Subject: ${email.subject} (${email.date})\n`;
      }
    }
  }

  // Build drafts section
  if (data.drafts && data.drafts.length > 0) {
    actionEmails += `\nUnsent Drafts (${data.drafts.length}):\n`;
    for (const draft of data.drafts.slice(0, 10)) {
      actionEmails += `- To: ${draft.to || 'No recipient'}, Subject: ${draft.subject || 'No subject'}\n`;
    }
  }

  // Get the prompt template
  const template = getPrompt('WEEKLY_BRIEFING_CLAUDE_PROMPT');

  // Apply variables to template
  return applyTemplate(template, {
    current_date: formatDate(startDate),
    action_today_emails: actionEmails || 'No action emails found.',
    action_this_week_emails: '', // Already included in actionEmails
    todoist_tasks: todoistTasks,
    calendar_summary: calendarSummary
  });
}

/**
 * Compiles data for the weekly outlook.
 *
 * @param {Date} startDate - The start date of the week
 * @returns {Object} Compiled report data
 */
function compileWeeklyData(startDate) {
  const data = {
    dayData: {},
    clientSummaries: {},
    conflicts: [],
    missingAgendas: [],
    overdueTasks: [],
    allMeetings: [],
    allTasks: [],
    unreadEmails: { recentUnread: [], olderBacklog: [], totalUnread: 0 },
    drafts: [],
    includeUnreadEmails: false
  };

  // Check if unread emails should be included (controlled by settings)
  const props = PropertiesService.getScriptProperties();
  const includeUnreadEmails = props.getProperty('INCLUDE_UNREAD_EMAILS') === 'true';
  data.includeUnreadEmails = includeUnreadEmails;

  if (includeUnreadEmails) {
    // Fetch unread emails from last 7 days for weekly report (older ones go to backlog)
    data.unreadEmails = fetchUnreadEmails(7);
  }

  // Fetch unsent drafts (always include these)
  data.drafts = fetchUnsentDrafts();

  // Initialize day data for each day of the week
  for (let i = 0; i < 7; i++) {
    const day = new Date(startDate);
    day.setDate(day.getDate() + i);
    const dayKey = formatDate(day);

    data.dayData[dayKey] = {
      date: day,
      meetings: [],
      tasks: []
    };
  }

  // Get week's events
  const calendar = CalendarApp.getDefaultCalendar();
  const endOfWeek = new Date(startDate);
  endOfWeek.setDate(endOfWeek.getDate() + 7);

  const events = calendar.getEvents(startDate, endOfWeek);

  // Process events
  for (const event of events) {
    if (event.isAllDayEvent()) continue;

    const eventDate = event.getStartTime();
    const dayKey = formatDate(eventDate);

    const eventInfo = {
      title: event.getTitle(),
      start: event.getStartTime(),
      end: event.getEndTime(),
      client: null,
      clientName: 'Unassigned',
      hasAgenda: false,
      dayKey: dayKey
    };

    // Identify client
    const client = identifyClientFromCalendarEvent(event);
    if (client) {
      eventInfo.client = client;
      eventInfo.clientName = client.client_name;
      eventInfo.hasAgenda = isAgendaGenerated(event.getId());

      if (!eventInfo.hasAgenda) {
        data.missingAgendas.push(eventInfo);
      }

      // Add to client summary
      if (!data.clientSummaries[client.client_name]) {
        data.clientSummaries[client.client_name] = {
          client: client,
          meetings: [],
          tasks: []
        };
      }
      data.clientSummaries[client.client_name].meetings.push(eventInfo);
    }

    if (data.dayData[dayKey]) {
      data.dayData[dayKey].meetings.push(eventInfo);
    }
    data.allMeetings.push(eventInfo);
  }

  // Detect conflicts for the entire week
  data.conflicts = detectScheduleConflicts(data.allMeetings);

  // Get tasks for all clients
  const clients = getClientRegistry();
  for (const client of clients) {
    if (client.todoist_project_id) {
      const tasks = fetchTodoistTasks(client.todoist_project_id);

      for (const task of tasks) {
        if (!task.due) continue;

        const dueDate = new Date(task.due.date);

        // Only include tasks due this week or overdue
        if (dueDate > endOfWeek && !isTaskOverdue(task, startDate)) {
          continue;
        }

        const dayKey = formatDate(dueDate);
        const taskInfo = {
          content: task.content,
          due: task.due.date,
          client: client,
          isOverdue: isTaskOverdue(task, startDate),
          dayKey: dayKey
        };

        if (taskInfo.isOverdue) {
          data.overdueTasks.push(taskInfo);
        }

        if (data.dayData[dayKey]) {
          data.dayData[dayKey].tasks.push(taskInfo);
        }

        data.allTasks.push(taskInfo);

        // Add to client summary
        if (!data.clientSummaries[client.client_name]) {
          data.clientSummaries[client.client_name] = {
            client: client,
            meetings: [],
            tasks: []
          };
        }
        data.clientSummaries[client.client_name].tasks.push(taskInfo);
      }
    }
  }

  return data;
}

/**
 * Formats the weekly outlook as HTML.
 *
 * @param {Object} data - The compiled report data
 * @param {Date} startDate - The start date of the week
 * @returns {string} HTML formatted report
 */
function formatWeeklyOutlookHtml(data, startDate) {
  let html = `<html><body style="font-family: Arial, sans-serif;">`;

  // Get header template from sheet
  const headerTemplate = getPrompt('WEEKLY_OUTLOOK_HEADER');
  html += applyTemplate(headerTemplate, { date: formatDate(startDate) });

  // Get summary template from sheet
  const summaryTemplate = getPrompt('WEEKLY_OUTLOOK_SUMMARY');
  html += applyTemplate(summaryTemplate, {
    meeting_count: data.allMeetings.length.toString(),
    task_count: data.allTasks.length.toString(),
    client_count: Object.keys(data.clientSummaries).length.toString()
  });

  // Alerts section
  if (data.conflicts.length > 0 || data.missingAgendas.length > 0 || data.overdueTasks.length > 0) {
    html += `<div style="background-color: #fff3cd; padding: 15px; margin-bottom: 20px; border-radius: 5px;">`;
    html += `<h2 style="margin-top: 0; color: #856404;">Alerts</h2>`;

    if (data.conflicts.length > 0) {
      html += `<h3>Schedule Conflicts</h3><ul>`;
      for (const conflict of data.conflicts) {
        html += `<li><strong>${conflict.event1.title}</strong> (${formatDate(conflict.event1.start)}) overlaps with <strong>${conflict.event2.title}</strong></li>`;
      }
      html += `</ul>`;
    }

    if (data.overdueTasks.length > 0) {
      html += `<h3>Overdue Tasks</h3><ul>`;
      for (const task of data.overdueTasks) {
        html += `<li>[${task.client.client_name}] ${task.content} <em>(Due: ${task.due})</em></li>`;
      }
      html += `</ul>`;
    }

    if (data.missingAgendas.length > 0) {
      html += `<h3>Meetings Without Agendas</h3><ul>`;
      for (const meeting of data.missingAgendas) {
        html += `<li>${meeting.title} - ${formatDate(meeting.start)} (${meeting.clientName})</li>`;
      }
      html += `</ul>`;
    }

    html += `</div>`;
  }

  // Day-by-day schedule
  html += `<h2>Day-by-Day Schedule</h2>`;

  const dayKeys = Object.keys(data.dayData).sort();
  for (const dayKey of dayKeys) {
    const dayInfo = data.dayData[dayKey];
    const dayName = getDayName(dayInfo.date);

    html += `<div style="margin-bottom: 20px; border: 1px solid #dee2e6; border-radius: 5px; overflow: hidden;">`;
    html += `<h3 style="background-color: #007bff; color: white; padding: 10px; margin: 0;">${dayName}, ${dayKey}</h3>`;

    html += `<div style="padding: 15px;">`;

    if (dayInfo.meetings.length === 0 && dayInfo.tasks.length === 0) {
      html += `<p style="color: #6c757d;">No meetings or tasks scheduled.</p>`;
    } else {
      if (dayInfo.meetings.length > 0) {
        html += `<strong>Meetings:</strong><ul>`;
        for (const meeting of dayInfo.meetings.sort((a, b) => a.start - b.start)) {
          const agendaStatus = meeting.hasAgenda ? '✓' : '✗';
          html += `<li>${formatTime(meeting.start)} - ${meeting.title} [${meeting.clientName}] Agenda: ${agendaStatus}</li>`;
        }
        html += `</ul>`;
      }

      if (dayInfo.tasks.length > 0) {
        html += `<strong>Tasks Due:</strong><ul>`;
        for (const task of dayInfo.tasks) {
          html += `<li>[${task.client.client_name}] ${task.content}</li>`;
        }
        html += `</ul>`;
      }
    }

    html += `</div></div>`;
  }

  // Client Summary
  html += `<h2>Client Summary</h2>`;
  const clientIds = Object.keys(data.clientSummaries);

  if (clientIds.length === 0) {
    html += `<p>No client activity this week.</p>`;
  } else {
    for (const clientId of clientIds) {
      const summary = data.clientSummaries[clientId];
      html += `<div style="margin-bottom: 15px; border: 1px solid #dee2e6; border-radius: 5px; overflow: hidden;">`;
      html += `<h3 style="background-color: #28a745; color: white; padding: 10px; margin: 0;">${summary.client.client_name}</h3>`;
      html += `<div style="padding: 15px;">`;

      html += `<p><strong>Meetings:</strong> ${summary.meetings.length} | <strong>Tasks:</strong> ${summary.tasks.length}</p>`;

      if (summary.meetings.length > 0) {
        html += `<strong>Scheduled Meetings:</strong><ul>`;
        for (const meeting of summary.meetings) {
          html += `<li>${formatDate(meeting.start)} ${formatTime(meeting.start)} - ${meeting.title}</li>`;
        }
        html += `</ul>`;
      }

      if (summary.tasks.length > 0) {
        html += `<strong>Tasks:</strong><ul>`;
        for (const task of summary.tasks) {
          const overdueStyle = task.isOverdue ? 'color: #dc3545;' : '';
          html += `<li style="${overdueStyle}">${task.content} <em>(Due: ${task.due})</em></li>`;
        }
        html += `</ul>`;
      }

      html += `</div></div>`;
    }
  }

  // Unread Emails Section - only if enabled in settings
  if (data.includeUnreadEmails) {
    html += `<h2>📧 Inbox Status</h2>`;
    if (data.unreadEmails.totalUnread === 0) {
      html += `<p style="color: #28a745;">✓ No unread emails!</p>`;
    } else {
      // Recent unread (last 7 days)
      html += formatUnreadEmailsHtml(
        data.unreadEmails.recentUnread,
        '📬 Unread Emails - Last 7 Days',
        '#17a2b8'
      );

      // Older backlog
      if (data.unreadEmails.olderBacklog.length > 0) {
        html += formatUnreadEmailsHtml(
          data.unreadEmails.olderBacklog,
          '⚠️ Older Unread Emails (Backlog)',
          '#dc3545'
        );
      }
    }
  }

  // Unsent Drafts Section
  html += `<h2>📝 Unsent Drafts</h2>`;
  if (data.drafts.length === 0) {
    html += `<p style="color: #28a745;">✓ No unsent drafts!</p>`;
  } else {
    html += formatDraftsHtml(data.drafts);
  }

  html += `</body></html>`;
  return html;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Detects schedule conflicts in a list of meetings.
 *
 * @param {Object[]} meetings - Array of meeting info objects
 * @returns {Object[]} Array of conflict objects
 */
function detectScheduleConflicts(meetings) {
  const conflicts = [];

  for (let i = 0; i < meetings.length; i++) {
    for (let j = i + 1; j < meetings.length; j++) {
      const event1 = meetings[i];
      const event2 = meetings[j];

      // Check if events overlap
      if (eventsOverlap(event1, event2)) {
        conflicts.push({
          event1: event1,
          event2: event2
        });
      }
    }
  }

  return conflicts;
}

/**
 * Checks if two events overlap.
 *
 * @param {Object} event1 - First event
 * @param {Object} event2 - Second event
 * @returns {boolean} True if events overlap
 */
function eventsOverlap(event1, event2) {
  const start1 = event1.start.getTime();
  const end1 = event1.end.getTime();
  const start2 = event2.start.getTime();
  const end2 = event2.end.getTime();

  // Events overlap if one starts before the other ends
  return start1 < end2 && start2 < end1;
}

/**
 * Checks if a Todoist task is overdue.
 *
 * @param {Object} task - Todoist task object
 * @param {Date} referenceDate - Date to compare against
 * @returns {boolean} True if task is overdue
 */
function isTaskOverdue(task, referenceDate) {
  if (!task.due) return false;

  const dueDate = new Date(task.due.date);
  dueDate.setHours(23, 59, 59, 999);

  const today = new Date(referenceDate);
  today.setHours(0, 0, 0, 0);

  return dueDate < today;
}

/**
 * Sends an outlook email with a label.
 *
 * @param {string} subject - Email subject
 * @param {string} htmlBody - HTML email body
 * @param {string} labelName - Label to apply
 */
function sendOutlookEmail(subject, htmlBody, labelName) {
  const userEmail = getCurrentUserEmail();

  // Ensure proper UTF-8 encoding by adding meta tag if not present
  let body = htmlBody;
  if (!body.match(/<meta[^>]+charset/i)) {
    // If body doesn't have HTML structure, wrap it
    // Use string concatenation instead of template literals to preserve UTF-8
    if (!body.match(/<html/i)) {
      body = '<!DOCTYPE html>\n' +
             '<html>\n' +
             '<head>\n' +
             '<meta charset="UTF-8">\n' +
             '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">\n' +
             '</head>\n' +
             '<body>\n' +
             body +
             '\n</body>\n' +
             '</html>';
    } else {
      // Insert meta tag in existing HTML
      body = body.replace(/<head[^>]*>/i, function(match) {
        return match + '\n<meta charset="UTF-8">\n<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">';
      });
    }
  }

  // Send email with explicit UTF-8 encoding and plain text fallback
  GmailApp.sendEmail(userEmail, subject, 'This email requires HTML support.', {
    htmlBody: body,
    charset: 'utf-8'
  });

  // Apply label to the sent email
  Utilities.sleep(2000); // Wait for email to be sent

  const query = `from:me to:me subject:"${subject}" newer_than:1h`;
  const threads = GmailApp.search(query, 0, 1);

  if (threads.length > 0) {
    const label = createLabelIfNotExists(labelName);
    threads[0].addLabel(label);
  }
}

/**
 * Gets the day name for a date.
 *
 * @param {Date} date - The date
 * @returns {string} Day name (e.g., "Monday")
 */
function getDayName(date) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return days[date.getDay()];
}

// ============================================================================
// UNREAD EMAILS & DRAFTS
// ============================================================================

/**
 * Fetches unread emails from inbox with filtering.
 *
 * @param {number} daysBack - How many days back to search
 * @returns {Object} Object with recentUnread and olderUnread arrays
 */
function fetchUnreadEmails(daysBack) {
  const result = {
    recentUnread: [],    // Unread from last `daysBack` days
    olderBacklog: [],    // Unread older than `daysBack` days
    totalUnread: 0
  };

  try {
    // Get all unread emails
    const threads = GmailApp.search('is:unread -from:me', 0, 100);

    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysBack);

    for (const thread of threads) {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      const date = lastMessage.getDate();

      // Get labels for context
      const labels = thread.getLabels().map(l => l.getName());

      const emailInfo = {
        subject: thread.getFirstMessageSubject(),
        from: lastMessage.getFrom(),
        date: date,
        snippet: lastMessage.getPlainBody().substring(0, 200),
        labels: labels,
        isStarred: thread.hasStarredMessages(),
        messageCount: messages.length
      };

      if (date >= cutoffDate) {
        result.recentUnread.push(emailInfo);
      } else {
        result.olderBacklog.push(emailInfo);
      }

      result.totalUnread++;
    }

    // Sort by date descending (newest first)
    result.recentUnread.sort((a, b) => b.date - a.date);
    result.olderBacklog.sort((a, b) => b.date - a.date);

  } catch (error) {
    Logger.log(`Error fetching unread emails: ${error.message}`);
  }

  return result;
}

/**
 * Fetches unsent drafts from Gmail.
 *
 * @returns {Object[]} Array of draft info objects
 */
function fetchUnsentDrafts() {
  const drafts = [];

  try {
    const gmailDrafts = GmailApp.getDrafts();

    for (const draft of gmailDrafts) {
      const message = draft.getMessage();
      const to = message.getTo();
      const subject = message.getSubject();
      const date = message.getDate();

      // Calculate age
      const ageMs = Date.now() - date.getTime();
      const ageDays = Math.floor(ageMs / (1000 * 60 * 60 * 24));

      drafts.push({
        subject: subject || '(No subject)',
        to: to || '(No recipient)',
        date: date,
        ageDays: ageDays,
        snippet: message.getPlainBody().substring(0, 150),
        draftId: draft.getId()
      });
    }

    // Sort by date descending (newest first)
    drafts.sort((a, b) => b.date - a.date);

  } catch (error) {
    Logger.log(`Error fetching drafts: ${error.message}`);
  }

  return drafts;
}

/**
 * Marks unread emails older than specified days as read.
 * Only runs if enabled in settings.
 *
 * @returns {Object} Result with count of emails marked
 */
function autoMarkOldEmailsAsRead() {
  const props = PropertiesService.getScriptProperties();
  const daysThreshold = parseInt(props.getProperty('AUTO_MARK_READ_AFTER_DAYS') || '0', 10);

  if (daysThreshold <= 0) {
    return { success: true, count: 0, message: 'Auto-mark-read is disabled' };
  }

  try {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysThreshold);

    // Search for unread emails older than threshold
    const query = `is:unread before:${formatDateForGmail(cutoffDate)}`;
    const threads = GmailApp.search(query, 0, 50);

    let markedCount = 0;
    for (const thread of threads) {
      thread.markRead();
      markedCount++;
    }

    if (markedCount > 0) {
      logProcessing('AUTO_MARK_READ', null, `Marked ${markedCount} old emails as read`, 'success');
    }

    return { success: true, count: markedCount, message: `Marked ${markedCount} emails as read` };

  } catch (error) {
    logProcessing('AUTO_MARK_READ', null, `Error: ${error.message}`, 'error');
    return { success: false, count: 0, message: error.message };
  }
}

/**
 * Formats a date for Gmail search query (YYYY/MM/DD).
 *
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDateForGmail(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

/**
 * Formats unread emails section for HTML output.
 *
 * @param {Object[]} emails - Array of email info objects
 * @param {string} title - Section title
 * @param {string} bgColor - Background color for header
 * @returns {string} HTML string
 */
function formatUnreadEmailsHtml(emails, title, bgColor) {
  if (emails.length === 0) {
    return '';
  }

  let html = `<div style="margin-bottom: 20px; border: 1px solid #dee2e6; border-radius: 5px; overflow: hidden;">`;
  html += `<h3 style="background-color: ${bgColor}; color: white; padding: 10px; margin: 0;">${title} (${emails.length})</h3>`;
  html += `<div style="padding: 15px;">`;
  html += `<p style="font-size: 12px; color: #666; margin-top: 0;"><em>⚠️ THESE ARE UNREAD EMAILS IN YOUR INBOX - NOT SENT BY YOU</em></p>`;
  html += `<table style="width: 100%; border-collapse: collapse; font-size: 13px;">`;
  html += `<tr style="background: #f8f9fa;"><th style="text-align: left; padding: 8px;">From</th><th style="text-align: left; padding: 8px;">Subject</th><th style="text-align: left; padding: 8px;">Date</th><th style="text-align: left; padding: 8px;">Labels</th></tr>`;

  for (const email of emails.slice(0, 15)) {
    const labelsStr = email.labels.length > 0 ? email.labels.join(', ') : '-';
    const starIcon = email.isStarred ? '⭐ ' : '';
    html += `<tr style="border-bottom: 1px solid #eee;">`;
    html += `<td style="padding: 8px;">${escapeHtml(email.from.split('<')[0].trim())}</td>`;
    html += `<td style="padding: 8px;">${starIcon}${escapeHtml(email.subject)}</td>`;
    html += `<td style="padding: 8px; white-space: nowrap;">${formatDateShort(email.date)}</td>`;
    html += `<td style="padding: 8px; font-size: 11px;">${escapeHtml(labelsStr)}</td>`;
    html += `</tr>`;
  }

  if (emails.length > 15) {
    html += `<tr><td colspan="4" style="padding: 8px; text-align: center; color: #666;">... and ${emails.length - 15} more</td></tr>`;
  }

  html += `</table></div></div>`;
  return html;
}

/**
 * Formats drafts section for HTML output.
 *
 * @param {Object[]} drafts - Array of draft info objects
 * @returns {string} HTML string
 */
function formatDraftsHtml(drafts) {
  if (drafts.length === 0) {
    return '';
  }

  let html = `<div style="margin-bottom: 20px; border: 1px solid #dee2e6; border-radius: 5px; overflow: hidden;">`;
  html += `<h3 style="background-color: #6c757d; color: white; padding: 10px; margin: 0;">📝 Unsent Drafts (${drafts.length})</h3>`;
  html += `<div style="padding: 15px;">`;
  html += `<p style="font-size: 12px; color: #dc3545; margin-top: 0;"><strong>⚠️ THESE ARE UNSENT DRAFTS - NOT SENT EMAILS</strong></p>`;
  html += `<table style="width: 100%; border-collapse: collapse; font-size: 13px;">`;
  html += `<tr style="background: #f8f9fa;"><th style="text-align: left; padding: 8px;">To</th><th style="text-align: left; padding: 8px;">Subject</th><th style="text-align: left; padding: 8px;">Age</th></tr>`;

  for (const draft of drafts) {
    const ageStr = draft.ageDays === 0 ? 'Today' : draft.ageDays === 1 ? '1 day' : `${draft.ageDays} days`;
    const ageColor = draft.ageDays > 7 ? '#dc3545' : draft.ageDays > 3 ? '#ffc107' : '#28a745';
    html += `<tr style="border-bottom: 1px solid #eee;">`;
    html += `<td style="padding: 8px;">${escapeHtml(draft.to)}</td>`;
    html += `<td style="padding: 8px;">${escapeHtml(draft.subject)}</td>`;
    html += `<td style="padding: 8px; color: ${ageColor}; font-weight: bold;">${ageStr}</td>`;
    html += `</tr>`;
  }

  html += `</table></div></div>`;
  return html;
}

/**
 * Escapes HTML special characters.
 *
 * @param {string} text - Text to escape
 * @returns {string} Escaped text
 */
function escapeHtml(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
