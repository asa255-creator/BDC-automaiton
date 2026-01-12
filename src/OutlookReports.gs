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
 */
function generateDailyOutlook() {
  Logger.log('Generating daily outlook...');

  const today = new Date();
  const reportData = compileDailyData(today);

  // Generate HTML report
  const htmlReport = formatDailyOutlookHtml(reportData, today);

  // Send email
  const subject = `Daily Outlook - ${formatDate(today)}`;
  sendOutlookEmail(subject, htmlReport, 'Brief: Daily');

  Logger.log('Daily outlook sent');
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
    overdueTasks: []
  };

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

  html += `</body></html>`;
  return html;
}

// ============================================================================
// WEEKLY OUTLOOK
// ============================================================================

/**
 * Generates and sends the weekly outlook report.
 * Called at 7:00 AM every Monday.
 */
function generateWeeklyOutlook() {
  Logger.log('Generating weekly outlook...');

  const today = new Date();
  const reportData = compileWeeklyData(today);

  // Generate HTML report
  const htmlReport = formatWeeklyOutlookHtml(reportData, today);

  // Send email
  const subject = `Weekly Outlook - Week of ${formatDate(today)}`;
  sendOutlookEmail(subject, htmlReport, 'Brief: Weekly');

  Logger.log('Weekly outlook sent');
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
    allTasks: []
  };

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

  GmailApp.sendEmail(userEmail, subject, '', {
    htmlBody: htmlBody
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
