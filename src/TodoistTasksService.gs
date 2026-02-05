/**
 * TodoistTasksService.gs - Todoist API integrations.
 */

/**
 * Fetches collaborators for a Todoist project.
 *
 * @param {string} projectId - The Todoist project ID
 * @returns {Object[]} Array of collaborator objects
 */
function fetchProjectCollaborators(projectId) {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    return [];
  }

  try {
    const url = `https://api.todoist.com/rest/v2/projects/${projectId}/collaborators`;

    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiToken}`
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      logProcessing('TODOIST', null, `Failed to fetch collaborators: ${responseCode}`, 'error');
      return [];
    }

    return JSON.parse(response.getContentText());

  } catch (error) {
    logProcessing('TODOIST', null, `Error fetching collaborators: ${error.message}`, 'error');
    return [];
  }
}

/**
 * Creates Todoist tasks with assignee matching.
 *
 * @param {Object[]} actionItems - Array of action items from AI extraction
 * @param {Object} client - The client object
 */
function createTodoistTasksWithAssignees(actionItems, client) {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    logProcessing('TODOIST', client.client_name, 'Todoist API token not configured', 'error');
    return;
  }

  const projectId = client.todoist_project_id;
  let createdCount = 0;

  for (const item of actionItems) {
    try {
      const url = 'https://api.todoist.com/rest/v2/tasks';

      const payload = {
        content: item.title || item.description.substring(0, 100),
        description: item.description,
        project_id: projectId
      };

      // Add assignee if we have one
      if (item.assignee_id) {
        payload.assignee_id = item.assignee_id;
      }

      // Add due date
      if (item.due_date) {
        payload.due_date = item.due_date;
      }

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

      if (responseCode === 200) {
        createdCount++;
        const assigneeInfo = item.assignee_name ? ` (assigned to ${item.assignee_name})` : '';
        logProcessing('TODOIST', client.client_name, `Created task: ${item.title}${assigneeInfo}`, 'success');
      } else {
        logProcessing('TODOIST', client.client_name, `Failed to create task: ${responseCode}`, 'error');
      }

    } catch (error) {
      logProcessing('TODOIST', client.client_name, `Error creating task: ${error.message}`, 'error');
    }
  }

  logProcessing('TODOIST', client.client_name, `Created ${createdCount}/${actionItems.length} tasks`, 'success');
}

/**
 * Creates Todoist tasks for action items.
 *
 * @param {Object[]} actionItems - Array of action item objects
 * @param {Object} client - The client object
 */
function createTodoistTasks(actionItems, client) {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    Logger.log('Todoist API token not configured');
    return;
  }

  const projectId = client.todoist_project_id;

  for (const item of actionItems) {
    try {
      createTodoistTask(apiToken, projectId, item, client.client_name);
    } catch (error) {
      Logger.log(`Failed to create Todoist task: ${error.message}`);
      logProcessing(
        'TODOIST_ERROR',
        client.client_name,
        `Failed to create task: ${item.description}`,
        'error'
      );
    }
  }
}

/**
 * Creates a single Todoist task.
 *
 * @param {string} apiToken - Todoist API token
 * @param {string} projectId - Todoist project ID
 * @param {Object} item - Action item object
 * @param {string} clientName - Client name for task content
 */
function createTodoistTask(apiToken, projectId, item, clientName) {
  const url = 'https://api.todoist.com/rest/v2/tasks';

  const taskContent = `[${clientName}] ${item.description}`;

  const payload = {
    content: taskContent,
    project_id: projectId
  };

  // Add due date if provided
  if (item.due_date) {
    payload.due_string = item.due_date;
  }

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
    throw new Error(`Todoist API error: ${responseCode}`);
  }

  Logger.log(`Created Todoist task: ${taskContent}`);
}

/**
 * Fetches tasks from Todoist for a specific project.
 *
 * @param {string} projectId - Todoist project ID
 * @returns {Object[]} Array of task objects
 */
function fetchTodoistTasks(projectId) {
  const apiToken = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

  if (!apiToken) {
    Logger.log('Todoist API token not configured');
    return [];
  }

  const url = `https://api.todoist.com/rest/v2/tasks?project_id=${projectId}`;

  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${apiToken}`
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`Todoist API error: ${responseCode}`);
      return [];
    }

    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log(`Failed to fetch Todoist tasks: ${error.message}`);
    return [];
  }
}

/**
 * Fetches tasks due today or overdue for a project.
 *
 * @param {string} projectId - Todoist project ID
 * @returns {Object[]} Array of task objects due today or overdue
 */
function fetchTodoistTasksDueToday(projectId, targetDate) {
  const tasks = fetchTodoistTasks(projectId);
  const day = targetDate ? new Date(targetDate) : new Date();
  day.setHours(23, 59, 59, 999);

  return tasks.filter(task => {
    if (!task.due) return false;

    const dueDate = new Date(task.due.date);
    return dueDate <= day;
  });
}
