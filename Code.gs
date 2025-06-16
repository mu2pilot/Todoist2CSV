// === GLOBAL CONFIGURATION ===
var TODOIST_API_TOKEN = PropertiesService.getScriptProperties().getProperty('TODOIST_API_TOKEN');

// === MENU SETUP ===
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Todoist Tools')
    .addItem('Update Sheet from Todoist', 'updateTodoistTasks')
    .addItem('Push Changes to Todoist', 'updateTodoistTaskFromSheet')
    .addSeparator()
    .addItem('Sort Tasks', 'sortTasks')
    .addSeparator()
    .addItem('Apply Colorful Formatting', 'applyTaskConditionalFormatting')
    .addToUi();
}

// === FETCH TASKS FROM TODOIST ===
function updateTodoistTasks() {
  var response = UrlFetchApp.fetch('https://api.todoist.com/rest/v2/tasks', {
    headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN }
  });
  var tasks = JSON.parse(response.getContentText());
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  var rows = [
    ['ID', 'Task', 'TaskLink', 'Project Name', 'Project ID', 'Due Date', 'Due Time', 'Recurring', 'Priority', 'Label1', 'Label2', 'Label3', 'Completed', 'Last Modified']
  ];

  // Fetch projects for project name mapping
  var projectsResponse = UrlFetchApp.fetch('https://api.todoist.com/rest/v2/projects', {
    headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN }
  });
  var projects = JSON.parse(projectsResponse.getContentText());
  var projectDict = {};
  projects.forEach(function(project) {
    projectDict[project.id] = project.name;
  });

  tasks.forEach(function(task) {
    var content = task.content || '';
    var taskVisible = content.replace(/\[(.*?)\]\((.*?)\)/g, '$1');
    var taskLink = '';
    var match = content.match(/\[(.*?)\]\((.*?)\)/);
    if (match) {
      var url = match[2];
      taskLink = '=HYPERLINK("' + url + '", "' + url + '")';
    }

    var labels = task.labels || [];
    var orderedLabels = labels.slice();
    var labelDict = {};
    if (orderedLabels.indexOf('GCFO') !== -1) {
      orderedLabels.splice(orderedLabels.indexOf('GCFO'), 1);
      orderedLabels.unshift('GCFO');
    }
    for (var i = 0; i < 3; i++) {
      labelDict['Label' + (i + 1)] = orderedLabels[i] || '';
    }

    var dueDate = task.due ? task.due.date : '';
    var dueTime = '';
    if (task.due && task.due.datetime) {
      // Extract the time part directly from the string (assume it's in local time)
      dueTime = task.due.datetime.substring(11, 16); // HH:MM
    }
    var recurrenceString = '';
    if (task.due && task.due.string && /^every/i.test(task.due.string.trim())) {
      recurrenceString = task.due.string;
    }

    Logger.log('TaskId: ' + task.id +
               ' | due.date: ' + (task.due ? task.due.date : '') +
               ' | due.datetime: ' + (task.due ? task.due.datetime : '') +
               ' | dueDate: ' + dueDate +
               ' | dueTime: ' + dueTime);

    // Defensive: Priority always 1-4, default 4
    var uiPriority = 4;
    if (typeof task.priority === 'number' && !isNaN(task.priority)) {
      uiPriority = 5 - task.priority;
      if (uiPriority < 1 || uiPriority > 4) uiPriority = 4;
    }

    rows.push([
      task.id,
      taskVisible,
      taskLink,
      projectDict[task.project_id] || '',
      task.project_id,
      dueDate,
      dueTime,
      recurrenceString,
      uiPriority,
      labelDict.Label1,
      labelDict.Label2,
      labelDict.Label3,
      false,
      '' // Last Modified is blank after refresh
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  // Set up checkboxes in Completed column (dynamically found)
  var completedCol = rows[0].indexOf('Completed') + 1;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, completedCol, sheet.getLastRow() - 1, 1).insertCheckboxes();
  }
  // Freeze the top row
  sheet.setFrozenRows(1);
  // Reset lastSyncTime after refresh
  var now = new Date().toISOString();
  PropertiesService.getDocumentProperties().setProperty('lastSyncTime', now);
  Logger.log('Reset last sync time to: ' + now);

  // Automatically sort and apply formatting after update
  sortTasks();
  applyTaskConditionalFormatting();
  hideIdColumns();
}

// === PUSH CHANGES TO TODOIST (SAFE, WITH PROJECT MOVE) ===
function getCentralOffset(dateString) {
  var date = new Date(dateString);
  // Central Time is UTC-6, but UTC-5 during DST
  var jan = new Date(date.getFullYear(), 0, 1);
  var jul = new Date(date.getFullYear(), 6, 1);
  var stdTimezoneOffset = Math.max(jan.getTimezoneOffset(), jul.getTimezoneOffset());
  var isDST = date.getTimezoneOffset() < stdTimezoneOffset;
  return isDST ? '-05:00' : '-06:00';
}

function formatDueTimeCell(cellValue) {
  if (typeof cellValue === 'string') {
    return cellValue.trim();
  }
  if (typeof cellValue === 'number') {
    // Only treat as time if less than 1 (i.e., a time, not a date)
    if (cellValue < 1) {
      var totalMinutes = Math.round(cellValue * 24 * 60);
      var hours = Math.floor(totalMinutes / 60);
      var minutes = totalMinutes % 60;
      return ('0' + hours).slice(-2) + ':' + ('0' + minutes).slice(-2);
    } else {
      // Not a valid time, ignore
      return '';
    }
  }
  if (cellValue instanceof Date) {
    // Use local hours and minutes (not UTC)
    var hours = cellValue.getHours();
    var minutes = cellValue.getMinutes();
    return ('0' + hours).slice(-2) + ':' + ('0' + minutes).slice(-2);
  }
  return '';
}

// === URL FORMATTING FUNCTION ===
function formatUrlToMarkdown(url, linkText) {
  if (!url) return '';
  // If it's a HYPERLINK formula, extract the URL and text
  if (url.toString().includes('=HYPERLINK')) {
    var match = url.toString().match(/=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)/);
    if (match) {
      url = match[1];
      linkText = match[2];
    }
  }
  url = url.toString().trim();
  // Add https:// only if the URL doesn't have a protocol prefix
  if (!url.match(/^[a-zA-Z]+:\/\//)) {
    url = 'https://' + url;
  }
  // Always use the provided linkText (task name) if present
  return '[' + (linkText || url.replace(/\/$/, '')) + '](' + url + ')';
}

// === ON EDIT TRIGGER: Automatically update 'Last Modified' column (L) ===
function onEdit(e) {
  Logger.log('onEdit triggered');
  Logger.log('Range: ' + e.range.getA1Notation());
  Logger.log('Column: ' + e.range.getColumn());
  Logger.log('Row: ' + e.range.getRow());
  
  var sheet = e.range.getSheet();
  var colToWatch = [1,2,3,4,5,6,7,8,9,10,11,12,13]; // columns A-M (1-based)
  var lastModifiedCol = 14; // column N (1-based index)
  
  Logger.log('Is column in watch list? ' + (colToWatch.indexOf(e.range.getColumn()) !== -1));
  Logger.log('Is row > 1? ' + (e.range.getRow() > 1));
  
  if (colToWatch.indexOf(e.range.getColumn()) !== -1 && e.range.getRow() > 1) {
    Logger.log('Setting timestamp in column N');
    sheet.getRange(e.range.getRow(), lastModifiedCol).setValue(new Date());
  }
}

// === PROJECT MANAGEMENT FUNCTIONS ===
function getProjectIdFromName(projectName) {
  // Fetch all projects
  var projectsResponse = UrlFetchApp.fetch('https://api.todoist.com/rest/v2/projects', {
    headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN }
  });
  var projects = JSON.parse(projectsResponse.getContentText());
  
  // Create name to ID mapping
  var projectNameToId = {};
  projects.forEach(function(project) {
    projectNameToId[project.name.toLowerCase()] = project.id;
  });
  
  // Check if project exists
  if (projectNameToId[projectName.toLowerCase()]) {
    return projectNameToId[projectName.toLowerCase()];
  }
  
  // Project doesn't exist, create it
  var createProjectUrl = 'https://api.todoist.com/rest/v2/projects';
  var createProjectPayload = JSON.stringify({
    name: projectName
  });
  var createProjectOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN },
    payload: createProjectPayload,
    muteHttpExceptions: true
  };
  
  var createResponse = UrlFetchApp.fetch(createProjectUrl, createProjectOptions);
  if (createResponse.getResponseCode() === 200) {
    var newProject = JSON.parse(createResponse.getContentText());
    Logger.log('Created new project: ' + projectName + ' with ID: ' + newProject.id);
    return newProject.id;
  } else {
    Logger.log('Failed to create project: ' + projectName);
    return null;
  }
}

// === MAIN UPDATE FUNCTION: Only update rows changed since last sync ===
function updateTodoistTaskFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var lastSyncTime = PropertiesService.getDocumentProperties().getProperty('lastSyncTime');
  var lastSyncDate = lastSyncTime ? new Date(lastSyncTime) : null;
  Logger.log('Last sync time: ' + lastSyncTime);

  var updatedRowsLog = [];
  var headers = data[0];
  var completedCol = headers.indexOf('Completed'); // 0-based
  var lastModifiedCol = headers.indexOf('Last Modified'); // 0-based
  var idCol = headers.indexOf('ID'); // 0-based
  var dueDateCol = headers.indexOf('Due Date'); // 0-based
  var dueTimeCol = headers.indexOf('Due Time'); // 0-based
  var recurringCol = headers.indexOf('Recurring'); // 0-based

  var rowsToDelete = [];
  var recurringTaskUpdated = false;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // Pad row to header length for safety
    if (row.length < headers.length) {
      row = row.concat(Array(headers.length - row.length).fill(''));
    }
    var lastModified = row[lastModifiedCol];
    Logger.log('Row ' + (i+1) + ' | Last Modified: ' + lastModified);
    var taskId = row[0]; // ID column (A)
    var isCompleted = row[completedCol] === true;
    if (isCompleted && taskId) {
      // Use Recurring column to determine if task is recurring
      var recurrenceCell = row[recurringCol] ? row[recurringCol].toString().toLowerCase() : '';
      var isRowRecurring = recurrenceCell && recurrenceCell.startsWith('every');
      if (isRowRecurring) {
        // 1. Close the task in Todoist (this advances the recurrence)
        var closeTaskUrl = 'https://api.todoist.com/rest/v2/tasks/' + taskId + '/close';
        var closeTaskOptions = {
          method: 'post',
          headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN },
          muteHttpExceptions: true
        };
        var closeTaskResponse = UrlFetchApp.fetch(closeTaskUrl, closeTaskOptions);
        Logger.log('TaskId: ' + taskId + ' | Close recurring task response code: ' + closeTaskResponse.getResponseCode());

        // 2. If successful, fetch the updated task to get the new due date/time
        if (closeTaskResponse.getResponseCode() === 204) {
          Utilities.sleep(1000); // Wait a moment for Todoist to update
          var getTaskUrl = 'https://api.todoist.com/rest/v2/tasks/' + taskId;
          var getTaskOptions = {
            method: 'get',
            headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN },
            muteHttpExceptions: true
          };
          var getTaskResponse = UrlFetchApp.fetch(getTaskUrl, getTaskOptions);
          Logger.log('TaskId: ' + taskId + ' | Fetch after close response code: ' + getTaskResponse.getResponseCode());
          if (getTaskResponse.getResponseCode() === 200) {
            var updatedTask = JSON.parse(getTaskResponse.getContentText());
            sheet.getRange(i + 1, dueDateCol + 1).setValue(updatedTask.due ? updatedTask.due.date : '');
            var newDueTime = '';
            if (updatedTask.due && updatedTask.due.datetime) {
              newDueTime = updatedTask.due.datetime.substring(11, 16);
            }
            sheet.getRange(i + 1, dueTimeCol + 1).setValue(newDueTime);
            // Uncheck Completed and clear Last Modified
            sheet.getRange(i + 1, completedCol + 1).setValue(false);
            sheet.getRange(i + 1, lastModifiedCol + 1).setValue('');
            recurringTaskUpdated = true;
            updatedRowsLog.push({ row: i+1, taskId: taskId, completed: 'recurring', newDueDate: updatedTask.due ? updatedTask.due.date : '', newDueTime: newDueTime });
          } else {
            Logger.log('TaskId: ' + taskId + ' | Could not fetch updated recurring task after close.');
          }
        } else {
          Logger.log('TaskId: ' + taskId + ' | Failed to close recurring task in Todoist. Response: ' + closeTaskResponse.getContentText());
        }
      } else {
        // Non-recurring: close the task in Todoist first
        var closeTaskUrl = 'https://api.todoist.com/rest/v2/tasks/' + taskId + '/close';
        var closeTaskOptions = {
          method: 'post',
          headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN },
          muteHttpExceptions: true
        };
        var closeTaskResponse = UrlFetchApp.fetch(closeTaskUrl, closeTaskOptions);
        Logger.log('TaskId: ' + taskId + ' | Close task response code: ' + closeTaskResponse.getResponseCode());
        
        if (closeTaskResponse.getResponseCode() === 204) { // 204 means success for close operation
          // Only delete row after successful close
          rowsToDelete.push(i + 1); // 1-based row index
          updatedRowsLog.push({ row: i+1, taskId: taskId, completed: 'fully (non-recurring, Recurring col)' });
        } else {
          Logger.log('TaskId: ' + taskId + ' | Failed to close task in Todoist. Response: ' + closeTaskResponse.getContentText());
        }
      }
      continue; // Skip further updates for this row
    }
    // === NEW TASK CREATION ===
    if ((!taskId || taskId.toString().trim() === '') && lastModified) {
      Logger.log('Row ' + (i+1) + ' | Detected new task (blank ID, has Last Modified)');
      var taskContent = row[1]; // Task (B)
      var taskLink = row[2];    // Link (C)
      var finalContent = taskContent;
      if (taskLink && taskLink.toString().trim() !== '') {
        finalContent = formatUrlToMarkdown(taskLink, taskContent);
      }
      
      var projectName = row[3];   // Project Name column (D)
      var projectId = null;
      if (projectName && projectName.toString().trim() !== '') {
        projectId = getProjectIdFromName(projectName);
        if (!projectId) {
          Logger.log('TaskId: ' + taskId + ' | Failed to get/create project: ' + projectName);
          continue; // Skip this task if project creation failed
        }
      }
      
      var dueDate = row[5];     // Due Date (F)
      if (dueDate instanceof Date) {
        dueDate = dueDate.getFullYear() + '-' +
                  ('0' + (dueDate.getMonth() + 1)).slice(-2) + '-' +
                  ('0' + dueDate.getDate()).slice(-2);
      }
      var dueTime = formatDueTimeCell(row[6]); // Due Time (G)
      // Label handling using header-based indexing
      var label1Col = headers.indexOf('Label1');
      var label2Col = headers.indexOf('Label2');
      var label3Col = headers.indexOf('Label3');
      var labelCols = [label1Col, label2Col, label3Col];
      var labels = [];
      labelCols.forEach(function(colIdx) {
        if (row[colIdx] !== undefined && row[colIdx] !== null && row[colIdx].toString().trim() !== '') {
          labels.push(row[colIdx]);
        }
      });
      var payloadObj = {
        content: finalContent,
        // Priority 4 in the sheet means priority 1 in the API (lowest)
        priority: 1
      };
      if (projectId && projectId.toString().trim() !== '') {
        payloadObj.project_id = projectId;
      }
      if (labels.length > 0) {
        payloadObj.labels = labels;
      }
      if (dueDate && dueDate.toString().trim() !== '') {
        if (dueTime && dueTime.toString().trim() !== '') {
          var offset = getCentralOffset(dueDate);
          var timeString = dueTime.length === 5 ? dueTime : ('0' + dueTime).slice(-5);
          var isoString = dueDate + 'T' + timeString + ':00' + offset;
          Logger.log('Creating new task | due_datetime: ' + isoString);
          payloadObj.due_datetime = isoString;
        } else {
          payloadObj.due_date = dueDate;
        }
      }
      var payload = JSON.stringify(payloadObj);
      Logger.log('Creating new task | Payload: ' + payload);
      var urlApi = 'https://api.todoist.com/rest/v2/tasks';
      var options = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN, 'X-Request-Id': Utilities.getUuid() },
        payload: payload,
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(urlApi, options);
      Logger.log('Creating new task | API response code: ' + response.getResponseCode() + ' | response body: ' + response.getContentText());
      if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
        var createdTask = JSON.parse(response.getContentText());
        var newTaskId = createdTask.id;
        Logger.log('Created new task with ID: ' + newTaskId);
        // Write new Task ID to column A
        sheet.getRange(i + 1, 1).setValue(newTaskId);
        // Clear Last Modified timestamp
        sheet.getRange(i + 1, lastModifiedCol + 1).setValue('');
        // Add to update log
        updatedRowsLog.push({
          row: i+1,
          newTaskId: newTaskId,
          content: taskContent,
          projectId: projectId,
          dueDate: dueDate,
          dueTime: dueTime,
          labels: labels
        });
      } else {
        Logger.log('Failed to create new task for row ' + (i+1));
      }
      continue; // Skip to next row
    }
    // === EXISTING TASK UPDATE ===
    if (!lastModified) continue;
    var lastModifiedDate = new Date(lastModified);
    if (lastSyncDate && lastModifiedDate <= lastSyncDate) {
      Logger.log('Row ' + (i+1) + ' skipped (not modified since last sync)');
      continue;
    }
    Logger.log('Row ' + (i+1) + ' will be updated');

    taskId = row[0];      // ID column (A)
    projectName = row[3];   // Project Name column (D)
    dueDate = row[5];     // Due Date column (F)
    Logger.log('TaskId: ' + taskId + ' | Raw Due Date value: ' + dueDate);
    if (dueDate instanceof Date) {
      dueDate = dueDate.getFullYear() + '-' +
                ('0' + (dueDate.getMonth() + 1)).slice(-2) + '-' +
                ('0' + dueDate.getDate()).slice(-2);
      Logger.log('TaskId: ' + taskId + ' | Formatted Due Date: ' + dueDate);
    }
    Logger.log('TaskId: ' + taskId + ' | Raw Due Time value: ' + row[6]);
    dueTime = formatDueTimeCell(row[6]);     // Due Time column (G)
    Logger.log('TaskId: ' + taskId + ' | Formatted Due Time: ' + dueTime);
    var uiPriority = row[7];  // Priority column (H)

    // Update project if name changed
    var projectId = null;
    if (projectName && projectName.toString().trim() !== '') {
      projectId = getProjectIdFromName(projectName);
      if (!projectId) {
        Logger.log('TaskId: ' + taskId + ' | Failed to get/create project: ' + projectName);
        continue; // Skip this task if project creation failed
      }
    }
    
    // Fetch current project ID from Todoist
    var getTaskUrl = 'https://api.todoist.com/rest/v2/tasks/' + taskId;
    var getTaskOptions = {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN }
    };
    var getTaskResponse = UrlFetchApp.fetch(getTaskUrl, getTaskOptions);
    var currentTask = JSON.parse(getTaskResponse.getContentText());
    var currentProjectId = currentTask.project_id;
    Logger.log('TaskId: ' + taskId + ' | Current Todoist Project ID: ' + currentProjectId);
    
    var projectIdChanged = false;
    if (projectId && projectId != currentProjectId) {
      Logger.log('TaskId: ' + taskId + ' | Project is changing from ' + currentProjectId + ' to ' + projectId);
      // Step 1: Move project using Sync API
      var syncUrl = 'https://api.todoist.com/sync/v9/sync';
      var syncPayload = JSON.stringify({
        commands: [{
          type: 'item_move',
          uuid: Utilities.getUuid(),
          args: {
            id: taskId,
            project_id: projectId
          }
        }]
      });
      var syncOptions = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN },
        payload: syncPayload,
        muteHttpExceptions: true
      };
      var syncResponse = UrlFetchApp.fetch(syncUrl, syncOptions);
      Logger.log('TaskId: ' + taskId + ' | Sync API move response code: ' + syncResponse.getResponseCode() + ' | response body: ' + syncResponse.getContentText());
      projectIdChanged = true;
      // Optionally, re-fetch the task to confirm project_id
      var verifyTaskResponse = UrlFetchApp.fetch(getTaskUrl, getTaskOptions);
      var verifyTask = JSON.parse(verifyTaskResponse.getContentText());
      Logger.log('TaskId: ' + taskId + ' | Project ID after Sync API move: ' + verifyTask.project_id);
    }
    
    // Step 2: Update other fields (if any) using REST API
    var payloadObj = {};
    
    // Update task content (name) and link
    var taskContent = row[1]; // Task (B)
    var taskLink = row[2];    // Link (C)
    var finalContent = taskContent;
    if (taskLink && taskLink.toString().trim() !== '') {
      finalContent = formatUrlToMarkdown(taskLink, taskContent);
    }
    
    payloadObj.content = finalContent;
    
    // Label handling using header-based indexing
    var label1Col = headers.indexOf('Label1');
    var label2Col = headers.indexOf('Label2');
    var label3Col = headers.indexOf('Label3');
    var labelCols = [label1Col, label2Col, label3Col];
    var labels = [];
    labelCols.forEach(function(colIdx) {
      if (row[colIdx] !== undefined && row[colIdx] !== null && row[colIdx].toString().trim() !== '') {
        labels.push(row[colIdx]);
      }
    });
    if (labels.length > 0) {
      payloadObj.labels = labels;
    }

    // First, update just the due date/recurrence if needed
    var recurrenceCell = row[recurringCol] ? row[recurringCol].toString().trim() : '';
    var dueDatePayload = {};
    if (recurrenceCell && recurrenceCell.toLowerCase().startsWith('every')) {
      // If recurrence, build due_string (append time if present and not already in string)
      var dueString = recurrenceCell;
      if (dueTime && dueTime !== '' && !dueString.match(/\d{1,2}:\d{2}/)) {
        dueString += ' at ' + dueTime;
      }
      dueDatePayload.due_string = dueString;
      Logger.log('TaskId: ' + taskId + ' | Sending only due_string for recurrence: ' + dueString);
    } else if (dueDate && dueDate.toString().trim() !== '') {
      // No recurrence, use due_date/due_datetime
      dueDatePayload.due_date = dueDate;
      if (dueTime && dueTime !== '') {
        var offset = getCentralOffset(dueDate);
        var timeString = dueTime.length === 5 ? dueTime : ('0' + dueTime).slice(-5);
        var isoString = dueDate + 'T' + timeString + ':00' + offset;
        dueDatePayload.due_datetime = isoString;
        delete dueDatePayload.due_date;
      }
      Logger.log('TaskId: ' + taskId + ' | Sending due_date/due_datetime: ' + JSON.stringify(dueDatePayload));
    }
    if (Object.keys(dueDatePayload).length > 0) {
      var dueDateUrl = 'https://api.todoist.com/rest/v2/tasks/' + taskId;
      var dueDateOptions = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN, 'X-Request-Id': Utilities.getUuid() },
        payload: JSON.stringify(dueDatePayload),
        muteHttpExceptions: true
      };
      Logger.log('TaskId: ' + taskId + ' | Due date/time/recurrence payload: ' + JSON.stringify(dueDatePayload));
      var dueDateResponse = UrlFetchApp.fetch(dueDateUrl, dueDateOptions);
      Logger.log('TaskId: ' + taskId + ' | Due date update response code: ' + dueDateResponse.getResponseCode() + ' | response body: ' + dueDateResponse.getContentText());
    }

    // Then update other fields if needed
    var otherFieldsPayload = {};
    if (finalContent !== taskContent) {
      otherFieldsPayload.content = finalContent;
    }
    if (labels.length > 0) {
      otherFieldsPayload.labels = labels;
    }
    if (uiPriority && uiPriority.toString().trim() !== '') {
      otherFieldsPayload.priority = 5 - Number(uiPriority);
    }

    if (Object.keys(otherFieldsPayload).length > 0) {
      var urlApi = 'https://api.todoist.com/rest/v2/tasks/' + taskId;
      var options = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN, 'X-Request-Id': Utilities.getUuid() },
        payload: JSON.stringify(otherFieldsPayload),
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(urlApi, options);
      Logger.log('TaskId: ' + taskId + ' | Other fields update response code: ' + response.getResponseCode() + ' | response body: ' + response.getContentText());
    }
    // Add to update log
    updatedRowsLog.push({
      row: i+1,
      taskId: taskId,
      dueDate: dueDate,
      dueTime: dueTime,
      priority: uiPriority,
      labels: labels,
      projectId: projectId,
      projectIdChanged: projectIdChanged
    });
    // Clear the Last Modified timestamp after successful update
    sheet.getRange(i + 1, lastModifiedCol + 1).setValue('');
  }
  // Delete fully completed (non-recurring) tasks, from bottom up
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort(function(a, b) { return b - a; });
    Logger.log('Rows deleted: ' + JSON.stringify(rowsToDelete));
    rowsToDelete.forEach(function(rowIdx) {
      sheet.deleteRow(rowIdx);
    });
  }
  // Only resort if a recurring task was updated
  if (recurringTaskUpdated) {
    sortTasks();
  }
  // Update last sync time
  var now = new Date().toISOString();
  PropertiesService.getDocumentProperties().setProperty('lastSyncTime', now);
  Logger.log('Updated last sync time to: ' + now);
  // Log summary of updated rows
  Logger.log('Summary of updated rows: ' + JSON.stringify(updatedRowsLog, null, 2));
}

// === TASK SORTING FUNCTION ===
function sortTasks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var tasks = data.slice(1); // Remove header row
  
  // Get column indices
  var idCol = headers.indexOf('ID');
  var taskCol = headers.indexOf('Task');
  var projectNameCol = headers.indexOf('Project Name');
  var projectIdCol = headers.indexOf('Project ID');
  var dueDateCol = headers.indexOf('Due Date');
  var dueTimeCol = headers.indexOf('Due Time');
  
  // Custom project order
  var projectOrder = {
    'Inbox': 0,
    'Focus': 1,
    'MainTasks': 2,
    'RecurringTasks': 3,
    'Sveta Tasks': 4,
    'Waiting': 5,
    'Backburner': 6,
    'Reference': 7
  };
  
  // Get current time in Central Time
  var now = new Date();
  var sevenDaysFromNow = new Date(now.getTime() + (7 * 24 * 60 * 60 * 1000));
  
  // Helper function to get task due datetime
  function getTaskDueDateTime(task) {
    var dueDate = task[dueDateCol];
    var dueTime = task[dueTimeCol];
    if (!dueDate) return null;
    var date = new Date(dueDate);
    if (!dueTime) return date;
    if (typeof dueTime === 'string') {
      var parts = dueTime.split(':');
      if (parts.length === 2) {
        var hours = parseInt(parts[0], 10);
        var minutes = parseInt(parts[1], 10);
        if (!isNaN(hours) && !isNaN(minutes)) {
          date.setHours(hours, minutes, 0, 0);
          return date;
        }
      }
      // If string but not HH:MM, ignore time
      return date;
    } else if (typeof dueTime === 'number') {
      // Google Sheets may store time as a fraction of a day
      var totalMinutes = Math.round(dueTime * 24 * 60);
      var hours = Math.floor(totalMinutes / 60);
      var minutes = totalMinutes % 60;
      date.setHours(hours, minutes, 0, 0);
      return date;
    } else if (dueTime instanceof Date) {
      date.setHours(dueTime.getHours(), dueTime.getMinutes(), 0, 0);
      return date;
    }
    // If not recognized, just return date
    return date;
  }
  
  // Helper function to get project order
  function getProjectOrder(projectName) {
    return projectOrder[projectName] !== undefined ? projectOrder[projectName] : 999;
  }
  
  // Sort tasks into groups
  var overdueTasks = [];
  var upcomingTasks = [];
  var otherTasks = [];
  
  tasks.forEach(function(task) {
    var dueDateTime = getTaskDueDateTime(task);
    if (!dueDateTime) {
      otherTasks.push(task);
      return;
    }
    
    if (dueDateTime < now) {
      overdueTasks.push(task);
    } else if (dueDateTime <= sevenDaysFromNow) {
      upcomingTasks.push(task);
    } else {
      otherTasks.push(task);
    }
  });
  
  // Sort function for tasks with due dates
  function sortByDueDate(a, b) {
    var aDue = getTaskDueDateTime(a);
    var bDue = getTaskDueDateTime(b);
    
    // First by due date
    if (aDue.getTime() !== bDue.getTime()) {
      return aDue.getTime() - bDue.getTime();
    }
    
    // Then by project
    var aProject = a[projectNameCol];
    var bProject = b[projectNameCol];
    return getProjectOrder(aProject) - getProjectOrder(bProject);
  }
  
  // Sort function for tasks without due dates or with due dates > 7 days out
  function sortByProjectDueDateTimeName(a, b) {
    var aProject = a[projectNameCol];
    var bProject = b[projectNameCol];
    var projectDiff = getProjectOrder(aProject) - getProjectOrder(bProject);
    if (projectDiff !== 0) return projectDiff;

    var aDue = getTaskDueDateTime(a);
    var bDue = getTaskDueDateTime(b);
    var aHasDue = !!aDue && a[dueDateCol];
    var bHasDue = !!bDue && b[dueDateCol];
    if (aHasDue && bHasDue) {
      if (aDue.getTime() !== bDue.getTime()) {
        return aDue.getTime() - bDue.getTime();
      }
    } else if (aHasDue && !bHasDue) {
      return -1;
    } else if (!aHasDue && bHasDue) {
      return 1;
    }
    // If both have no due date or same due date/time, sort alphabetically by task name
    return a[taskCol].localeCompare(b[taskCol]);
  }
  
  // Sort each group
  overdueTasks.sort(sortByDueDate);
  upcomingTasks.sort(sortByDueDate);
  otherTasks.sort(sortByProjectDueDateTimeName);
  
  // Combine all tasks
  var sortedTasks = [...overdueTasks, ...upcomingTasks, ...otherTasks];
  
  // Write back to sheet
  var output = [headers, ...sortedTasks];
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  
  return {
    overdue: overdueTasks.length,
    upcoming: upcomingTasks.length,
    other: otherTasks.length
  };
}

// === APPLY COLOR-CODED FORMATTING TO TASK LIST ===
function applyTaskConditionalFormatting() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('No data rows found.');
    return;
  }
  var headers = data[0];
  Logger.log('Headers: ' + JSON.stringify(headers));
  var lastRow = data.length;
  var lastCol = data[0].length;
  Logger.log('Last row: ' + lastRow + ', Last col: ' + lastCol);

  // Find column indices
  var dueDateCol = headers.indexOf('Due Date') + 1; // 1-based
  var dueTimeCol = headers.indexOf('Due Time') + 1;
  var priorityCol = headers.indexOf('Priority') + 1;
  var completedCol = headers.indexOf('Completed') + 1;
  var idCol = headers.indexOf('ID') + 1;
  var projectNameCol = headers.indexOf('Project Name') + 1;
  var recurringCol = headers.indexOf('Recurring') + 1;
  var label1Col = headers.indexOf('Label1') + 1;
  Logger.log('Due Date col: ' + dueDateCol + ', Due Time col: ' + dueTimeCol + ', Priority col: ' + priorityCol + ', Completed col: ' + completedCol);

  var dueDateLetter = columnToLetter(dueDateCol);
  var priorityLetter = columnToLetter(priorityCol);
  var completedLetter = columnToLetter(completedCol);
  var idLetter = columnToLetter(idCol);

  var range = sheet.getRange(2, 1, lastRow - 1, lastCol); // Exclude header
  Logger.log('Applying formatting to range: ' + range.getA1Notation());
  range.setBackground(null).setFontColor(null).setFontWeight('normal').setFontLine('none'); // Reset

  // Format header row (row 1)
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#595959') // Dark gray 2
    .setFontColor('#ffffff') // White
    .setHorizontalAlignment('center');

  // Auto-resize all columns to fit data, except TaskLink
  var taskLinkCol = headers.indexOf('TaskLink') + 1;
  for (var col = 1; col <= lastCol; col++) {
    if (col !== taskLinkCol) {
      sheet.autoResizeColumn(col);
    }
  }

  // Center contents for specified columns
  var columnsToCenter = [projectNameCol, dueDateCol, dueTimeCol, recurringCol, priorityCol, label1Col].filter(function(idx) { return idx > 0; });
  columnsToCenter.forEach(function(colIdx) {
    sheet.getRange(2, colIdx, lastRow - 1, 1).setHorizontalAlignment('center');
  });

  var rules = [];

  // Fully completed (non-recurring, ID is blank and Completed is checked): darker gray, white text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + completedLetter + '2=TRUE, $' + idLetter + '2="")')
    .setBackground('#888888')
    .setFontColor('#ffffff')
    .setRanges([range])
    .build());

  // Completed: Gray background
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + completedLetter + '2=TRUE')
    .setBackground('#cccccc')
    .setFontColor('#666666')
    .setRanges([range])
    .build());

  // Overdue: Red background, white bold text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + dueDateLetter + '2 <> "", $' + dueDateLetter + '2 < TODAY())')
    .setBackground('#f4cccc')
    .setFontColor('#990000')
    .setBold(true)
    .setRanges([range])
    .build());

  // Due Today: Orange/Yellow background, bold text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + dueDateLetter + '2 = TODAY()')
    .setBackground('#ffe599')
    .setFontColor('#b45f06')
    .setBold(true)
    .setRanges([range])
    .build());

  // Due in next 7 days: Green background, bold text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + dueDateLetter + '2 > TODAY(), $' + dueDateLetter + '2 <= TODAY()+7)')
    .setBackground('#d9ead3')
    .setFontColor('#274e13')
    .setBold(true)
    .setRanges([range])
    .build());

  // No Due Date: Light gray background
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + dueDateLetter + '2 = ""')
    .setBackground('#eeeeee')
    .setFontColor('#666666')
    .setRanges([range])
    .build());

  // Priority 1: Red text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + priorityLetter + '2 = 1')
    .setFontColor('#d80000')
    .setBold(true)
    .setRanges([range])
    .build());

  // Priority 2: Orange text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + priorityLetter + '2 = 2')
    .setFontColor('#e69138')
    .setBold(true)
    .setRanges([range])
    .build());

  // Priority 3: Blue text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + priorityLetter + '2 = 3')
    .setFontColor('#3c78d8')
    .setBold(true)
    .setRanges([range])
    .build());

  // Priority 4: Gray text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + priorityLetter + '2 = 4')
    .setFontColor('#999999')
    .setRanges([range])
    .build());

  sheet.setConditionalFormatRules(rules);
}

// Helper function to convert column number to letter
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// === HIDE ID COLUMNS FUNCTION ===
function hideIdColumns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Trim header values to avoid issues with extra spaces
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.trim(); });
  var taskIdCol = headers.indexOf('ID') + 1; // 1-based index
  var projectIdCol = headers.indexOf('Project ID') + 1; // 1-based index

  Logger.log('Headers: ' + JSON.stringify(headers));
  Logger.log('Task ID Col: ' + taskIdCol + ', Project ID Col: ' + projectIdCol);

  if (taskIdCol > 0) {
    sheet.hideColumn(sheet.getRange(1, taskIdCol));
  }
  if (projectIdCol > 0) {
    sheet.hideColumn(sheet.getRange(1, projectIdCol));
  }
}