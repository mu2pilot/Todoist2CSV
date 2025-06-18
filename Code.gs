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
    .addSeparator()
    .addItem('Show All Labels', 'showAllRows')
    .addItem('Select Labels...', 'showLabelFilterDialog')
    .addToUi();
}

// === LABEL FILTERING FUNCTIONS ===
function showAllRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.showRows(2, lastRow - 1);
  }
}

function showLabelFilterDialog() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Get column indices for labels
  var label1Col = headers.indexOf('Label1');
  var label2Col = headers.indexOf('Label2');
  var label3Col = headers.indexOf('Label3');
  
  // Get all unique labels from the sheet
  var allLabels = new Set();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[label1Col]) allLabels.add(row[label1Col]);
    if (row[label2Col]) allLabels.add(row[label2Col]);
    if (row[label3Col]) allLabels.add(row[label3Col]);
  }
  
  // Convert Set to Array and sort
  var labelArray = Array.from(allLabels).sort();
  
  // Create HTML for the dialog
  var html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; }
          .checkbox-container { margin-bottom: 10px; }
          .button-container { margin-top: 20px; text-align: center; }
          .submit-btn {
            background-color: #4CAF50;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
          }
          .submit-btn:hover { background-color: #45a049; }
        </style>
      </head>
      <body>
        <form id="labelForm">
          <div class="checkbox-container">
            ${labelArray.map(label => 
              `<div>
                <input type="checkbox" name="labels" value="${label}" id="${label}">
                <label for="${label}">${label}</label>
              </div>`
            ).join('')}
          </div>
          <div class="button-container">
            <input type="submit" value="Apply Filter" class="submit-btn">
          </div>
        </form>
        <script>
          document.getElementById('labelForm').addEventListener('submit', function(e) {
            e.preventDefault();
            var selectedLabels = Array.from(document.querySelectorAll('input[name="labels"]:checked'))
              .map(cb => cb.value);
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .filterBySelectedLabels(selectedLabels);
          });
        </script>
      </body>
    </html>
  `;
  
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(300)
    .setHeight(400)
    .setTitle('Select Labels to Show');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Labels to Show');
}

function filterBySelectedLabels(selectedLabels) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Get column indices for labels
  var label1Col = headers.indexOf('Label1');
  var label2Col = headers.indexOf('Label2');
  var label3Col = headers.indexOf('Label3');
  
  if (selectedLabels.length === 0) {
    showAllRows();
    return;
  }
  
  // Hide all rows first
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.hideRows(2, lastRow - 1);
  }
  
  // Show only rows that have any of the selected labels
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowLabels = [
      row[label1Col],
      row[label2Col],
      row[label3Col]
    ].filter(function(label) { return label; });
    
    // If row has any of the selected labels, show it
    if (rowLabels.some(function(label) { return selectedLabels.includes(label); })) {
      sheet.showRows(i + 1);
    }
  }
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
    ['ID', 'Task', 'Note', 'TaskLink', 'Project Name', 'Project ID', 'Due Date', 'Due Time', 'Recurring', 'Priority', 'Label1', 'Label2', 'Label3', 'Completed', 'Last Modified']
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
      var localDate = new Date(task.due.datetime);
      dueDate = localDate.getFullYear() + '-' +
                ('0' + (localDate.getMonth() + 1)).slice(-2) + '-' +
                ('0' + localDate.getDate()).slice(-2);
      var localHours = localDate.getHours();
      var localMinutes = localDate.getMinutes();
      dueTime = ('0' + localHours).slice(-2) + ':' + ('0' + localMinutes).slice(-2);
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
      task.description || '', // Note column between Task and TaskLink
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
  // Get the local timezone offset in minutes and convert to hours
  var offset = -date.getTimezoneOffset() / 60;
  // Format as ±HH:00
  var sign = offset >= 0 ? '+' : '-';
  var hours = Math.abs(Math.floor(offset));
  return sign + ('0' + hours).slice(-2) + ':00';
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
  Logger.log('formatUrlToMarkdown called with url: "' + url + '", linkText: "' + linkText + '"');
  if (!url || url === 'undefined' || url === 'null') return linkText || '';
  if (url.toString().includes('=HYPERLINK')) {
    var match = url.toString().match(/=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)/);
    if (match) {
      url = match[1];
      linkText = match[2];
    }
  }
  url = url ? url.toString().trim() : '';
  if (!url.match(/^[a-zA-Z]+:\/\//)) {
    url = 'https://' + url;
  }
  var result = '[' + (linkText || url.replace(/\/$/, '')) + '](' + url + ')';
  Logger.log('formatUrlToMarkdown result: "' + result + '"');
  return (typeof result === 'string') ? result : (linkText || '');
}

// === ON EDIT TRIGGER: Automatically update 'Last Modified' column (L) ===
function onEdit(e) {
  Logger.log('onEdit triggered');
  Logger.log('Range: ' + e.range.getA1Notation());
  Logger.log('Column: ' + e.range.getColumn());
  Logger.log('Row: ' + e.range.getRow());
  
  var sheet = e.range.getSheet();
  var colToWatch = [1,2,3,4,5,6,7,8,9,10,11,12,13,14]; // columns A-N (1-based), including Note column
  var lastModifiedCol = 15; // column O (1-based index), shifted due to new Note column
  
  Logger.log('Is column in watch list? ' + (colToWatch.indexOf(e.range.getColumn()) !== -1));
  Logger.log('Is row > 1? ' + (e.range.getRow() > 1));
  
  if (colToWatch.indexOf(e.range.getColumn()) !== -1 && e.range.getRow() > 1) {
    Logger.log('Setting timestamp in column O');
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
  var headers = data[0];
  var lastSyncTime = PropertiesService.getDocumentProperties().getProperty('lastSyncTime');
  var lastSyncDate = lastSyncTime ? new Date(lastSyncTime) : null;
  Logger.log('Last sync time: ' + lastSyncTime);

  var updatedRowsLog = [];
  var idCol = data[0].indexOf('ID'); // 0-based
  var taskNameCol = data[0].indexOf('Task');
  var noteCol = data[0].indexOf('Note'); // Note column should be right after Task
  var taskLinkCol = data[0].indexOf('TaskLink');
  var projectNameCol = data[0].indexOf('Project Name');
  var dueDateCol = data[0].indexOf('Due Date');
  var dueTimeCol = data[0].indexOf('Due Time');
  var recurringCol = data[0].indexOf('Recurring');
  var priorityCol = data[0].indexOf('Priority');
  var completedCol = data[0].indexOf('Completed');
  var lastModifiedCol = data[0].indexOf('Last Modified');
  var label1Col = data[0].indexOf('Label1');
  var label2Col = data[0].indexOf('Label2');
  var label3Col = data[0].indexOf('Label3');

  // Log column indices for debugging
  Logger.log('Column indices: ID=' + idCol + ', Task=' + taskNameCol + ', Note=' + noteCol + 
             ', TaskLink=' + taskLinkCol + ', Project Name=' + projectNameCol + 
             ', Due Date=' + dueDateCol + ', Due Time=' + dueTimeCol + ', Recurring=' + recurringCol + 
             ', Priority=' + priorityCol + ', Completed=' + completedCol + ', Last Modified=' + lastModifiedCol + 
             ', Label1=' + label1Col + ', Label2=' + label2Col + ', Label3=' + label3Col);

  var rowsToDelete = [];
  var recurringTaskUpdated = false;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row.length < headers.length) {
      row = row.concat(Array(headers.length - row.length).fill(''));
    }
    var lastModified = row[lastModifiedCol];
    Logger.log('Headers: ' + JSON.stringify(headers));
    Logger.log('Row ' + (i+1) + ' length: ' + row.length + ', data: ' + JSON.stringify(row));
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
      var taskNameCol = headers.indexOf('Task');
      var taskLinkCol = headers.indexOf('TaskLink');
      var projectNameCol = headers.indexOf('Project Name');
      var dueDateCol = headers.indexOf('Due Date');
      var dueTimeCol = headers.indexOf('Due Time');
      var recurringCol = headers.indexOf('Recurring');
      var priorityCol = headers.indexOf('Priority');
      var completedCol = headers.indexOf('Completed');
      var lastModifiedCol = headers.indexOf('Last Modified');
      var idCol = headers.indexOf('ID');
      var label1Col = headers.indexOf('Label1');
      var label2Col = headers.indexOf('Label2');
      var label3Col = headers.indexOf('Label3');
      var noteCol = headers.indexOf('Note');

      // Log column indices for debugging
      Logger.log('Column indices: Task=' + taskNameCol + ', TaskLink=' + taskLinkCol + ', Project Name=' + projectNameCol + ', Due Date=' + dueDateCol + ', Due Time=' + dueTimeCol + ', Recurring=' + recurringCol + ', Priority=' + priorityCol + ', Completed=' + completedCol + ', Last Modified=' + lastModifiedCol + ', ID=' + idCol + ', Label1=' + label1Col + ', Label2=' + label2Col + ', Label3=' + label3Col + ', Note=' + noteCol);

      // Defensive: Only process if Task Name column exists
      if (taskNameCol < 0) {
        Logger.log('Task Name column not found. Skipping row ' + (i+1));
        continue;
      }
      var taskContentRaw = row[taskNameCol];
      Logger.log('Raw taskContent: ' + taskContentRaw);
      if (
        taskContentRaw === undefined ||
        taskContentRaw === null ||
        (typeof taskContentRaw === 'string' && taskContentRaw.trim() === '') ||
        (typeof taskContentRaw === 'string' && taskContentRaw.trim().toLowerCase() === 'undefined') ||
        (typeof taskContentRaw === 'string' && taskContentRaw.trim().toLowerCase() === 'null')
      ) {
        Logger.log('Skipping row ' + (i+1) + ' because Task Name is blank or invalid: "' + taskContentRaw + '"');
        continue;
      }
      var taskContent = taskContentRaw.toString();
      var taskLink = (taskLinkCol >= 0 && row[taskLinkCol] !== undefined && row[taskLinkCol] !== null && row[taskLinkCol] !== 'undefined' && row[taskLinkCol] !== 'null') ? row[taskLinkCol].toString() : '';
      Logger.log('taskContent: "' + taskContent + '", taskLink: "' + taskLink + '"');
      var finalContent = taskContent;
      if (typeof taskLink === 'string' && taskLink.trim() !== '') {
        finalContent = formatUrlToMarkdown(taskLink, taskContent);
      }
      Logger.log('finalContent: "' + finalContent + '"');
      // If Task Name is blank, log error and skip row
      if (!finalContent || finalContent.trim() === '') {
        Logger.log('ERROR: Task Name is blank for row ' + (i+1) + '. Skipping this row.');
        continue;
      }
      
      var projectName = row[projectNameCol];   // Project Name column (D)
      var projectId = null;
      if (projectName && projectName.toString().trim() !== '') {
        projectId = getProjectIdFromName(projectName);
        if (!projectId) {
          Logger.log('TaskId: ' + taskId + ' | Failed to get/create project: ' + projectName);
          continue; // Skip this task if project creation failed
        }
      }
      
      var dueDate = row[dueDateCol];     // Due Date (F)
      if (dueDate instanceof Date) {
        dueDate = dueDate.getFullYear() + '-' +
                  ('0' + (dueDate.getMonth() + 1)).slice(-2) + '-' +
                  ('0' + dueDate.getDate()).slice(-2);
      }
      var dueTime = formatDueTimeCell(row[dueTimeCol]); // Due Time (G)
      // Label handling using header-based indexing
      var labels = [];
      var labelCols = [label1Col, label2Col, label3Col];
      labelCols.forEach(function(colIdx) {
        if (colIdx >= 0 && row[colIdx] !== undefined && row[colIdx] !== null && row[colIdx].toString().trim() !== '') {
          labels.push(row[colIdx]);
        }
      });
      Logger.log('labelCols: ' + JSON.stringify(labelCols));
      Logger.log('labels: ' + JSON.stringify(labels));
      var taskNote = (noteCol >= 0 && row[noteCol] !== undefined && row[noteCol] !== null) ? row[noteCol].toString() : '';
      Logger.log('Task Note: ' + taskNote);
      // Define payloadObj before assigning any properties
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
          // Use the date and time as-is, no timezone or offset
          var dueDatetime = dueDate + 'T' + dueTime;
          payloadObj.due_datetime = dueDatetime;
        } else {
          payloadObj.due_date = dueDate;
        }
      }
      if (taskNote && taskNote.trim() !== '') {
        payloadObj.description = taskNote;
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
    projectName = row[projectNameCol];   // Project Name column (D)
    dueDate = row[dueDateCol];     // Due Date column (F)
    Logger.log('TaskId: ' + taskId + ' | Raw Due Date value: ' + dueDate);
    if (dueDate instanceof Date) {
      dueDate = dueDate.getFullYear() + '-' +
                ('0' + (dueDate.getMonth() + 1)).slice(-2) + '-' +
                ('0' + dueDate.getDate()).slice(-2);
      Logger.log('TaskId: ' + taskId + ' | Formatted Due Date: ' + dueDate);
    }
    Logger.log('TaskId: ' + taskId + ' | Raw Due Time value: ' + row[dueTimeCol]);
    dueTime = formatDueTimeCell(row[dueTimeCol]);     // Due Time column (G)
    Logger.log('TaskId: ' + taskId + ' | Formatted Due Time: ' + dueTime);
    var uiPriority = row[priorityCol];  // Priority column (H)

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
    
    // Set finalContent for existing task update
    var taskContent = row[taskNameCol];
    var taskLink = (taskLinkCol >= 0 && row[taskLinkCol] !== undefined && row[taskLinkCol] !== null && row[taskLinkCol] !== 'undefined' && row[taskLinkCol] !== 'null') ? row[taskLinkCol].toString() : '';
    var finalContent = (taskContent !== undefined && taskContent !== null) ? taskContent.toString() : '';
    if (typeof taskLink === 'string' && taskLink.trim() !== '') {
      finalContent = formatUrlToMarkdown(taskLink, finalContent);
    }
    Logger.log('finalContent (update block): "' + finalContent + '"');
    // If Task Name is blank, log error and skip row
    if (!finalContent || finalContent.trim() === '') {
      Logger.log('ERROR: Task Name is blank for row ' + (i+1) + '. Skipping this row.');
      continue;
    }

    // Build labels array for update block
    var labels = [];
    var labelCols = [label1Col, label2Col, label3Col];
    labelCols.forEach(function(colIdx) {
      if (colIdx >= 0 && row[colIdx] !== undefined && row[colIdx] !== null && row[colIdx].toString().trim() !== '') {
        labels.push(row[colIdx]);
      }
    });
    Logger.log('labelCols (update block): ' + JSON.stringify(labelCols));
    Logger.log('labels (update block): ' + JSON.stringify(labels));

    // Then update other fields if needed
    var otherFieldsPayload = {};
    // Always include content if task content or link has changed
    if (finalContent !== currentTask.content) {
      otherFieldsPayload.content = finalContent;
      Logger.log('TaskId: ' + taskId + ' | Content changed from "' + currentTask.content + '" to "' + finalContent + '"');
    }

    // Add note/description if changed
    var taskNote = (noteCol >= 0 && row[noteCol] !== undefined && row[noteCol] !== null) ? row[noteCol].toString() : '';
    if (taskNote !== currentTask.description) {
      otherFieldsPayload.description = taskNote;
      Logger.log('TaskId: ' + taskId + ' | Description changed from "' + currentTask.description + '" to "' + taskNote + '"');
    }

    if (Array.isArray(labels) && labels.length > 0) {
      otherFieldsPayload.labels = labels;
    }
    if (uiPriority && uiPriority.toString().trim() !== '') {
      otherFieldsPayload.priority = 5 - Number(uiPriority);
    }

    // Update due date/time if changed
    if (dueDate && dueDate.toString().trim() !== '') {
      if (dueTime && dueTime.toString().trim() !== '') {
        // Use the date and time as-is, no timezone or offset
        var dueDatetime = dueDate + 'T' + dueTime;
        otherFieldsPayload.due_datetime = dueDatetime;
      } else {
        otherFieldsPayload.due_date = dueDate;
      }
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
 // Automatically sort and apply formatting after update
  sortTasks();
  applyTaskConditionalFormatting();
  hideIdColumns();
  
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
  var noteCol = headers.indexOf('Note');
  var taskLinkCol = headers.indexOf('TaskLink');
  var projectNameCol = headers.indexOf('Project Name');
  var projectIdCol = headers.indexOf('Project ID');
  var dueDateCol = headers.indexOf('Due Date');
  var dueTimeCol = headers.indexOf('Due Time');
  var label1Col = headers.indexOf('Label1');

  Logger.log('Sort column indices: ID=' + idCol + ', Task=' + taskCol + ', Note=' + noteCol + 
             ', TaskLink=' + taskLinkCol + ', Project=' + projectNameCol + 
             ', Due Date=' + dueDateCol + ', Time=' + dueTimeCol);

  // Custom project order for Group 4
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
  var todayStr = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2) + '-' + ('0' + now.getDate()).slice(-2);
  var sevenDaysFromNow = new Date(now.getTime() + (7 * 24 * 60 * 60 * 1000));

  // Helper function to get task due date as a string (YYYY-MM-DD)
  function getTaskDueDateString(task) {
    var dueDate = task[dueDateCol];
    if (!dueDate) return '';
    if (dueDate instanceof Date) {
      return dueDate.getFullYear() + '-' + ('0' + (dueDate.getMonth() + 1)).slice(-2) + '-' + ('0' + dueDate.getDate()).slice(-2);
    }
    return dueDate;
  }

  // Helper function to get project order
  function getProjectOrder(projectName) {
    return projectOrder[projectName] !== undefined ? projectOrder[projectName] : 999;
  }

  // Sort function for Groups 1, 2, 3: by due date (date only), then task name
  function sortByDueDateThenTask(a, b) {
    var aDue = getTaskDueDateString(a);
    var bDue = getTaskDueDateString(b);
    if (aDue && bDue && aDue !== bDue) {
      return aDue.localeCompare(bDue);
    }
    // If due dates are equal or missing, sort alphabetically by task name
    var aTask = a[taskCol] || '';
    var bTask = b[taskCol] || '';
    return aTask.localeCompare(bTask);
  }

  // Sort function for Group 4: by project order, then Label1, then due date (date only), then task name
  function sortByProjectLabelDueDateTask(a, b) {
    var aProject = a[projectNameCol];
    var bProject = b[projectNameCol];
    var projectDiff = getProjectOrder(aProject) - getProjectOrder(bProject);
    if (projectDiff !== 0) return projectDiff;
    var aLabel1 = (label1Col >= 0 && a[label1Col]) ? a[label1Col] : '';
    var bLabel1 = (label1Col >= 0 && b[label1Col]) ? b[label1Col] : '';
    var labelDiff = aLabel1.localeCompare(bLabel1);
    if (labelDiff !== 0) return labelDiff;
    var aDue = getTaskDueDateString(a);
    var bDue = getTaskDueDateString(b);
    if (aDue && bDue && aDue !== bDue) {
      return aDue.localeCompare(bDue);
    }
    var aTask = a[taskCol] || '';
    var bTask = b[taskCol] || '';
    return aTask.localeCompare(bTask);
  }

  // Group tasks
  var overdueTasks = [];
  var todayTasks = [];
  var upcomingTasks = [];
  var otherTasks = [];

  // Get today's date string and 7 days from today string
  var todayDate = new Date();
  var todayDateStr = todayDate.getFullYear() + '-' + ('0' + (todayDate.getMonth() + 1)).slice(-2) + '-' + ('0' + todayDate.getDate()).slice(-2);
  var sevenDaysFromNowDate = new Date(todayDate.getTime() + (7 * 24 * 60 * 60 * 1000));
  var sevenDaysFromNowStr = sevenDaysFromNowDate.getFullYear() + '-' + ('0' + (sevenDaysFromNowDate.getMonth() + 1)).slice(-2) + '-' + ('0' + sevenDaysFromNowDate.getDate()).slice(-2);

  tasks.forEach(function(task) {
    var dueDateStr = getTaskDueDateString(task);
    var projectName = task[projectNameCol];
    var isTodayProject = (projectName && projectName.trim().toLowerCase() === 'today');
    var isDueToday = (dueDateStr && dueDateStr === todayDateStr);
    if (dueDateStr && dueDateStr < todayDateStr) {
      overdueTasks.push(task);
    } else if (isTodayProject || isDueToday) {
      todayTasks.push(task);
    } else if (dueDateStr && dueDateStr > todayDateStr && dueDateStr <= sevenDaysFromNowStr) {
      upcomingTasks.push(task);
    } else {
      otherTasks.push(task);
    }
  });

  overdueTasks.sort(sortByDueDateThenTask);
  todayTasks.sort(sortByDueDateThenTask);
  upcomingTasks.sort(sortByDueDateThenTask);
  otherTasks.sort(sortByProjectLabelDueDateTask);

  // Combine all tasks in the new order
  var sortedTasks = [...overdueTasks, ...todayTasks, ...upcomingTasks, ...otherTasks];

  // Write back to sheet
  var output = [headers, ...sortedTasks];
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  return {
    overdue: overdueTasks.length,
    today: todayTasks.length,
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
  var taskCol = headers.indexOf('Task') + 1;
  var noteCol = headers.indexOf('Note') + 1;
  var taskLinkCol = headers.indexOf('TaskLink') + 1;
  var projectNameCol = headers.indexOf('Project Name') + 1;
  var recurringCol = headers.indexOf('Recurring') + 1;
  var label1Col = headers.indexOf('Label1') + 1;

  Logger.log('Formatting column indices: ID=' + idCol + ', Task=' + taskCol + ', Note=' + noteCol + 
             ', TaskLink=' + taskLinkCol + ', Project=' + projectNameCol + 
             ', Due Date=' + dueDateCol + ', Time=' + dueTimeCol);

  var dueDateLetter = columnToLetter(dueDateCol);
  var priorityLetter = columnToLetter(priorityCol);
  var completedLetter = columnToLetter(completedCol);
  var idLetter = columnToLetter(idCol);
  var projectNameLetter = columnToLetter(projectNameCol);

  var range = sheet.getRange(2, 1, lastRow - 1, lastCol); // Exclude header
  Logger.log('Applying formatting to range: ' + range.getA1Notation());
  range.setBackground(null).setFontColor(null).setFontWeight('normal').setFontLine('none'); // Reset

  // Format header row (row 1)
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#595959') // Dark gray 2
    .setFontColor('#ffffff') // White
    .setHorizontalAlignment('center');

  // Auto-resize all columns to fit data, except TaskLink and Note
  for (var col = 1; col <= lastCol; col++) {
    if (col !== taskLinkCol && col !== noteCol) {
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

  // Overdue: Red background, white bold text (Group 1)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + dueDateLetter + '2 <> "", $' + dueDateLetter + '2 < TODAY())')
    .setBackground('#f4cccc')
    .setFontColor('#990000')
    .setBold(true)
    .setRanges([range])
    .build());

  // Today: Green background, bold text, includes Project Name 'Today' (Group 2)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($' + dueDateLetter + '2 = TODAY(), LOWER($' + projectNameLetter + '2) = "today")')
    .setBackground('#d9ead3')
    .setFontColor('#274e13')
    .setBold(true)
    .setRanges([range])
    .build());

  // Upcoming: Yellow background, bold text (Group 3)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + dueDateLetter + '2 > TODAY(), $' + dueDateLetter + '2 <= TODAY()+7)')
    .setBackground('#ffe599')
    .setFontColor('#b45f06')
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

  Logger.log('Hide columns: Task ID Col: ' + taskIdCol + ', Project ID Col: ' + projectIdCol);

  if (taskIdCol > 0) {
    sheet.hideColumn(sheet.getRange(1, taskIdCol));
  }
  if (projectIdCol > 0) {
    sheet.hideColumn(sheet.getRange(1, projectIdCol));
  }
}