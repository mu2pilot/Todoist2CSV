// === GLOBAL CONFIGURATION ===
var TODOIST_API_TOKEN_1 = 'afbcef9f4c486c11967c841aacda94512e8f85f2'; // Test Account
var TODOIST_API_TOKEN_2 = '0a37e89a4121430b636eb99f20bcd802b5b1ae11'; // Active Account
var TODOIST_API_TOKEN = TODOIST_API_TOKEN_2; // Set which token to use

// === MENU SETUP ===
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Todoist Tools')
    .addItem('Update Sheet from Todoist', 'updateTodoistTasks')
    .addItem('Push Changes to Todoist', 'updateTodoistTaskFromSheet')
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
    ['ID', 'Task', 'TaskLink', 'Project Name', 'Project ID', 'Due Date', 'Due Time', 'Priority', 'Label1', 'Label2', 'Label3', 'Last Modified']
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
      taskLink = '=HYPERLINK("' + url + '", "Link")';
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
      uiPriority,
      labelDict.Label1,
      labelDict.Label2,
      labelDict.Label3,
      '' // Last Modified is blank after refresh
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  // Reset lastSyncTime after refresh
  var now = new Date().toISOString();
  PropertiesService.getDocumentProperties().setProperty('lastSyncTime', now);
  Logger.log('Reset last sync time to: ' + now);
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

// === ON EDIT TRIGGER: Automatically update 'Last Modified' column (L) ===
function onEdit(e) {
  Logger.log('onEdit triggered');
  Logger.log('Range: ' + e.range.getA1Notation());
  Logger.log('Column: ' + e.range.getColumn());
  Logger.log('Row: ' + e.range.getRow());
  
  var sheet = e.range.getSheet();
  var colToWatch = [1,2,3,4,5,6,7,8,9,10,11]; // columns A-K (1-based)
  var lastModifiedCol = 12; // column L (1-based index)
  
  Logger.log('Is column in watch list? ' + (colToWatch.indexOf(e.range.getColumn()) !== -1));
  Logger.log('Is row > 1? ' + (e.range.getRow() > 1));
  
  if (colToWatch.indexOf(e.range.getColumn()) !== -1 && e.range.getRow() > 1) {
    Logger.log('Setting timestamp in column L');
    sheet.getRange(e.range.getRow(), lastModifiedCol).setValue(new Date());
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

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var lastModified = row[11]; // Column L (index 11)
    Logger.log('Row ' + (i+1) + ' | Last Modified: ' + lastModified);
    if (!lastModified) continue;
    var lastModifiedDate = new Date(lastModified);
    if (lastSyncDate && lastModifiedDate <= lastSyncDate) {
      Logger.log('Row ' + (i+1) + ' skipped (not modified since last sync)');
      continue;
    }
    Logger.log('Row ' + (i+1) + ' will be updated');

    var taskId = row[0];      // ID column (A)
    var projectId = row[4];   // Project ID column (E)
    var dueDate = row[5];     // Due Date column (F)
    if (dueDate instanceof Date) {
      dueDate = dueDate.getFullYear() + '-' +
                ('0' + (dueDate.getMonth() + 1)).slice(-2) + '-' +
                ('0' + dueDate.getDate()).slice(-2);
    }
    Logger.log('TaskId: ' + taskId + ' | Raw Due Time value: ' + row[6]);
    var dueTime = formatDueTimeCell(row[6]);     // Due Time column (G)
    Logger.log('TaskId: ' + taskId + ' | Formatted Due Time: ' + dueTime);
    var uiPriority = row[7];  // Priority column (H)

    // Only try to update labels if the columns exist
    var labels = [];
    if (row.length > 8) {
      for (var j = 8; j <= 10 && j < row.length; j++) { // columns I, J, K (Label1, Label2, Label3)
        if (row[j] && row[j].toString().trim() !== '') {
          labels.push(row[j]);
        }
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
    if (projectId && projectId.toString().trim() !== '' && projectId != currentProjectId) {
      Logger.log('TaskId: ' + taskId + ' | Project ID is changing from ' + currentProjectId + ' to ' + projectId);
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
    if (labels.length > 0) {
      payloadObj.labels = labels;
    }
    if (dueDate && dueDate.toString().trim() !== '') {
      if (dueTime && dueTime.toString().trim() !== '') {
        // Use correct US Central offset (DST aware), and use the time as entered in the sheet
        var offset = getCentralOffset(dueDate);
        var timeString = dueTime.length === 5 ? dueTime : ('0' + dueTime).slice(-5); // Ensure HH:MM
        var isoString = dueDate + 'T' + timeString + ':00' + offset;
        Logger.log('Updating TaskId: ' + taskId + ' | due_datetime: ' + isoString);
        payloadObj.due_datetime = isoString;
      } else {
        payloadObj.due_date = dueDate;
      }
    }
    if (uiPriority && uiPriority.toString().trim() !== '') {
      // Convert UI value back to API value
      payloadObj.priority = 5 - Number(uiPriority);
    }
    var payload = JSON.stringify(payloadObj);
    Logger.log('TaskId: ' + taskId + ' | Payload for other fields: ' + payload);
    if (taskId && Object.keys(payloadObj).length > 0) {
      var urlApi = 'https://api.todoist.com/rest/v2/tasks/' + taskId;
      var options = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN, 'X-Request-Id': Utilities.getUuid() },
        payload: payload,
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(urlApi, options);
      Logger.log('TaskId: ' + taskId + ' | API response code (other fields): ' + response.getResponseCode() + ' | response body: ' + response.getContentText());
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
    sheet.getRange(i + 1, 12).setValue('');
  }
  // Update last sync time
  var now = new Date().toISOString();
  PropertiesService.getDocumentProperties().setProperty('lastSyncTime', now);
  Logger.log('Updated last sync time to: ' + now);
  // Log summary of updated rows
  Logger.log('Summary of updated rows: ' + JSON.stringify(updatedRowsLog, null, 2));
}