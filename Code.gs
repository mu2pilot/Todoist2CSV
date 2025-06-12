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
    ['ID', 'TaskRaw', 'Task', 'TaskLink', 'Project Name', 'Project ID', 'Due Date', 'Priority', 'Label1', 'Label2', 'Label3']
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

    // Task: Only the visible text (replace all [text](url) with text)
    var taskVisible = content.replace(/\[(.*?)\]\((.*?)\)/g, '$1');

    // TaskLink: Clickable link for the first [text](url)
    var taskLink = '';
    var match = content.match(/\[(.*?)\]\((.*?)\)/);
    if (match) {
      var url = match[2];
      taskLink = '=HYPERLINK("' + url + '", "Link")';
    }

    // Labels logic
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

    rows.push([
      task.id,
      content,           // TaskRaw: full markdown content
      taskVisible,       // Task: only visible text
      taskLink,          // TaskLink: clickable link for first [text](url)
      projectDict[task.project_id] || '',
      task.project_id,   // Project ID
      task.due ? task.due.date : '',
      task.priority,
      labelDict.Label1,
      labelDict.Label2,
      labelDict.Label3
    ]);
  });

  // Write all rows at once
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

// === PUSH CHANGES TO TODOIST (SAFE, WITH PROJECT MOVE) ===
function updateTodoistTaskFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var taskId = row[0];      // ID column
    var contentText = row[1]; // TaskRaw column (full markdown)
    var projectId = row[5];   // Project ID column (F)
    var dueDate = row[6];     // Due Date column (G)
    var priority = row[7];    // Priority column (H)

    // Only try to update labels if the columns exist
    var labels = [];
    if (row.length > 8) {
      for (var j = 8; j <= 10 && j < row.length; j++) { // columns I, J, K (Label1, Label2, Label3)
        if (row[j] && row[j].toString().trim() !== '') {
          labels.push(row[j]);
        }
      }
    }

    // --- Project Move Logic ---
    if (projectId && projectId.toString().trim() !== '') {
      // Fetch current project ID from Todoist
      var getTaskUrl = 'https://api.todoist.com/rest/v2/tasks/' + taskId;
      var getTaskOptions = {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN }
      };
      var getTaskResponse = UrlFetchApp.fetch(getTaskUrl, getTaskOptions);
      var currentTask = JSON.parse(getTaskResponse.getContentText());
      var currentProjectId = currentTask.project_id;

      // If the project ID in the sheet is different, move the task
      if (projectId != currentProjectId) {
        var moveUrl = 'https://api.todoist.com/rest/v2/tasks/' + taskId + '/move';
        var movePayload = JSON.stringify({ project_id: projectId });
        var moveOptions = {
          method: 'post',
          contentType: 'application/json',
          headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN, 'X-Request-Id': Utilities.getUuid() },
          payload: movePayload,
          muteHttpExceptions: true
        };
        var moveResponse = UrlFetchApp.fetch(moveUrl, moveOptions);
        Logger.log('Moved task ' + taskId + ' to project ' + projectId + ': ' + moveResponse.getContentText());
      }
    }

    // Build payload: only include fields if they are present and non-blank
    var payloadObj = { content: contentText };
    if (labels.length > 0) {
      payloadObj.labels = labels;
    }
    if (dueDate && dueDate.toString().trim() !== '') {
      payloadObj.due_date = dueDate;
    }
    if (priority && priority.toString().trim() !== '') {
      payloadObj.priority = Number(priority);
    }
    // Do NOT include project_id here; it's handled by the move logic above

    var payload = JSON.stringify(payloadObj);

    if (taskId && contentText) {
      var urlApi = 'https://api.todoist.com/rest/v2/tasks/' + taskId;
      var options = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + TODOIST_API_TOKEN, 'X-Request-Id': Utilities.getUuid() },
        payload: payload,
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(urlApi, options);
      Logger.log('Updated task ' + taskId + ': ' + response.getContentText());
    }
  }
}