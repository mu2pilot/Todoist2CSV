import requests
import csv
import re

print("Script started")

import os
API_TOKEN = os.environ.get("TODOIST_API_TOKEN")
headers = {"Authorization": f"Bearer {API_TOKEN}"}

# Get all projects first
projects_response = requests.get("https://api.todoist.com/rest/v2/projects", headers=headers)
projects = projects_response.json()
project_dict = {project['id']: project['name'] for project in projects}

# Get all active tasks
response = requests.get("https://api.todoist.com/rest/v2/tasks", headers=headers)
tasks = response.json()

def parse_content(content):
    # Pattern to match [text](url)
    pattern = r'\[(.*?)\]\((.*?)\)'
    match = re.search(pattern, content)
    if match:
        return match.group(1), match.group(2)
    return content, None

def parse_labels(labels, max_labels=3):
    # Ensure 'GCFO' is always Label1 if present
    label_dict = {}
    labels = labels.copy()  # Avoid modifying the original list
    if 'GCFO' in labels:
        labels.remove('GCFO')
        ordered_labels = ['GCFO'] + labels
    else:
        ordered_labels = labels
    for i, label in enumerate(ordered_labels[:max_labels], 1):
        label_dict[f'Label{i}'] = label
    # Fill remaining slots with None
    for i in range(len(ordered_labels) + 1, max_labels + 1):
        label_dict[f'Label{i}'] = None
    return label_dict

# Write to CSV
with open("todoist_tasks.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow([
        "ID", "Task", "Project Name", 
        "Due Date", "Priority", "Label1", "Label2", "Label3"
    ])
    for task in tasks:
        content = task.get("content", "")
        # Remove all markdown links, leaving just the visible text
        visible_text = re.sub(r'\[(.*?)\]\((.*?)\)', r'\1', content)
        match = re.search(r'\[(.*?)\]\((.*?)\)', content)
        if match:
            url = match.group(2)
            task_field = f'=HYPERLINK("{url}", "{visible_text}")'
        else:
            task_field = visible_text
        project_id = task.get("project_id")
        project_name = project_dict.get(project_id, "Unknown Project")
        labels = task.get("labels", [])
        label_dict = parse_labels(labels)
        ui_priority = 5 - task.get("priority", 1)
        writer.writerow([
            task.get("id"),
            task_field,
            project_name,
            task.get("due", {}).get("date") if task.get("due") else None,
            ui_priority,
            label_dict['Label1'],
            label_dict['Label2'],
            label_dict['Label3']
        ])
print("Export complete: todoist_tasks.csv")