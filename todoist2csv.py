import requests
import csv
import re

print("Script started")

API_TOKEN = '0a37e89a4121430b636eb99f20bcd802b5b1ae11'
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

def parse_labels(labels, max_labels=5):
    # Convert labels list to dictionary with Label1, Label2, etc.
    label_dict = {}
    for i, label in enumerate(labels[:max_labels], 1):
        label_dict[f'Label{i}'] = label
    # Fill remaining slots with None
    for i in range(len(labels) + 1, max_labels + 1):
        label_dict[f'Label{i}'] = None
    return label_dict

# Write to CSV
with open("todoist_tasks.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow([
        "ID", "Content", "Content2", "Project Name", 
        "Due Date", "Priority", "Label1", "Label2", "Label3", "Label4", "Label5"
    ])
    for task in tasks:
        content = task.get("content", "")
        content1, content2 = parse_content(content)
        project_id = task.get("project_id")
        project_name = project_dict.get(project_id, "Unknown Project")
        labels = task.get("labels", [])
        label_dict = parse_labels(labels)
        
        writer.writerow([
            task.get("id"),
            content1,
            content2,
            project_name,
            task.get("due", {}).get("date") if task.get("due") else None,
            task.get("priority"),
            label_dict['Label1'],
            label_dict['Label2'],
            label_dict['Label3'],
            label_dict['Label4'],
            label_dict['Label5']
        ])
print("Export complete: todoist_tasks.csv")