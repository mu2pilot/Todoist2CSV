import msal
import requests
import re

# === CONFIGURATION ===
CLIENT_ID = 'YOUR_CLIENT_ID'
CLIENT_SECRET = 'YOUR_CLIENT_SECRET'
TENANT_ID = 'YOUR_TENANT_ID'
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# SharePoint/OneDrive info
SITE_ID = 'YOUR_SITE_ID'
ITEM_ID = 'YOUR_ITEM_ID'  # The Excel file's item ID

# === AUTHENTICATE ===
app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)
result = app.acquire_token_for_client(scopes=SCOPE)
access_token = result['access_token']

# === FETCH TODOIST TASKS ===
API_TOKEN = 'YOUR_TODOIST_API_TOKEN'
headers = {"Authorization": f"Bearer {API_TOKEN}"}
response = requests.get("https://api.todoist.com/rest/v2/tasks", headers=headers)
tasks = response.json()

# === PREPARE DATA ===
rows = [
    ["ID", "Task", "Project Name", "Due Date", "Priority", "Label1", "Label2", "Label3"]
]
for task in tasks:
    content = task.get("content", "")
    visible_text = re.sub(r'\[(.*?)\]\((.*?)\)', r'\1', content)
    match = re.search(r'\[(.*?)\]\((.*?)\)', content)
    if match:
        url = match.group(2)
        task_field = f'=HYPERLINK("{url}", "{visible_text}")'
    else:
        task_field = visible_text
    # Add your project/label logic here as needed
    rows.append([
        task.get("id"),
        task_field,
        "",  # Project Name (add if needed)
        task.get("due", {}).get("date") if task.get("due") else None,
        task.get("priority"),
        "", "", ""  # Labels (add if needed)
    ])

# === WRITE TO EXCEL VIA GRAPH API ===
headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}
worksheet = 'Sheet1'
range_address = f'A1:H{len(rows)}'
url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/items/{ITEM_ID}/workbook/worksheets/{worksheet}/range(address='{range_address}')"
data = {"values": rows}
response = requests.patch(url, headers=headers, json=data)
print(response.status_code, response.text)