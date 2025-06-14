import msal
import requests
import re
import webbrowser
import json
import os

# === CONFIGURATION ===
CLIENT_ID = '7285d70c-bee8-4576-b727-8b252e0fa9fa'
# CLIENT_SECRET is not needed for device code flow / PublicClientApplication
TENANT_ID = '62a342d2-a4ba-48a3-bb9a-5078990c7015'
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["Files.ReadWrite.All", "Sites.ReadWrite.All"]

# OneDrive Excel file info
EXCEL_FILE_NAME = 'TodoistTasks.xlsx'  # Change if your file is named differently
WORKSHEET_NAME = 'Sheet1'  # Change if your worksheet is named differently

# Todoist API
TODOIST_API_TOKEN = '0a37e89a4121430b636eb99f20bcd802b5b1ae11'

# === AUTHENTICATE WITH MICROSOFT ===
app = msal.PublicClientApplication(
    CLIENT_ID, authority=AUTHORITY
)

# Try to load token from cache
token_cache = app.get_accounts()
if token_cache:
    result = app.acquire_token_silent(SCOPE, account=token_cache[0])
else:
    # If no cached token, do interactive login
    flow = app.initiate_device_flow(scopes=SCOPE)
    if "user_code" not in flow:
        raise ValueError("Failed to create device flow")
    
    print(flow["message"])
    webbrowser.open(flow["verification_uri"])
    
    result = app.acquire_token_by_device_flow(flow)

if 'access_token' not in result:
    print('Failed to obtain access token. Error details:', result.get('error_description', 'No error description'))
    exit(1)
access_token = result['access_token']

# === FETCH TODOIST TASKS ===
headers = {"Authorization": f"Bearer {TODOIST_API_TOKEN}"}
response = requests.get("https://api.todoist.com/rest/v2/tasks", headers=headers)
tasks = response.json()

# === PREPARE DATA ===
rows = [
    ["ID", "Task", "Due Date", "Priority"]
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
    rows.append([
        task.get("id"),
        task_field,
        task.get("due", {}).get("date") if task.get("due") else None,
        task.get("priority")
    ])

# === GET EXCEL FILE ITEM ID FROM ONEDRIVE ===
headers_graph = {'Authorization': f'Bearer {access_token}'}
file_path = EXCEL_FILE_NAME
url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}"
response = requests.get(url, headers=headers_graph)
if response.status_code != 200:
    print('Could not find the Excel file in OneDrive. Check the file name and location.')
    print('Response status:', response.status_code)
    print('Response body:', response.text)
    print('\nTrying to create the file...')
    
    # Try to create the file if it doesn't exist
    create_url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    create_data = {
        "name": EXCEL_FILE_NAME,
        "file": {},
        "@microsoft.graph.conflictBehavior": "replace"
    }
    create_response = requests.post(create_url, headers=headers_graph, json=create_data)
    if create_response.status_code not in (200, 201):
        print('Failed to create Excel file:', create_response.status_code, create_response.text)
        exit(1)
    item_id = create_response.json()['id']
else:
    item_id = response.json()['id']

# === WRITE TO EXCEL VIA GRAPH API ===
range_address = f'A1:D{len(rows)}'
url = f"https://graph.microsoft.com/v1.0/me/drive/items/{item_id}/workbook/worksheets/{WORKSHEET_NAME}/range(address='{range_address}')"
headers_graph['Content-Type'] = 'application/json'
data = {"values": rows}
response = requests.patch(url, headers=headers_graph, json=data)
print('Excel update response:', response.status_code, response.text) 