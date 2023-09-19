import requests
from requests.auth import HTTPBasicAuth
import json

def get_epic_link(issue_key):
    # Replace 'your_jira_url' and 'your_issue_key' with the appropriate values
    url = f"https://your_jira_url/rest/api/latest/issue/{issue_key}?fields=your_issue_key"

    # Replace 'your_username' and 'your_password' with your Jira credentials
    response = requests.get(url, auth=('your_username', 'your_api_key'))

    if response.status_code == 200:
        issue_data = response.json()
        epic_link = issue_data['fields']['customfield_10014']
        return epic_link
    else:
        print(f"Failed to retrieve Epic Link for issue {issue_key}")
        return None

# URL to get all fields.
url = "https://your_jira_url/rest/api/2/field"
  
# Create an authentication object, using registered emailID, and, token received.
auth = HTTPBasicAuth('your_username', 'your_api_key')
  
# The Header parameter, should mention, the desired format of data.
headers = {
    "Accept": "application/json"
}

# Create a request object with above parameters.
response = requests.request(
    "GET",
    url,
    headers=headers,
    auth=auth
)
  
# Get all fields, by using the json loads method.
fields = json.loads(response.text)

for field in fields:
    print(f'Field ID: {field["id"]}, Field Name: {field["name"]}')

get_epic_link('EBA-39')