# Jira Tasks Status Friendly Report
This process generates mobile/email friendly summary tasks status report from Jira  
## Requirements
- Python3.*
- Python Libs
    - jira
    - mime
    - yaml
## Setup
### Python env Setup 
## Initial App Config
Create and fill Jira credentials file YAML File ./credentials.yaml
```
jira:
  domain: 'https://jira.domain/'
  user: 'jira_user'
  pass: 'jira_pass'
mail:
  sender: 'sender@sender.com'
  recipient: 'recipient@recipient.com' 
  host: 'smpt host'
  port: 23
```
Fill Jira query standard params YAML file ./jira_query_params.yaml
```
not_in_progress_status: ' Completed, Done, "To Do", Backlog'
days_back: 8
order_by: ' ORDER BY resolved DESC,  priority DESC, due ASC, status ASC, project ASC, key ASC'
```
## Usage
Run from console 
```
python3 status_email_report.py
```
## Pseudocode
- Get secret credentials
- Get jira query global params
- Set Jira Connection 
- Set are Jira projects IDs 
- Get the counts of each type of ticket
- Create the HTML report
- Initiate with Headers and Top Counter Boxes 
- Create Detailed List View within the X days back
- Add the Completed detailed ticket list to the report
- Add the In-Progress detailed ticket list to the report
- Add the New detailed ticket list to the report
- Distribute the report by Email