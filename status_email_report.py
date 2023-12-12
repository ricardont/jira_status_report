from jira import JIRA
from jira.exceptions import JIRAError
import smtplib
from email.mime.text import MIMEText
import yaml
# get secret credentials
with open('credentials.yaml', 'r') as file:
    credentials = yaml.safe_load(file)
# get jira query global params
with open('jira_query_params.yaml', 'r') as file:
    jira_query_params = yaml.safe_load(file)

domain = credentials['jira']['domain']
user = credentials['jira']['user']
password = credentials['jira']['pass']
try:
    jira = JIRA(
        server=domain,
        basic_auth=(user, password)
    )
except JIRAError as e:
   print(e.status_code, e.text)

# Get the counts of each type of ticket
# projects_names  = "VMDS","VMDD","DMAS"
projects_names  = '"VMDS","VMDD","DMAS"'
area = 'Visual Merch'  
if area == 'Merch Analysis' :
    projects_names += '"DMAS"' 
elif area == 'Visual Merch':
    projects_names = 'VMDS,VMDD'
else:
    area = 'Merch'
      
projects  = '( PROJECT in ( ' + projects_names + ') and issuetype != Epic and status != "Cancelled" ) ' 
days_back  = jira_query_params['days_back']
not_in_progress_status  = jira_query_params['not_in_progress_status']
order_by = jira_query_params['order_by']
completed_tickets = jira.search_issues(projects + ' AND  resolved >= -' + str(days_back) + 'd ' + order_by)
in_progress_tickets = jira.search_issues(projects + ' AND (status not in ( ' + not_in_progress_status + ') ) ' + order_by)
new_tickets = jira.search_issues(projects + ' and resolved = null and created >= -' + str(days_back) + 'd ' + order_by)
completed_count = len(completed_tickets)
in_progress_count = len(in_progress_tickets)
new_count = len(new_tickets)

# Create the report
report = f'''
<h3>Tasks Status within the last {str(days_back)} days</h3>
<table>
    <tr align="center">
        <th>Completed</th>
        <th>In Progress</th>
        <th>New</th>
    </tr>
    <tr style="font-size:40px" >
        <td align="center" width="150" >{completed_count}</td>
        <td align="center" width="150" >{in_progress_count}</td>
        <td align="center" width="150" >{new_count}</td>
    </tr>
</table>
'''

# Add the Completed tickets to the report
report += '''
    <ul>
    <h2>Completed</h2>
    </ul>
'''
for ticket in completed_tickets:
    report += f'<li>{ticket.key}: {ticket.fields.summary}</li>'

# Add the In Progress tickets to the report
report += '''
    <ul>
    <h2>In Progress</h2>
    </ul>
'''
for ticket in in_progress_tickets:
    report += f'<li>{ticket.key}: {ticket.fields.summary}</li>'

# Add the New tickets to the report
report += '''
    <ul>
    <h2>New</h2>
    </ul>
'''
for ticket in new_tickets:
    report += f'<li>{ticket.key}: {ticket.fields.summary}</li>'
    
# Email the report
recipient    = credentials['mail']['recipient']
sender       = credentials['mail']['sender']
mail_host    = credentials['mail']['host']
mail_port    = credentials['mail']['port']
subject = area + ' Projects Status'

import win32com.client as win32
outlook=win32.Dispatch('outlook.application')
mail=outlook.CreateItem(0)
mail.To=recipient
mail.Subject= subject
mail.HTMLBody=report

mail.Send()