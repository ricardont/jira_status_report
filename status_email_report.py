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
projects_names  = jira_query_params['projects'] 
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
    <h2>Tasks Status within the last {str(days_back)} days</h2>
    <h2>Counts of Each Type of Ticket</h2>
        <li>Completed: {completed_count}</li>
    <ul>
        <li>In Progress: {in_progress_count}</li>
        <li>New: {new_count}</li>
    </ul>
'''

# Add the Completed tickets to the report
report += '''
    <ul>
    <h2>Completed</h2>
    </ul>
'''
for ticket in completed_tickets:
    print(f'<li>{ticket.key}: {ticket.fields.summary}</li>')

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
print(report)


# Email the report

# import win32com.client as win32
# outlook=win32.Dispatch('outlook.application')
# mail=outlook.CreateItem(0)
# mail.To='youremail@yourdomain.com;'
# mail.Subject= 'Projects Status'
# mail.Body=report
# # mail.HTMLBody=body

# mail.Send()



# msg = MIMEText(report, 'html')
# msg['Subject'] = 'Visual Merch Status Report'
# msg['From'] = 'youremail@yourdomain.com'
# msg['To'] = 'youremail@yourdomain.com'

# s = smtplib.SMTP('smtp.gmail.com', 587)
# s.starttls()
# s.login('youremail@yourdomain.com', 'yourpassword')
# s.sendmail('youremail@yourdomain.com', 'recipientemail@recipientdomain.com', msg.as_string())
# s.quit()
