from jira import JIRA
from jira.exceptions import JIRAError
import smtplib
from email.mime.text import MIMEText

# Connect to Jira
jira = JIRA(
    server='https://jira.domain/',
    basic_auth=('user', 'pass')
)

# Get the counts of each type of ticket
query_project_filter  = ' PROJECT in ("VMDS","Visual Merch DZ Development","DMAS") and issuetype != Epic ' 
days_back  = 8 
query_order_by  = ' ORDER BY resolved DESC,  priority DESC, due ASC, status ASC, project ASC, key ASC ' 
completed_tickets = jira.search_issues(query_project_filter + ' AND  resolved >= -' + str(days_back) + 'd ' + query_order_by)
in_progress_tickets = jira.search_issues(query_project_filter + ' AND (status not in (Completed, Completed, Done, Cancelled, "To Do", Backlog) ) ' + query_order_by)
new_tickets = jira.search_issues(query_project_filter + ' and resolved = null and created >= -' + str(days_back) + 'd ' + query_order_by)
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

import win32com.client as win32
outlook=win32.Dispatch('outlook.application')
mail=outlook.CreateItem(0)
mail.To='youremail@yourdomain.com;'
mail.Subject= 'Projects Status'
mail.Body=report
# mail.HTMLBody=body

mail.Send()



# msg = MIMEText(report, 'html')
# msg['Subject'] = 'Visual Merch Status Report'
# msg['From'] = 'youremail@yourdomain.com'
# msg['To'] = 'youremail@yourdomain.com'

# s = smtplib.SMTP('smtp.gmail.com', 587)
# s.starttls()
# s.login('youremail@yourdomain.com', 'yourpassword')
# s.sendmail('youremail@yourdomain.com', 'recipientemail@recipientdomain.com', msg.as_string())
# s.quit()
