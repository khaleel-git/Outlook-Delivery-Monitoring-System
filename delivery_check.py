import random
import win32com.client
import datetime
import pymssql

today = datetime.date.today()

'''
Function: Send_Email()
Arguments: 
	table_subjects, email_times, table_expected_times, email_status
'''
def Send_Email(table_subjects, email_times, table_expected_times, email_status, table_owners):
	# Sort all lists based on email_times
	# sorted_lists = sorted(zip(email_times, table_subjects, table_expected_times, email_status, table_owners))
	sorted_lists = sorted(zip(table_expected_times, table_subjects, email_times, email_status, table_owners))


	html_table = "<table>\n"
	html_table += "<tr><th><h3>Email Subject</h3></th><th><h3>Email Time</h3></th><th><h3>Expected Time</h3></th><th><h3>Email Status</h3></th><th><h3>Owner</h3></th></tr>\n"

	# for time, subject, expected_time, status, owner in sorted_lists:
	for expected_time,subject,time,status,owner  in sorted_lists:
		if status == "Delivery Unsuccessful":
			html_table += '<tr class="warning">'
		elif status == "Success":
			html_table += '<tr class="success">'
		elif status == "Late Delivery":
			html_table += '<tr class="success">'
		else:
			html_table += "<tr>"
			
		html_table += f"<td>{subject}</td>"
		html_table += f"<td>{time}</td>"
		html_table += f"<td>{expected_time}</td>"
		html_table += f"<td>{status}</td>"
		html_table += f"<td>{owner}</td>"
		html_table += "</tr>\n"

	html_table += "</table>"

	with open('E:\\Khaleel\\EricMon\\EricMon.css', 'r') as css_file:
		css_content = css_file.read()

	html_head = f"<head><style>{css_content}</style></head>"

	header = f"<h1>Delivery Monitoring</h1>"
	boday =  f'''
					<h2>Delivery Status So far, Time: {datetime.datetime.now().strftime("%H:%M")}</h2>
			  '''

	footer = "<p><font size=""2"" face=""Bahnschrift Light Condensed"" color=""Gray""><br>This is a system generated email, please report anomalies to  <a href=""mailto:ENPMPERFORMANCETEAM@jazz.com.pk"">ENPMPERFORMANCETEAM@jazz.com.pk</a></p></html>"

	HTMLBody = html_head + header + boday + html_table + footer


	obj = win32com.client.Dispatch("Outlook.Application")
	outlook_obj = obj.CreateItem(0)
	outlook_obj.SentOnBehalfOfName = 'Outlook-configured-email@outlook'
	# outlook_obj.To = "khaleel Ahmad <khaleel.org@gmail.com>;"
	outlook_obj.To = ""
	outlook_obj.Subject = f"EricMon Delivery Monitoring | " + today.strftime('%d-%m-%Y')
	outlook_obj.HTMLBody = HTMLBody
	outlook_obj.Send()

'''
End of Function: Send_Email
'''

if __name__ == '__main__':
	# Get delivey Info from EricMon Database -> Table: Meta_Delivery
	conn = pymssql.connect(host="", user="", password="", database="EricMon")
	cur = conn.cursor()
	query = "SELECT * FROM [EricMon].[dbo].[Meta_Delivery]"
	cur.execute(query)
	rows = cur.fetchall()

	# Outlook configurations
	outlook_app = win32com.client.Dispatch("Outlook.Application")
	namespace = outlook_app.GetNamespace("MAPI")
	sent_folder = namespace.GetDefaultFolder(5)  # 5 represents the index of the sent folder

	# Filter only today's Sent Emails
	filter_str = f"[SentOn] >= '{today.strftime('%m/%d/%Y')} 00:00 AM' AND [SentOn] <= '{today.strftime('%m/%d/%Y')} 11:59 PM'"
	sent_emails = sent_folder.Items.Restrict(filter_str)  # Apply the filter

	# Initialization
	table_subjects = []
	table_expected_times = []
	table_owners = []

	email_times = []
	email_status = []

	bot_counts, repeat_counts = 0, 0

	for row in rows:
		table_subject = row[2].rsplit("|", 1)[0].strip()
		expected_delivery_time = row[3]

		table_subjects.append(table_subject)
		table_owners.append(row[5])

		email_found = False  # Flag to check if an email was found for the subject
		for email in sent_emails:
			email_subject = email.Subject.rsplit("|", 1)[0].strip()

			if "::" in email.Subject:
				email_subject = email.Subject.rsplit("::", 1)[0].strip()

			if "||" in email.Subject:
				email_subject = email.Subject.rsplit("||", 1)[0].strip()

			if email_subject.lower() == table_subject.lower():
				email_times.append(email.SentOn.strftime("%H:%M"))
				table_expected_times.append(expected_delivery_time)
				if expected_delivery_time < email.SentOn.strftime("%H:%M"):
					email_status.append('Late Delivery')
				else:
					email_status.append('Success')
				email_found = True
				break  # Found the email, exit the loop

		if not email_found:
			current_time = datetime.datetime.now().strftime("%H:%M")
			if current_time > expected_delivery_time:
				email_times.append(' ')
				email_status.append('Delivery Unsuccessful')
			else:
				email_times.append(' ')
				email_status.append('Pending Delivery')
			table_expected_times.append(expected_delivery_time)

	# Execute Send_Email Function
	Send_Email(table_subjects, email_times, table_expected_times, email_status, table_owners)

	# Print the counts of the lists
	print(f"Email Subjects, length: {len(table_subjects)}")
	print(f"Email Times, length: {len(email_times)}")
	print(f"Email Statuses, length: {len(email_status)}")
	print(f"Owners Count, length: {len(table_owners)}")

	# Release resources from RAM
	namespace = None
	outlook_app = None
