
import csv
import win32com.client as win32

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

file = csv.reader(open('csv_test.csv'), delimiter = '\t')

for line in file:	

	split_line = line[0].split(',', 1) 
	split_line[1] = split_line[1].replace('"', '')
	#Creates a new list where the CSV's are seperate strings, and gets rid of the quotations in the second string
	# print(split_line)


	mailItem = olApp.CreateItem(0)
	mailItem.Subject = 'Fixed CSV Reader and Emailer'
	mailItem.BodyFormat = 1
	mailItem.To = split_line[0]
	mailItem.Body = split_line[1]
	#Creates an email that is sent to the address in the first column, and fills the body with what is in the second column

	#It will send from your primary Outlook email, so be careful

	mailItem.Send()
