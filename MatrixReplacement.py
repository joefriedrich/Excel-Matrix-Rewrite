'''
	Title:  SAP Security Interactive tool 1.3
	Author:  Joe Friedrich
	License:  MIT
'''

print('Importing libraries')
import os
import pyperclip
from company import Company
from sys import stdin

def get_menu(list_of_things):
	'''
		This takes a list as input.
		It prints the list to the screen as menu options 1 through len(list).
		It calls and returns the results of get_menu_input.
	'''
	menu_number = 0
	print('--------------------------------')
	for thing in list_of_things:
		menu_number += 1
		print(str(menu_number) +  ':  ' + thing)
	print('--------------------------------')
	return get_menu_input(menu_number)

	
def get_menu_input(menu_total):
	'''
		This takes the final number of possible selections (whole number).
		It takes user input and
			-if the input is between 1 and the last possible selection
				it returns the number the user typed in.
			-if the number is greater than the last possible selection
				or less than 1
				or not a whole number
				or not a number at all
				it returns 0.
	'''
	while True:
		try:
			print('Type a number from the menu and hit enter.')
			user_input = int(input('Make any other entry and enter to stop:  '))
			print('\n')
			if 1 <= user_input <= menu_total:
				return user_input
			else:  #the numerical value is not a menu option
				print('\nAre you sure you are finished?')
				print('Press y and enter to quit.')
				quitting = input('Press any other key and enter to continue:  ')
				if quitting in ['y', 'Y']:
					return 0
		except ValueError:  #if the value is a non-numeral or float
			print('\nAre you sure you are finished?')
			print('Press y and enter to quit.')
			quitting = input('Press any other key and enter to continue:  ')
			if quitting in ['y', 'Y']:
				return 0

				
def get_role_input():
	'''
		This is just a placeholder for user input that is not in the form
			of get_menu_input.
	'''
	return stdin.readlines()

def find_and_sort_roles(company, user_input):
	'''
		Takes a list of strings (user_input) and an int user_region.
		It filters each string of user_input through that companies regex.
		[It forces the input string to be all uppercase letters.]
		Takes the results of that filter and 
			tries to find it in that company's role_lookup.
		If it is a Regional role, the lenght of appovers will be > 1.
			Collects tuples (name, description, approver).
		Returns the list of tuples sorted by approver name.
	'''
	user_region = -1
	roles_and_approvers = []
	for line in user_input:
		find_role = company.role_format.search(line.upper())
		role_not_found = True
		if(find_role != None):
			for row in company.role_lookup:
				if find_role.group() == row.name:
					if len(row.approvers) > 1:
						if user_region < 0:
							print('\nIn which region does the user work?')
							select_region = get_menu(company.region_lookup)
							while select_region == 0:
								print('\nPlease select a valid region.  You cannot quit from here.\n')
								select_region = get_menu(company.region_lookup)
							user_region = select_region - 1
						roles_and_approvers.append((row.name,
										row.description, 
										row.approvers[user_region]))
					else:
						roles_and_approvers.append((row.name, 
										row.description, 
										row.approvers[0]))
					role_not_found = False
			if role_not_found:
				print('\r\nRole ' + find_role.group() + ' was not found.')
	return sorted(roles_and_approvers, key = lambda entry: entry[2])
	
def parse_role_tuples(output, company, clipboard):
	'''
		Takes the list of tuples (output), company object, and the clipboard.
			Output is expected to be sorted by approver (output[2]).
			Clipboard should be blank.
		For every unique approver in the tuples, it will print and append a 
			header to the clipboard.
			It will also append the email address to approver_emails.
		These are followed by the role names until a new approver name is 
			found or the list ends.
		Returns a list of strings (approver_emails) and the clipboard.
	'''
	approvers_without_email = ['Do Not Assign',
					'NO APPROVAL REQUIRED',
					'No Approval Needed',
					'Do Not Assign - Parent Role',
					'Parent Role: ASSIGN ONLY CHILD ROLES FOR THIS ACCESS']
	current_approver = ''
	approver_emails = []
	for role_tuple in output:
		if(role_tuple[2] != current_approver):
			current_approver = role_tuple[2]
			print('\n' + user_client +
					' -- awaiting approval from ' + current_approver)
			clipboard += '\r\n' + user_client + ' -- awaiting approval from ' + current_approver + '\r\n'
			if current_approver not in approvers_without_email:
				if current_approver in company.email_lookup:
					approver_emails.append(company.email_lookup[
											current_approver])
				else:
					approver_emails.append(current_approver +
											"'s email is missing")
		print(role_tuple[0] + '\t ' + role_tuple[1])
		clipboard += role_tuple[0] + '\t ' + role_tuple[1] + '\r\n'
	return approver_emails, clipboard

	
def single_approvers(company, email_list, clipboard):
	'''
		Takes list of tuples, list of strings, and clipboard.
		Will take one single approver at a time and append it's data to 
			email_list and clipboard.
		Returns list of strings (email_list) and clipboard.
	'''
	while True:
		print('\n*********Single Approver Client Section********************')
		print('***This will print your single client approvers UNTIL *****')
		print('*****you make a selection that is NOT on the menu.*********')
		print('***********************************************************')
		print('\nDoes the user need additional '
				'access in single approver clients?')
		menu_options = [client[0] for client in company.single_approver_lookup]

		select_single_client = get_menu(menu_options)
		if select_single_client == 0:
			break
		
		selected_client = company.single_approver_lookup[select_single_client - 1]
		current_approver = selected_client[3]
		
		print('\n' + selected_client[1] + ' -- ' + selected_client[2] +
				' -- awaiting approval from ' + current_approver + '\n' +
				selected_client[4])
		clipboard += '\r\n' + selected_client[1] + ' -- ' + selected_client[2] + ' -- awaiting approval from ' + current_approver + '\r\n' + selected_client[4] + '\r\n'

		#need a check for multiple email approvers
		if current_approver in company.email_lookup: 
			email_list.append(company.email_lookup[current_approver])
		else:
			email_list.append(current_approver + "'s email is missing")

	return email_list, clipboard

	
def email_format(email_list):
	'''
		Takes a list of strings.
		Returns a concatination of this list.
	'''
	email_output = ''
	email_separator_character = ','
	clean_email = -1 * len(email_separator_character)

	for email in email_list:
		email_output += email + email_separator_character
	
	print('\r\n' + email_output[:clean_email] + '\r\n')
	return str('\r\n' + email_output[:clean_email] + '\r\n')	

	
#--------------------------End Local Functions--------------------------------

print('Loading companies.')
company1 = Company('Company1',
			r'/home/joe/Code/github/Excel-Matrix-Rewrite/Company1Data',
			r'[A-Z]{1,3}(:|_)\S+',
			['ProdC1', 'QaC1', 'ProdC1/QaC1', 'DevC1', 'QaC1/DevC1'])
company2 = Company('Company2',
			r'/home/joe/Code/github/Excel-Matrix-Rewrite/Company2Data',
			r'Z:\S{4}:\S{7}:\S{4}:\S',
			['ProdC2', 'QaC2', 'ProdC2/QaC2', 'DevC2', 'QaC1/DevC2'])
company3 = Company('Company3',
			r'/home/joe/Code/github/Excel-Matrix-Rewrite/Company3Data',
			r'\S+',
			['ProdC3', 'QaC3', 'ProdC3/QaC3', 'DevC3', 'QaC3/DevC3'])

list_companies = [company1, company2, company3]
company_names = [company.name for company in list_companies]

#----------------------------Begin Program------------------------------------
while (True):
	print('\n************Welcome to the SAP Access Request Tool************')
	if len(company_names) == 1:
		print('This request is for ' + company_names[0] + '.')
		company = list_companies[0]
	else:
		print('To which company does the user belong?')
		select_company = get_menu(company_names)
		if select_company == 0:
			break
		company = list_companies[select_company - 1]
	
	print("\nPaste the roles in, one per line.")
	print("On a new line, hit Ctrl+Z and Enter to continue.")
	requested_roles = get_role_input()
	
	role_tuples = find_and_sort_roles(company, requested_roles)
	
	clipboard = ''
	pyperclip.copy(clipboard) #Clears the clipboard.  New data coming.
	
	if requested_roles != []:
		print('\nIn which SAP client does the user want the access?')
		select_client = get_menu(company.clients)
		if select_client == 0:
			break
		user_client = company.clients[select_client - 1]
	
		email_list, clipboard = parse_role_tuples(role_tuples, company, clipboard)
	else:
		email_list = []
	
	email_list, clipboard = single_approvers(company, email_list, clipboard)
	
	clipboard += email_format(email_list)
	pyperclip.copy(clipboard)
	print('\n**YOUR OUTPUT IS IN THE CLIPBOARD. PASTE IT INTO YOUR TICKET.**')
	pause = input('Press enter to continue.')

	os.system('clear') #clears a terminal/powershell screen
	#os.system('cls')  #clears the cmd screen [Windows]
