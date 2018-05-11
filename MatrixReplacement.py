'''
	Title:  SAP Security Interactive tool 1.2
	Author:  Joe Friedrich
	License:  MIT
'''

print('Importing libraries')
import csv
import re
import os
from sys import stdin
from collections import namedtuple
from datetime import date
from tkinter import Tk

#------------------Begin Company Class----------------------------------------
class Company:
	def __init__(self, name, folder_path, regular_expression, clients):
		self.name = name
		self.folder_path = folder_path #location of .csv files
		self.role_format = re.compile(regular_expression)
		self.clients = clients
		self.vacation_lookup = self.load_vacations()
		self.region_lookup = self.load_regions()
		self.role_lookup = self.load_roles()
		self.email_lookup = self.load_emails()
		self.single_approver_lookup = self.load_single_approvers()


	def load_regions(self):
		'''   *** Rewrite ***  if we put this into 'load roles', we can take the file
						loading process out and make another function
			This grabs the header contents of the Roles .csv file.
			It returns all entries after the first 4.
			These are the headings that correspond with the companies regions.
		'''
		role_file = open(self.folder_path + '/Roles.csv')
		role_data = csv.reader(role_file)
		header_text = role_data.__next__()
		role_file.close()
		
		return header_text[4:]


	def vacation_check(self, names):
		'''
			Takes a list of strings (names)
			For each name in names, 
				it looks for name in vacation_lookup (list of tuples).
				If the name is there,
					checks the date against the dates in vacation_lookup.
				If both of these check out,
					it replaces name with the name in vacation_lookup.
			Returns a list of strings (checked_names)
		'''	
		checked_names = []
		for name in names:
			for vacation in self.vacation_lookup:
				if name == None: 
					break
				elif name == vacation.standard:
					if vacation.start <= date.today() <= vacation.finish:
						name = vacation.replacement
						break
			checked_names.append(name)
		return checked_names


	def load_approvers(self, row):
		'''
			Takes the row and a blank approvers list.
			Approvers occur row[4] and beyond.
			Fills the approvers list with all approvers in the row 
				when row[2] is Regional.
			Otherwise, only grabs the first approver it finds.
		'''
		approvers = []
		if row[2] == 'Regional':
			approvers = row[4: 4 + len(self.region_lookup)]
		else:
			for approver in row[4:4 + len(self.region_lookup)]:
				if approver != 'N/A':
					approvers.append(approver)
					break
		return self.vacation_check(approvers)

	
	def load_roles(self):
		'''   ***REWRITE***   Consider pulling (load regions into this) ***
			This creates a list of tuples that contain
				the role name.
				the role definition.
				and a list of the role approvers (created by load_approvers).
			This information comes from the Roles.csv file.
		'''
		print('Loading lookup dictionary...')
		
		role_file = open(self.folder_path + '/Roles.csv')
		role_reader = csv.reader(role_file)
		role_data = list(role_reader)
		role_file.close()
		
		role = namedtuple('role',
						'name description approvers')
		role_list = []
		for row in role_data:
			if row[0] in ('', None):
				print('Issue with a row.')   #Can we pull row data?
				break
			role_data = role(row[0], row[1],
						self.load_approvers(row))
			role_list.append(role_data)
		
		return role_list


	def load_emails(self):
		''' 
			Creates a dictionary from the information in the Email.csv file.
		'''
		email_file = open(self.folder_path + '/Email.csv')
		email_reader = csv.reader(email_file)
		email_data = list(email_reader)
		email_file.close()
		
		emails = {}
		for row in email_data:
			emails[row[0]] = row[1]
		
		return emails


	def load_vacations(self):
		''' 
			Creates a list of tuples from the information in the Vacations.csv
			file.
		'''		
		vacation_file = open(self.folder_path + '/Vacations.csv')
		vacation_reader = csv.reader(vacation_file)
		vacation_data = list(vacation_reader)
		vacation_file.close()
		
		garbage = vacation_data.pop(0) #remove header
		date_format = re.compile(r'\d+')
		
		vacation = namedtuple('vacation',
				'standard replacement start finish')
		vacation_list = []
		for row in vacation_data:
			find_date = date_format.findall(row[2])
			start_date = date(int(find_date[2]), 
							int(find_date[0]), 
							int(find_date[1]))
			
			find_date = date_format.findall(row[3])
			finish_date = date(int(find_date[2]), 
							int(find_date[0]), 
							int(find_date[1]))
			
			vacation_tuple = vacation(row[0],
								row[1],
								start_date,
								finish_date)
			vacation_list.append(vacation_tuple)
		return vacation_list


	def load_single_approvers(self):
		'''
			Creates a list of tuples from the information in the Singles.csv
			file.
		'''
		single_file = open(self.folder_path + '/Singles.csv')
		single_reader = csv.reader(single_file)
		single_data = list(single_reader)
		single_file.close()
		
		single_approver = namedtuple('single_Approver',
				'menu_name client client_name approver role_text')
		single_approver_list = []
		for row in single_data:
			single_approver_tuple = single_approver(row[0],
											row[1],
											row[2],
											row[3],
											row[4])
			single_approver_list.append(single_approver_tuple)
		
		title_row = single_approver_list.pop(0) #removes title row
		return single_approver_list


	def find_and_sort_roles(self, user_input, user_region):
		'''
			Takes a list of strings (user_input) and an int user_region.
			It filters each string of user_input through that companies regex.
			[It forces the input string to be all uppercase letters.]
			Takes the results of that filter and 
				tries to find it in that company's role_lookup.
			If it is a Regional role, the lenght of appovers will be > 1.
			
			***Rewrite***  Need to separate this functionality out for 
							greater flexability.
				Collects tuples (name, description, approver).
			Returns the list of tuples sorted by approver name.
		'''
		roles_and_approvers = []
		for line in user_input:
			find_role = self.role_format.search(line.upper())
			role_not_found = True
			if(find_role != None):
				for row in self.role_lookup:
					if find_role.group() == row.name:
						if len(row.approvers) > 1:
							roles_and_approvers.append((row.name,
											row.description, 
											row.approvers[user_region]))
						else:
							roles_and_approvers.append((row.name, 
											row.description, 
											row.approvers[0]))
						role_not_found = False
				if role_not_found:
					print('\nRole ' + find_role.group() + ' was not found.')
		return sorted(roles_and_approvers, key = lambda entry: entry[2])

		
#---------------------------End Company Class---------------------------------

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
			clipboard.clipboard_append('\n' + user_client +
					' -- awaiting approval from ' + current_approver + '\n')
			if current_approver not in approvers_without_email:
				if current_approver in company.email_lookup:
					approver_emails.append(company.email_lookup[
											current_approver])
				else:
					approver_emails.append(current_approver +
											"'s email is missing")
		print(role_tuple[0] + '\t ' + role_tuple[1])
		clipboard.clipboard_append(role_tuple[0] + '\t ' + 
									role_tuple[1] + '\n')
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
		clipboard.clipboard_append('\n' + selected_client[1] + ' -- ' + 
				selected_client[2] + ' -- awaiting approval from ' + current_approver +
				'\n' + selected_client[4] + '\n')

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
	clean_last_separator = -1 * len(email_separator_character)
	
	#make this a list comprehension
	for email in email_list:
		email_output += email + email_separator_character
	
	print('\n' + email_output[:clean_last_separator] + '\n')
	return str('\n' + email_output[:clean_last_separator] + '\n')	

	
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

clipboard = Tk()
clipboard.withdraw() #removes Tk window that is not needed at this time.
#need a way to bring focus back to the cmd window in which this is running

#----------------------------Begin Program------------------------------------
while (True):
	print('\n************Welcome to the SAP Access Request Tool************')
	print('To which company does the user belong?')
	select_company = get_menu(company_names)
	if select_company == 0:
		clipboard.destroy()
		break
	company = list_companies[select_company - 1]
	
	print('\nIn which region does the user work?')
	select_region = get_menu(company.region_lookup)
	if select_region == 0:
		clipboard.destroy()
		break
	region = select_region - 1
	
	print("\nPaste the roles in, hit Ctrl+D (Ctrl+Z and Enter for Windows).")
	requested_roles = get_role_input()
	
	role_tuples = company.find_and_sort_roles(requested_roles, region)
	
	clipboard.clipboard_clear() #Clears the clipboard.  New data coming.
	
	if requested_roles != []:
		print('\nIn which SAP client does the user want the access?')
		select_client = get_menu(company.clients)
		if select_client == 0:
			clipboard.destroy()
			break
		user_client = company.clients[select_client - 1]
	
		email_list, clipboard = parse_role_tuples(role_tuples, company, clipboard)
	else:
		email_list = []
	
	email_list, clipboard = single_approvers(company, email_list, clipboard)
	
	clipboard.clipboard_append(email_format(email_list))
	print('\n**YOUR OUTPUT IS IN THE CLIPBOARD. PASTE IT INTO YOUR TICKET.**')
	pause = input('Press enter to continue.')

	os.system('clear') #clears a terminal/powershell screen
	#os.system('cls')  #clears the cmd screen [Windows]
