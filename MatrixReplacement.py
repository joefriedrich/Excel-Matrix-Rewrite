'''
	Title:  Excel Matrix tool
	Author:  Joe Friedrich
	License:  MIT
'''

print('Importing libraries')
import openpyxl
import re
from sys import stdin
from warnings import simplefilter
from collections import namedtuple
from datetime import datetime

#-----------------------------------------------------------------------------
class Company:
	def __init__(self, name, matrix, regex, clients):
		self.name = name
		self.matrix = matrix
		self.regex = regex
		self.clients = clients
		self.roles = self.matrix.get_sheet_by_name("Roles")
		self.emails = self.matrix.get_sheet_by_name("Names and Email")
		self.vacations = self.matrix.get_sheet_by_name("OnHoliday")
		self.regions = self.load_regions()
		self.role_lookup = self.load_roles()
		self.email_lookup = self.load_emails()
		self.vacation_lookup = self.load_vacations()
	
	def load_regions(self):
		header_text = []
		for cell in self.roles.rows[0]:
			if cell.value in ('', ' ', None):
				break
			header_text.append(cell.value)
		return header_text[4:]
		'''
			This grabs the contents of the cells in the top row of the Roles worksheet.
			It stores them in a list and stops when it gets to a blank cell.
			It returns all entries after the first 4.
			These are the headings that correspond with the companies regions.
		'''

	def vacation_check(self, names):
		return names #needs implemented
	'''
	'''	

	def load_approvers(self, row):
		approvers = []
		if row[2].value == 'Regional':
			approvers = [approver.value for approver in row[4:4 + len(self.regions)]]
		else:
			for approver in row[4:4 + len(self.regions)]:
				if approver.value != 'N/A':
					approvers.append(approver.value)
					break
		approvers = self.vacation_check(approvers)
		return approvers
		'''
			Takes the row and a blank approvers list.
			Approvers occur row[4] and beyond.
			Fills the approvers list with all approvers in the row 
				when row[2] is Regional.
			Otherwise, only grabs the first approver it finds.
		'''
	
	def load_roles(self):
		print('Loading lookup dictionary...')
		role = namedtuple('role',
						  'name description approvers')
		role_list = []
		for row in self.roles.rows:
			if row[0].value in ('', None):
				print('Issue with row ' + str(row[0]))
				break
			role_tuple = role(row[0].value, row[1].value, self.load_approvers(row))
			role_list.append(role_tuple)
		return role_list
		'''
			This creates a list of tuples that contain
				the role name.
				the role definition.
				and a list of the role approvers (created by load_approvers).
			Region was not included, because we can get that from
				the length of the list of role approvers (when greater than 1).
		'''

	def load_emails(self):
		emails = {}
		for row in self.emails.rows:
			emails[row[0].value] = row[1].value
		return emails
	'''
		Creates a dictionary of emails from the table of names/emails in the sprreadsheet.
	'''

	def load_vacations(self):
		vacation = namedtuple('vacation',
				'standard replacement start finish')
		vacation_list = []
		for row in self.vacations.rows:
			vacation_tuple = vacation(row[0].value,
						row[1].value,
						row[2].value,
						row[3].value)
			vacation_list.append(vacation_tuple)
		return vacation_list
	'''
		Creates a list of tuples from the information in the OnHoliday spreadsheet.
	'''

	def check_and_sort_roles(self, user_input, user_region):
		roles_and_approvers = []
		for line in user_input:
			find_this = self.regex.search(line.upper())
			if(find_this != None):
				for row in self.role_lookup:
					if find_this.group() == row.name:
						if len(row.approvers) > 1:
							roles_and_approvers.append((row.name,
											row.description, 
											row.approvers[user_region]))
						else:
							roles_and_approvers.append((row.name, 
											row.description, 
											row.approvers[0]))
		return sorted(roles_and_approvers, key = lambda entry: entry[2])
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

#-------------------------------------------------------------------------------
def get_menu(list_of_things):
	menu_number = 0
	for thing in list_of_things:
		menu_number += 1
		print(str(menu_number) +  ':  ' + thing)
	return get_menu_input(menu_number)
	'''
		This takes a list as input.
		It prints the list to the screen as menu options 1 through len(list).
		It calls and returns the results of get_menu_input.
	'''

def get_menu_input(menu_total):
	while True: #left while true in incase we need to make it harder to quit
		try:
			user_input = int(input('Any other entry to quit:  '))
			if 1 <= user_input <= menu_total:
				return user_input
			else:
				return 0
		except ValueError:  #if the value is a non-numeral, return 0
			return 0
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
	
def get_role_input():
	return stdin.readlines()
	'''
		This is just a placeholder for user input that is not in the form
			of get_menu_input.
	'''

def output_to_screen_and_clipboard(output, company):
	current_approver = ''
	approver_emails = ''
	for role_tuple in output:
		if(role_tuple[2] != current_approver):
			current_approver = role_tuple[2]
			print('\n' + user_client + ' -- awaiting approval from ' + current_approver)
			if current_approver in company.email_lookup:
				approver_emails += company.email_lookup[current_approver] + ', '
			else:
				approver_emails += current_approver + "'s email is missing, "
		print(role_tuple[0] + '\t' + role_tuple[1])
	print('\n' + approver_emails[:-2] + '\n')
	'''
	'''
	
#-------------------------------------------------------------------------------
#simplefilter('ignore')  #uncomment this line when we go live

print('Loading companies')
company1 = Company('Company1',
			openpyxl.load_workbook(r'/{file path}/matrixCompany1.xlsx'),
			re.compile(r'[A-Z]{1,3}(:|_)\S+'),
			['ProdC1', 'QaC1', 'ProdC1/QaC1', 'DevC1', 'QaC1/DevC1'])
company2 = Company('Company2',
			openpyxl.load_workbook(r'/{file path}/matrixCompany2.xlsx'),
			re.compile(r'Z:\S{4}:\S{7}:\S{4}:\S'),
			['ProdC2', 'QaC2', 'ProdC2/QaC2', 'DevC2', 'QaC1/DevC2'])
company3 = Company('Company3',
			openpyxl.load_workbook(r'/{file path}/matrixCompany3.xlsx'),
			re.compile(r'\S+'),
			['ProdC3', 'QaC3', 'ProdC3/QaC3', 'DevC3', 'QaC3/DevC3'])

list_companies = [company1, company2, company3]

list_company_names = [company.name for company in list_companies]

list_single_approver_clients = [(),
				()] #needs implemented
			
#Begin Program
while (True):	
	print('\nWhich company?')
	select_company = get_menu(list_company_names)
	if (select_company == 0):
		break
	matrix = list_companies[select_company - 1]
	
	print('\nWhich region?')
	select_region = get_menu(matrix.regions)
	if (select_region == 0):
		break
	region = select_region - 1
	
	print("\nPaste the roles in, hit Ctrl+D (Ctrl+Z and Enter for Windows).")
	requested_roles = get_role_input()
	
	organized_output = matrix.check_and_sort_roles(requested_roles, region)
	
	print('\nWhich client?')
	select_client = get_menu(matrix.clients)
	if (select_client == 0):
		break
	user_client = matrix.clients[select_client - 1]
	
	output_to_screen_and_clipboard(organized_output, matrix)

	#Insert single approver client functionality here.  Append output.
