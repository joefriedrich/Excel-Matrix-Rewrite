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
	def __init__(self, name, file_name, regex, clients):
		self.name = name
		self.excel_file = openpyxl.load_workbook(file_name, read_only = True)
		self.regex = regex
		self.clients = clients
		self.vacation_lookup = self.load_vacations()
		self.regions = self.load_regions()
		self.role_lookup = self.load_roles()
		self.email_lookup = self.load_emails()
		self.excel_file = None

	def load_regions(self):
		'''
			This grabs the contents of the cells in the top row of the Roles worksheet.
			It stores them in a list and stops when it gets to a blank cell.
			It returns all entries after the first 4.
			These are the headings that correspond with the companies regions.
		'''
		header_text = []
		for role in self.excel_file["Roles"]:
			for cell in role:
				if cell.value == None:
					break
				header_text.append(cell.value)
			break
		return header_text[4:]


	def vacation_check(self, names):
		'''
		'''	
		checked_names = []
		for name in names:
			for vacation in self.vacation_lookup:
				if name == None:
					break
				elif name == vacation.standard:
					if vacation.start <= datetime.now() <= vacation.finish:
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
		if row[2].value == 'Regional':
			approvers = [approver.value for approver in row[4:4 + len(self.regions)]]
		else:
			for approver in row[4:4 + len(self.regions)]:
				if approver.value != 'N/A':
					approvers.append(approver.value)
					break
		approvers = self.vacation_check(approvers)
		return approvers

	
	def load_roles(self):
		'''
			This creates a list of tuples that contain
				the role name.
				the role definition.
				and a list of the role approvers (created by load_approvers).
			Region was not included, because we can get that from
				the length of the list of role approvers (when greater than 1).
		'''
		print('Loading lookup dictionary...')
		role = namedtuple('role',
						  'name description approvers')
		role_list = []
		for row in self.excel_file['Roles'].rows:
			if row[0].value in ('', None):
				print('Issue with a row.')   #Can we pull the excel row number from this?
				break
			role_tuple = role(row[0].value, row[1].value, self.load_approvers(row))
			role_list.append(role_tuple)
		return role_list


	def load_emails(self):
		'''
			Creates a dictionary of emails from the table of names/emails in the sprreadsheet.
		'''
		emails = {}
		for row in self.excel_file['Names and Email'].rows:
			emails[row[0].value] = row[1].value
		return emails


	def load_vacations(self):
		'''
			Creates a list of tuples from the information in the OnHoliday spreadsheet.
		'''
		vacation = namedtuple('vacation',
				'standard replacement start finish')
		vacation_list = []
		for row in self.excel_file['OnHoliday'].rows:
			vacation_tuple = vacation(row[0].value,
						row[1].value,
						row[2].value,
						row[3].value)
			vacation_list.append(vacation_tuple)
		return vacation_list


	def check_and_sort_roles(self, user_input, user_region):
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

#-------------------------------------------------------------------------------
def get_menu(list_of_things):
	'''
		This takes a list as input.
		It prints the list to the screen as menu options 1 through len(list).
		It calls and returns the results of get_menu_input.
	'''
	menu_number = 0
	for thing in list_of_things:
		menu_number += 1
		print(str(menu_number) +  ':  ' + thing)
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
	while True: #left while true in incase we need to make it harder to quit
		try:
			user_input = int(input('Any other entry to quit:  '))
			if 1 <= user_input <= menu_total:
				return user_input
			else:
				return 0
		except ValueError:  #if the value is a non-numeral or float, return 0
			return 0

	
def get_role_input():
	'''
		This is just a placeholder for user input that is not in the form
			of get_menu_input.
	'''
	return stdin.readlines()


def output_to_screen_and_clipboard(output, company):
	'''
	'''
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

	
#-------------------------------------------------------------------------------
simplefilter('ignore')  #blocks some warnings that openpyxl throws when loading files

print('Loading companies.')
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

print('Loading single approvers.')		
single_approver_clients = []
single_approver_clients.append(('Company One Special Test Client', 'CP1TST -- awaiting approval from singleApproverCP1/nSpecialRole', 'singleApprover@company1.com'))
single_approver_clients.append(('Company Two Special Test Client', 'CP2TST -- awaiting approval from singleApproverCP2/nSpecialRole', 'singleApprover@company2.com'))
single_approver_clients.append(('Company Three Special Test Client', 'CP3TST -- awaiting approval from singleApproverCP3/nSpecialRole', 'singleApprover@company3.com'))
single_approver_clients.append(('Company One Test Email Duplication', 'CP1TST -- awaiting approval from ApproverOneCP1/nSpecialRole', 'Approver1@company1.com'))
single_approver_clients.append(('Company Two Test Email Duplication', 'CP2TST -- awaiting approval from ApproverOneCP2/nSpecialRole', 'Approver1@company2.com'))
single_approver_clients.append(('Company Three Test Email Duplication', 'CP3TST -- awaiting approval from ApproverOneCP3/nSpecialRole', 'Approver1@company3.com'))

list_company_names = [company.name for company in list_companies]
list_single_approvers = [client[0] for client in single_approver_clients]
			
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
	
	output_to_screen_and_clipboard(organized_output, matrix) #need to pass it clipboard. this will change to return emails and clipboard.
	
	#Insert single approver client functionality here.  Print output.  Append to clipboard.
