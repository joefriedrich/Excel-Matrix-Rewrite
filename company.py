'''
	Title:  Excel Matrix tool (Company Class)
	Author:  Joe Friedrich
	License:  MIT
'''
import csv
import re
from collections import namedtuple
from datetime import date

class Company:
	def __init__(self, name, folder_path, regular_expression, clients):
		self.name = name
		self.folder_path = folder_path
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
			This grabs the contents of the cells in the top row of the Roles
				worksheet.
			It stores them in a list and stops when it gets to a blank cell.
			It returns all entries after the first 4.
			These are the headings that correspond with the companies regions.
		'''
		role_file = open(self.folder_path + '/Roles.csv', encoding='ansi')
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
			approvers = row[4:4 + len(self.region_lookup) + 1] #the +1 was added when we removed the 'global' option from the approval list.
		else:
			for approver in row[4:4 + len(self.region_lookup) + 1]:
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
			Region was not included, because we can get that from
				the length of the list of role approvers (when greater than 1)
		'''
		print('Loading lookup dictionary...')
		
		role_file = open(self.folder_path + '/Roles.csv', encoding='ansi')
		role_reader = csv.reader(role_file)
		role_data = list(role_reader)
		role_file.close()
		
		role = namedtuple('role',
						'name description approvers')
		role_list = []
		for row in role_data:
			if row[0] in ('', None):
				print('Issue with a row.  ' + row)   #Can we pull row data?
				break
			role_data = role(row[0], row[1],
						self.load_approvers(row))
			role_list.append(role_data)
		
		return role_list


	def load_emails(self):
		''' ***REWRITE***
			Creates a dictionary from the information in the 'Names and Email'
				spreadsheet.
		'''
		email_file = open(self.folder_path + '/Email.csv', encoding='ansi')
		email_reader = csv.reader(email_file)
		email_data = list(email_reader)
		email_file.close()
		
		emails = {}
		for row in email_data:
			emails[row[0]] = row[1]
		
		return emails


	def load_vacations(self):
		''' ***REWRITE***
			Creates a list of tuples from the information in the OnHoliday
				spreadsheet.
		'''		
		vacation_file = open(self.folder_path + '/Vacations.csv', encoding='ansi')
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
			#Needs written!!
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