from accounter import Accounter
import datetime
import xlrd
import xlsxwriter

# NOTE: You must have the template excel file with Date, Balance, Cash flow as your columns.

class Exceler:
	def __init__(self, workbook='template.xlsx'):
		self.excel_workbook = xlrd.open_workbook(workbook) 
		self.accounter = Accounter()
		self.workbook_name = workbook
		self.python_workbook = None


	def read_workbook_to_python(self):
		self.python_workbook = []
		for sheet_ite, sheet in enumerate(self.excel_workbook.sheets()):
			self.python_workbook.append({sheet.name:{}})
			for i in range(sheet.nrows):
				for j in range(sheet.ncols):
					if i == 0:
						self.python_workbook[sheet_ite][sheet.name][sheet.cell(0, j).value] = []
					else:
						self.python_workbook[sheet_ite][sheet.name][sheet.cell(0,j).value].append(sheet.cell(i,j).value)
		return self.python_workbook				

	def write_workbook_to_excel(self):
		if self.python_workbook:
			new_workbook = xlsxwriter.Workbook(self.workbook_name)
			worksheet = None 
			# Rewriting tracked transactions
			for sheet_obj in self.python_workbook:
				sheet_name = next(iter(sheet_obj))
				worksheet = new_workbook.add_worksheet(sheet_name)
				for i,col in enumerate(sheet_obj[sheet_name]):
					worksheet.write(0, i, col, new_workbook.add_format({'bold': True, 'italic': True}))
					for j,row in enumerate(sheet_obj[sheet_name][col]):
						j+=1
						worksheet.write(j,i,row)
			# Updating with untracked transactions
			last_sheet_obj = self.python_workbook[-1]
			last_sheet_name = next(iter(last_sheet_obj))	
			current_number_of_rows_excel = len(last_sheet_obj[last_sheet_name]['Date']) + 1		
			row_n = current_number_of_rows_excel if current_number_of_rows_excel > 1 else 0
			stored_days = last_sheet_obj[last_sheet_name]['Date']
			new_transactions = self.accounter.get_new_transactions(stored_days)
			last_recorded_month = last_sheet_name.split('_')
			last_month = int(last_recorded_month[0])
			last_year = int(last_recorded_month[1])
			last_recorded_balance = last_sheet_obj[last_sheet_name]['Balance'][-1] if last_sheet_obj[last_sheet_name]['Balance']!=[] else self.accounter.get_balance() - self.accounter.get_net_transaction_balance(new_transactions)
			have_added_spreadsheet = False
			same_day_transactions = []
			for transaction in new_transactions[::-1]:
				row_n+=1
				transaction_date = datetime.datetime.strptime(transaction['date'], '%Y-%m-%d')
				if (not have_added_spreadsheet) and (transaction_date.day >= self.accounter.start_day) and (last_year < transaction_date.year or (last_year == transaction_date.year and last_month < transaction_date.month)):
					worksheet.write(row_n, 3, "Total spent this month:")
					worksheet.write(row_n, 4, '=SUM(C2:C'+str(row_n)+')')
					worksheet = new_workbook.add_worksheet(str(transaction_date.month)+'_'+str(transaction_date.year))
					row_n = 1
					have_added_spreadsheet = True
					for ite, col in enumerate(('Date', 'Balance', 'Cash flow')):
						worksheet.write(0, ite, col, new_workbook.add_format({'bold': True, 'italic': True}))	
				if (same_day_transactions==[]) or same_day_transactions[0]['date']==transaction['date']:
					row_n-=1
					same_day_transactions.append(transaction)
				else:
					net_daily_expense = 0
					for daily_transactions in same_day_transactions:
						net_daily_expense-=daily_transactions['amount']
					date = same_day_transactions[0]['date']
					last_recorded_balance+=net_daily_expense
					for ite, cell_data in enumerate((date, last_recorded_balance,net_daily_expense)):
						worksheet.write(row_n,ite, cell_data)
					same_day_transactions=[transaction]
			if same_day_transactions!=[]:
				row_n+=1
				net_daily_expense = 0
				for daily_transactions in same_day_transactions:
					net_daily_expense-=daily_transactions['amount']
				date = same_day_transactions[0]['date']
				last_recorded_balance+=net_daily_expense
				for ite, cell_data in enumerate((date, last_recorded_balance,net_daily_expense)):
					worksheet.write(row_n,ite, cell_data)
			new_workbook.close()
			return True
		return None	


