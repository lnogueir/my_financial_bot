import plaid
import datetime
from my_config import *


def format_error(e):
  return {'error': {'display_message': e.display_message, 'error_code': e.code, 'error_type': e.type, 'error_message': e.message } }

class Accounter(plaid.Client):
	def __init__(self):
		plaid.Client.__init__(self,client_id = client_id, secret=secret_development_key,
		          public_key=public_key, environment=environment, api_version='2019-05-29')
		self.access_token=access_token
		self.start_day = start_date # this is the day you will start counting each month

	def get_net_transaction_balance(self,transactions):
		net_transaction_balance = 0
		for transaction in transactions:
			net_transaction_balance-=transaction['amount']		
		return net_transaction_balance	

	def get_balance(self):
		try:
			balance_response = self.Accounts.balance.get(self.access_token)
		except plaid.errors.PlaidError as e:
			print(format_error(e))
			return None
		return balance_response['accounts'][1]['balances']['available']

	def get_new_transactions(self, old_dates):
		new_transactions = []
		all_transactions = self.get_transactions()['transactions']
		for transaction in all_transactions:
			if transaction['date'] not in old_dates: 
				new_transactions.append(transaction)
		return new_transactions

	def get_transactions(self):
	  # Pull transactions from the last 30 days
		start_date = '{:%Y-%m-%d}'.format(datetime.datetime.now() + datetime.timedelta(-30))
		end_date = '{:%Y-%m-%d}'.format(datetime.datetime.now())
		try:
			transactions_response = self.Transactions.get(self.access_token, start_date, end_date)
		except plaid.errors.PlaidError as e:
			print(format_error(e))
		return transactions_response

