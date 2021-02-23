'''	QuestradeReconcile is free software: you can redistribute and/or modify it under the terms of the GNU General Public
	License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later
	version.

	QuestradeReconcile is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
	warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU General Public License along with this program. If not, see
	https://www.gnu.org/licenses/.

	Copyright (C) 2019 - 2021 kerouac01850
'''

import logging
from spreadsheet import Spreadsheet, Accounts

class Balances( Spreadsheet ):
	CAD = 0
	USD = 1

	def __init__( self ):
		balances_name = 'Balances'
		balances_fields = [
			( 'balanceType', 's', ),
			( 'account', 's', ),
			( 'accountType', 's', ),
			( 'currency', 's', ),
			( 'cash', '$', ),
			( 'marketValue', '$', ),
			( 'totalEquity', '$', ),
			( 'buyingPower', '$', ),
			( 'maintenanceExcess', '$', ),
			( 'isRealTime', 'b', )
		]
		super( ).__init__( balances_name, balances_fields )

	def fetch( self, connection, account ):
		try:
			quest_balances = connection.questrade.account_balances( account['number'] )
			if 'combinedBalances' not in quest_balances or 'perCurrencyBalances' not in quest_balances:
				ctrl = 'ASSERT combinedBalances or perCurrencyBalances expected but not found! Balances::fetch( {} ) = {}.'
				raise RuntimeError( ctrl.format( account['number'], quest_balances ) )
		except:
			logging.exception( 'Balances::fetch( {} )'.format( account['number'] ) )
			return
		for balanceType in [ 'combinedBalances', 'perCurrencyBalances' ]:
			for currency in [ Balances.CAD, Balances.USD ]:
				quest_balances[balanceType][currency]['balanceType'] = balanceType
				quest_balances[balanceType][currency]['account'] = account['number']
				quest_balances[balanceType][currency]['accountType'] = account['type']
				logging.debug( 'Balances::fetch( {} ) = {}'.format( account['number'], quest_balances[balanceType][currency] ) )
				yield quest_balances[balanceType][currency]

	def reconcile( self, connection ):
		self.clear_contents( )
		for account in Accounts( ).get_rows( ):
			for balance in self.fetch( connection, account ):
				self.update_row( balance )
		self.default_sort( )

	def default_sort( self ):
		# balanceType (Column A = 0), account (Column B = 1), currency (Column D = 3)
		self.sort_by_indicies( ( 0, 1, 3, ) )
