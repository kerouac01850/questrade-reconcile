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

class Activities( Spreadsheet ):
	from datetime import date, timedelta

	startDate = ( date.today( ) - timedelta( days = 29 ) ).isoformat( ) + 'T00:00:00-04:00'
	endDate = date.today( ).isoformat( ) + 'T00:00:00-04:00'

	def __init__( self ):
		activities_name = 'Activities'
		activities_fields = [
			( 'account', 's', ),
			( 'accountType', 's', ),
			( 'currency', 's', ),
			( 'transactionDate', 'd', ),
			( 'symbol', 's', ),
			( 'symbolId', 'n', ),
			( 'type', 's', ),
			( 'action', 's', ),
			( 'quantity', 'n3', ),
			( 'price', '$4', ),
			( 'grossAmount', '$', ),
			( 'commission', '$', ),
			( 'netAmount', '$', ),
			( 'tradeDate', 'd', ),
			( 'settlementDate', 'd', ),
			( 'description', 's', )
		]
		super( ).__init__( activities_name, activities_fields )

	def fetch( self, connection, account ):
		try:
			quest_activities = connection.questrade.account_activities( account['number'],
				startTime = self.startDate,
				endTime = self.endDate )
			if 'activities' not in quest_activities:
				ctrl = 'Activities::account_activities( {}, {}, {} ) = {}'
				raise RuntimeError( ctrl.format( account['number'], self.startDate, self.endDate, quest_activities ) )
		except:
			logging.exception( 'Activities::fetch( {} ) failed!'.format( account ) )
			return
		for activity in quest_activities['activities']:
			activity['account'] = account['number']
			activity['accountType'] = account['type']
			activity['transactionDate'] = self.questrade_to_python_date( activity['transactionDate'] )
			activity['tradeDate'] = self.questrade_to_python_date( activity['tradeDate'] )
			activity['settlementDate'] = self.questrade_to_python_date( activity['settlementDate'] )
			logging.debug( 'Activities::fetch( {} ) = {}'.format( account['number'], activity ) )
			yield activity

	def reconcile( self, connection ):
		self.clear_contents( )
		for account in Accounts( ).get_rows( ):
			for activity in self.fetch( connection, account ):
				self.update_row( activity )
		self.default_sort( )

	def default_sort( self ):
		# account (Column A = 0), currency (Column C = 2), transactionDate (Column D = 3), Symbol (Column E = 4)
		self.sort_by_indicies( ( 0, 2, 3, 4, ) )
