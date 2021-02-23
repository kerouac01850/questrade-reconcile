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

class Positions( Spreadsheet ):
	def __init__( self ):
		positions_name = 'Positions'
		positions_fields = [
			( 'account', 's', ),
			( 'accountType', 's', ),
			( 'currency', 's', ),
			( 'symbol', 's', ),
			( 'symbolId', 'n', ),
			( 'openQuantity', 'n3', ),
			( 'currentPrice', '$', ),
			( 'currentMarketValue', '$', ),
			( 'averageEntryPrice', '$3', ),
			( 'totalCost', '$', ),
			( 'openPnl', '$', ),
			( 'dayPnl', '$', ),
			( 'closedQuantity', 'n3', ),
			( 'closedPnl', '$', ),
			( 'isUnderReorg', 'b', ),
			( 'isRealTime', 'b', )
		]
		super( ).__init__( positions_name, positions_fields )

	def set_currency( self, row_id ):
		currency_id = self.column_by_name( "currency" )
		currency_cell = self.sheet.getCellByPosition( currency_id, row_id )
		symbol_alias = self.alias_for( "symbol", row_id )
		currency_cell.setFormula( '=CURRENCYOF({})'.format( symbol_alias ) )

	def fetch( self, connection, account ):
		try:
			quest_positions = connection.questrade.account_positions( account['number'] )
			if 'positions' not in quest_positions:
				raise RuntimeError( 'Activities::account_positions( {} ) = {}'.format( account['number'], quest_positions ) )
		except:
			logging.exception( 'Positions::fetch( {} ) failed'.format( account ) )
			return
		for position in quest_positions['positions']:
			position['account'] = account['number']
			position['accountType'] = account['type']
			logging.debug( 'Positions::fetch( {} ) = {}'.format( account['number'], position ) )
			yield position

	def reconcile( self, connection ):
		self.clear_contents( )
		for account in Accounts( ).get_rows( ):
			for position in self.fetch( connection, account ):
				row_id = self.update_row( position )
				self.set_currency( row_id )
		self.default_sort( )

	def default_sort( self ):
		# account (column A = 0), currency (column C = 2), symbol (column D = 3)
		self.sort_by_indicies( ( 0, 2, 3, ) )
