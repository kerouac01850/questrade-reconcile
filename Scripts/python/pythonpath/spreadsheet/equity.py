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
from spreadsheet import Spreadsheet, Positions, Configuration

class Equities( Spreadsheet ):
	def __init__( self ):
		equities_name = 'Equities'
		equities_fields = [
			( 'account', 's', ),
			( 'accountType', 's', ),
			( 'currency', 's', ),
			( 'symbol', 's', ),
			( 'symbolId', 'n', ),
			( 'description', 's', ),
			( 'listingExchange', 's', ),
			( 'securityType', 's', ),
			( 'prevDayClosePrice', '$', ),
			( 'yield', 'n4', ),
			( 'pe', 'n4', ),
			( 'eps', 'n4', ),
			( 'outstandingShares', 'n0', ),
			( 'marketCap', 'n0', ),
			( 'averageVol20Days', 'n0', ),
			( 'averageVol3Months', 'n0', ),
			( 'dividend', 'n4', ),
			( 'dividendDate', 'd', ),
			( 'exDate', 'd', ),
			( 'lowPrice52', '$', ),
			( 'highPrice52', '$', ),
			( 'tradeUnit', 'b', ),
			( 'pay', '$', )
		]
		super( ).__init__( equities_name, equities_fields )

	def set_pay( self, row_id ):
		date_id = self.column_by_name( "dividendDate" )
		date_cell = self.sheet.getCellByPosition( date_id, row_id )
		date_value = self.spreadsheet_to_python_date( date_cell.getValue( ) )
		if self.within_monthly_range( date_value ):
			account_alias = self.alias_for( "account", row_id )
			symbol_alias = self.alias_for( "symbol", row_id )
			dividend_alias = self.alias_for( "dividend", row_id )
			pay_id = self.column_by_name( "pay" )
			pay_cell = self.sheet.getCellByPosition( pay_id, row_id )
			pay_cell.setFormula( '={}*UNITQTY({};{})'.format( dividend_alias, account_alias, symbol_alias ) )

	def fetch( self, connection, symbol_names ):
		try:
			if len( symbol_names ) == 0:
				logging.debug( 'Equities::fetch( len( symbol_names ) is 0 )' )
				return
			quest_equities = connection.questrade.symbols( names = symbol_names )
			if 'symbols' not in quest_equities:
				raise RuntimeError( 'Equities::symbols( {} ) = {}'.format( quest_equities ) )
		except:
			logging.exception( 'Equities::fetch( {} ) failed!'.format( symbol_names ) )
			return
		for equity in quest_equities['symbols']:
			equity['account'] = None
			equity['accountType'] = None
			equity['dividendDate'] = self.questrade_to_python_date( equity['dividendDate'] )
			equity['exDate'] = self.questrade_to_python_date( equity['exDate'] )
			logging.debug( 'Equities::fetch( {} ) = {}'.format( equity['account'], symbol_names ) )
			yield equity

	def fetch_unique( self, connection, position ):
		try:
			quest_equities = connection.questrade.symbol( position['symbolId'] )
			if 'symbols' not in quest_equities:
				raise RuntimeError( 'Activities::symbol( {} ) = {}'.format( position['symbolId'], quest_equities ) )
		except:
			logging.exception( 'Equities::fetch_unique( {} ) failed'.format( position ) )
			return None
		equity = quest_equities['symbols'][0]
		equity['account'] = position['account']
		equity['accountType'] = position['accountType']
		equity['dividendDate'] = self.questrade_to_python_date( equity['dividendDate'] )
		equity['exDate'] = self.questrade_to_python_date( equity['exDate'] )
		equity['pay'] = 0.0
		logging.debug( 'Equities::fetch( {} ) = {}'.format( equity['account'], position ) )
		return equity

	def reconcile( self, connection ):
		self.clear_contents( )
		for position in Positions( ).get_rows( ):
			equity = self.fetch_unique( connection, position )
			if equity:
				row_id = self.update_row( equity )
				self.set_pay( row_id )
		for equity in self.fetch( connection, Configuration.get_equitylist( ) ):
			if equity:
				self.update_row( equity )
		self.default_sort( )

	def default_sort( self ):
		# account (Column A = 0), currency (Column C = 2), symbol (Column C = 3)
		self.sort_by_indicies( ( 0, 2, 3, ) )
