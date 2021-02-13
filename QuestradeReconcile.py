'''	The QuestradeReconcile macro is python code that uses the Questrade application programming interface (API) to fetch
	account, position, balance, equity, and 30 day activity into a LibreOffice spreadsheet file.

	The spreadsheet file must have six sheets with the following names: Accounts, Positions, Balances, Equities, Activity,
	and Configuration. Except for the Configuration sheet any existing data on these sheets will be cleared when the
    QuestradeReconcile python macro is run.

	Data from the Questrade on-line platform will be used to update these sheets. The sheets can be moved within the file:
	sheet order does not affect anything. Any other sheets in the file are ignored.
	
	The Configuration sheet must have the following cells defined for the QuestradeReconcile macro to function correctly.
		$Configuration.B1  : Questrade authentication token_value text string.
		$Configuration.B3  : A date cell indicating the last time the macro was run.
		$Configuration.B5  : RateLimit API remaining is number of API requests allowed against the current limit.
		$Configuration.B7  : RateLimit API reset is time when the current limit will expire ( Unix timestamp ).
		$Configuration.B10 : A text cell with a comma separated list of equity symbols.
		$Configuration.B12 : A text cell with default logging level.
		$Configuration.B15 : A text cell for logging macro status.
		$Configuration.B17 : Questrade API token cache

	To change these locations the python code must be edited. See the Configuration class.
	
	This script must be copied into the following directory:
		%APPDATA%\\LibreOffice\\4\\user\\Scripts\\python

	QuestradeReconcile is free software: you can redistribute and/or modify it under the terms of the GNU General Public
	License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later
	version.

	QuestradeReconcile is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
	warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU General Public License along with this program. If not, see
	https://www.gnu.org/licenses/.

	Copyright (C) 2019 - 2021 kerouac01850
'''

import logging

class Spreadsheet( ):
	__setvalue = {
		's': lambda o, c, v: o.set_and_format_string( c, v, '@' ),
		'd': lambda o, c, v: o.set_and_format_date( c, v, 'MMM D, YYYY' ),
		'b': lambda o, c, v: o.set_and_format_float( c, v, 'BOOLEAN' ),
		'$': lambda o, c, v: o.set_and_format_float( c, v, '[$$-409]#,##0.00;[RED]([$$-409]#,##0.00)' ),
		'$3': lambda o, c, v: o.set_and_format_float( c, v, '[$$-409]#,##0.000;[RED]([$$-409]#,##0.000)' ),
		'$4': lambda o, c, v: o.set_and_format_float( c, v, '[$$-409]#,##0.0000;[RED]([$$-409]#,##0.0000)' ),
		'n': lambda o, c, v: o.set_and_format_float( c, v, 'General' ),
		'n0': lambda o, c, v: o.set_and_format_float( c, v, '#,##0;[RED]-#,##0' ),
		'n1': lambda o, c, v: o.set_and_format_float( c, v, '#,##0.0;[RED]-#,##0.0' ),
		'n2': lambda o, c, v: o.set_and_format_float( c, v, '#,##0.00;[RED]-#,##0.00' ),
		'n3': lambda o, c, v: o.set_and_format_float( c, v, '#,##0.000;[RED]-#,##0.000' ),
		'n4': lambda o, c, v: o.set_and_format_float( c, v, '#,##0.0000;[RED]-#,##0.0000' ),
		'n5': lambda o, c, v: o.set_and_format_float( c, v, '#,##0.00000;[RED]-#,##0.00000' ),
		'n6': lambda o, c, v: o.set_and_format_float( c, v, '#,##0.000000;[RED]-#,##0.000000' )
	}

	__getvalue = {
		's': lambda o, v: v.getString( ),
		'n': lambda o, v: int( v.getValue( ) ),
		'b': lambda o, v: True if v.getValue( ) else False,
		'd': lambda o, v: o.spreadsheet_to_python_date( v.getValue( ) )
	}

	@classmethod
	def questrade_to_python_date( cls, ds ):
		''' Sample date: '2019-07-02T00:00:00.000000-04:00'
            return: python date
		'''
		from datetime import date
		return None if ds is None else date( int( ds[0:4] ), int( ds[5:7] ), int( ds[8:10] ) )

	@classmethod
	def spreadsheet_to_python_date( cls, serial, mode = 0 ):
		''' mode: 0 for 1900-based, 1 for 1904-based
		'''
		from datetime import timedelta, date
		return date( 1899, 12, 30 ) +  timedelta( days = serial + 1462 * mode )

	@classmethod
	def python_to_spreadsheet_date( cls, dt ):
		from datetime import date
		t = date( 1899, 12, 30 )
		delta = dt - t
		return float( delta.days ) + float( delta.seconds ) / 86400

	@classmethod
	def month_range( cls, startdate = None, range = 1 ):
		from datetime import date
		from calendar import monthrange
		t = date.today( ) if startdate is None else startdate
		year = t.year if t.month + range - 1 <= 12 else t.year + 1
		month = t.month + range - 1 if year == t.year else t.month + range - 1 - 12
		current, last = monthrange( year, month )
		t1 = date( t.year, t.month, 1 )
		t2 = date( year, month, last )
		return ( t1, t2, )

	@classmethod
	def within_monthly_range( cls, dt, months = 1 ):
		t1, t2 = cls.month_range( range = months )
		return dt >= t1 and dt <= t2

	@classmethod
	def format_cell( cls, cell, format ):
		from com.sun.star.uno import RuntimeException
		document = XSCRIPTCONTEXT.getDocument( )
		try:
			cell.NumberFormat = document.NumberFormats.addNew( format, document.CharLocale )
		except RuntimeException:
			cell.NumberFormat = document.NumberFormats.queryKey( format, document.CharLocale, False )

	@classmethod
	def set_and_format_string( cls, cell, value, format ):
		if value is None:
			cell.setString( '' )
			cls.format_cell( cell, format )
		else:
			cell.setString( value )
			cls.format_cell( cell, format )

	@classmethod
	def set_and_format_date( cls, cell, value, format ):
		from datetime import date
		if value is None:
			cls.set_and_format_string( cell, value, format )
		elif isinstance( value, date ):
			cell.setValue( cls.python_to_spreadsheet_date( value ) )
			cls.format_cell( cell, format )
			if cls.within_monthly_range( value ):
				cell.CellBackColor = 0xccffcc	# light green
		else:
			cls.set_and_format_string( cell, value, 'General' )

	@classmethod
	def set_and_format_float( cls, cell, value, format ):
		if value is None:
			cls.set_and_format_string( cell, value, format )
		else:
			cell.setValue( float( value ) )
			cls.format_cell( cell, format )

	@classmethod
	def sheet_by_name( cls, name ):
		desktop = XSCRIPTCONTEXT.getDesktop( )
		model = desktop.getCurrentComponent( )
		return model.Sheets.getByName( name )

	def __init__( self, name, fields ):
		self.name = name
		self.fields = fields
		self.sheet = self.sheet_by_name( self.name )
		self.range = self.sheet.getCellRangeByName( 'A1:AMJ1048576' )
		cursor = self.sheet.createCursor( )
		cursor.gotoEndOfUsedArea( False )
		self.rows = cursor.RangeAddress.EndRow

	def get_row( self, row ):
		data = dict( )
		for field_name, field_type in self.fields:
			value_handler = self.__getvalue.get( field_type, lambda o, v: v.getValue( ) )
			value_cell = self.sheet.getCellByPosition( len( data ), row )
			data[field_name] = value_handler( self, value_cell )
		data[ 'row' ] = row
		return data

	def get_rows( self ):
		for row in range( 0, self.rows ):
			data = self.get_row( row + 1 )
			yield data

	def update_row( self, data ):
		column = 0
		for field_name, field_type in self.fields:
			if self.rows == 0:
				try:
					name_cell = self.sheet.getCellByPosition( column, 0 )
					name_cell.setString( field_name )
					name_cell.CellBackColor = 0xb4c7dc	# light blue / grey
				except:
					logging.exception( 'Failed to set column name cell[{},0] = {}'.format( column, field_name ) )
			row = self.rows + 1 if 'row' not in data else data[ 'row' ]
			if field_name in data:
				try:
					default_handler = lambda o, c, v: o.set_and_format_string( c, v, '@' )
					value_handler = self.__setvalue.get( field_type, default_handler )
					value_cell = self.sheet.getCellByPosition( column, row )
					value_handler( self, value_cell, data[ field_name ] )
				except:
					ctrl = '{}::update_row( n={} t={} r={} c={} v={} )'
					logging.exception( ctrl.format( self.name, field_name, field_type, row, column, data[ field_name] ) )
			column = column + 1
		if row > self.rows:
			self.rows = row
		return row

	def clear_contents( self ):
		self.range.clearContents( 1 | 2 | 4 | 8 | 16 | 32 | 64 | 128 | 256 | 512 )
		self.rows = 0

	def search_descriptor( self, range, target ):
		descriptor = range.createSearchDescriptor( )
		descriptor.setSearchString( target )
		descriptor.SearchCaseSensitive = True
		descriptor.SearchWords = True
		return descriptor

	def find_cell( self, sRange, sTarget ):
		oRange = self.sheet.getCellRangeByName( sRange )
		oDescriptor = self.search_descriptor( oRange, sTarget )
		return oRange.findFirst( oDescriptor )

	def sortfield_by_index( self, index, ascending = True ):
		from com.sun.star.util import SortField
		sf = SortField( )
		sf.Field = index
		sf.SortAscending = ascending
		return sf

	def property_value( self, name, value ):
		from com.sun.star.beans import PropertyValue
		pv = PropertyValue( )
		pv.Name = name
		pv.Value = value
		return pv

	def sort_by_indicies( self, indicies, ascending = True ):
		from uno import Any
		sort_fields = tuple( self.sortfield_by_index( index, ascending ) for index in indicies )
		property_fields = self.property_value( 'SortFields', Any( '[]com.sun.star.util.SortField', sort_fields ) )
		property_header = self.property_value( 'HasHeader', True )
		self.range.sort( [ property_fields, property_header ] )

class Configuration( Spreadsheet ):
	__token         = '$Configuration.B1'
	__timestamp     = '$Configuration.B3'
	__api_remaining = '$Configuration.B5'
	__api_reset     = '$Configuration.B7'
	__equitylist    = '$Configuration.B10'
	__loglevel      = '$Configuration.B12'
	__log           = '$Configuration.B15'
	__apicache      = '$Configuration.B17'

	@classmethod
	def cellbyref( cls, reference ):
		values = reference.split( '.' )
		sheet = cls.sheet_by_name( values[0].strip( '$' ) )
		return sheet.getCellRangeByName( values[1].strip( '$' ) )

	@classmethod
	def getbyref( cls, reference, typeof = str ):
		cell = cls.cellbyref( reference )
		return cell.getString( ) if typeof == str else cell.getValue( )

	@classmethod
	def setbyref( cls, reference, value ):
		cell = cls.cellbyref( reference )
		if value is None:
			cell.setString( '' )
		elif type( value ) == str:
			cell.setString( value )
		else:
			cell.setValue( float( value ) )
		return value

	@classmethod
	def get_token( cls ):
		return cls.getbyref( cls.__token )

	@classmethod
	def set_timestamp( cls, value ):
		cls.setbyref( cls.__timestamp, str( value ) )

	@classmethod
	def set_remaining( cls, value ):
		cls.setbyref( cls.__api_remaining, float( value ) )

	@classmethod
	def set_reset( cls, value ):
		cls.setbyref( cls.__api_reset, float( value ) )

	@classmethod
	def get_equitylist( cls ):
		return cls.getbyref( cls.__equitylist )

	@classmethod
	def get_loglevel( cls ):
		return cls.getbyref( cls.__loglevel )

	@classmethod
	def get_log( cls ):
		return cls.getbyref( cls.__log )

	@classmethod
	def set_log( cls, value ):
		cls.setbyref( cls.__log, str( value ) )

	@classmethod
	def get_apicache( cls ):
		return cls.getbyref( cls.__apicache )

	@classmethod
	def set_apicache( cls, value ):
		cls.setbyref( cls.__apicache, str( value ) )

	def __init__( self ):
		configuration_name = "Configuration"
		configuration_fields = None
		super( ).__init__( configuration_name, configuration_fields )

class Accounts( Spreadsheet ):
	def __init__( self ):
		accounts_name = 'Accounts'
		accounts_fields = [
			( 'number', 's', ),
			( 'type', 's', ),
			( 'clientAccountType', 's', ),
			( 'status', 's', ),
			( 'isPrimary', 'b', ),
			( 'isBilling', 'b', )
		]
		super( ).__init__( accounts_name, accounts_fields )

	def fetch( self, connection ):
		try:
			quest_accounts = connection.questrade.accounts
			if 'accounts' not in quest_accounts:
				raise RuntimeError( 'Accounts::accounts = {}'.format( quest_accounts ) )
		except:
			logging.exception( 'Accounts::fetch( ) failed!' )
			return
		for account in quest_accounts['accounts']:
			logging.debug( 'Accounts::fetch( ) = {}'.format( account ) )
			yield account

	def reconcile( self, connection ):
		self.clear_contents( )
		for account in self.fetch( connection ):
			self.update_row( account )
		self.default_sort( )

	def default_sort( self ):
		# account (number A = 0)
		self.sort_by_indicies( ( 0, ) )

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
		equities = Equities( )
		for account in Accounts( ).get_rows( ):
			for position in self.fetch( connection, account ):
				equity = equities.search_by_symbol( position['symbol'] )
				if equity:
					position['currency'] = equity['currency']
				self.update_row( position )
		self.default_sort( )

	def default_sort( self ):
		# account (column A = 0), currency (column C = 2), symbol (column D = 3)
		self.sort_by_indicies( ( 0, 2, 3, ) )

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

	def search_by_symbol( self, sTarget, sRange = 'D1:D1048576' ):
		cell = self.find_cell( sRange, sTarget )
		if cell is None:
			return None
		address = cell.getCellAddress( )
		row = address.Row
		return self.get_row( row )

	def column_from( self, oCell ):
		return oCell.getCellAddress( ).Column

	def column_by_name( self, name ):
		return self.column_from( self.find_cell( "A1:AMJ1", name ) )

	def alias_for( self, name, row ):
		return chr( ord( 'A' ) + self.column_by_name( name ) ) + str( row + 1 )

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
				raise RuntimeError( 'Activities::symbols( {} ) = {}'.format( quest_equities ) )
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

class Connection( ):
	def __init__( self ):
		self.__questrade = self.__questrade_cache_connect( )
		if self.__questrade is None:
			self.__questrade = self.__questrade_token_connect( )
		if self.__questrade is None:
			ctrl = 'MBOX: Failed to authenticate. Generate new Questrade token for {}'
			logging.critical( ctrl.format( Configuration.get_token( ) ) )
			raise RuntimeError( 'Unable to authenticate connection with Questrade.' )

	def __questrade_cache_connect( self ):
		from questrade_api import Questrade
		if len( Configuration.get_apicache( ) ) == 0:
			return None
		try:
			logging.info( 'MBOX: Authenticate with credentials from {}'.format( Configuration.get_apicache( ) ) )
			q = Questrade(
				logger = logging.debug,
				storage_adaptor = ( lambda : self.__read_token_cache( ), lambda s: self.__write_token_cache( s ) ) )
			Configuration.set_timestamp( q.time['time'] )
		except:
			logging.exception( 'Failed to authenticate using application key {}'.format( Configuration.get_apicache( ) ) )
			q = None
		return q

	def __questrade_token_connect( self ):
		from questrade_api import Questrade
		if len( Configuration.get_token( ) ) == 0:
			return None
		try:
			logging.info( 'MBOX: Authenticate with credentials from {}'.format( Configuration.get_token( ) ) )
			q = Questrade(
				logger = logging.debug,
				storage_adaptor = ( lambda : self.__read_token_cache( ), lambda s: self.__write_token_cache( s ), ),
				refresh_token = Configuration.get_token( ) )
			logging.info( 'MBOX: Authenticate with credentials from {}'.format( Configuration.get_apicache( ) ) )
			Configuration.set_timestamp( q.time['time'] )
		except:
			logging.exception( 'Failed to authenticate using Questrade token {}'.format( Configuration.get_token( ) ) )
			q = None
		return q

	def __read_token_cache( self ):
		return Configuration.get_apicache( )

	def __write_token_cache( self, s ):
		Configuration.set_apicache( s )

	@property
	def questrade( self ):
		return self.__questrade

	def finalize_ratelimits( self ):
		remaining, reset = self.questrade.ratelimit
		Configuration.set_remaining( remaining )
		Configuration.set_reset( reset )

	def update_from_server( self ):
		logging.info( 'MBOX: Reconcile from Questrade started!' )
		Accounts( ).reconcile( self )
		Balances( ).reconcile( self )
		Activities( ).reconcile( self )
		Positions( ).reconcile( self )
		Equities( ).reconcile( self )
		self.finalize_ratelimits( )
		logging.info( 'MBOX: Reconcile complete.' )

def RunMacro( title, macro ):
	from com.sun.star.awt.MessageBoxType import MESSAGEBOX
	from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
	from uno import getComponentContext
	from io import StringIO

	logger = logging.getLogger( )
	formatter = logging.Formatter( '%(asctime)s : %(levelname)s : %(message)s' )
	logger.setLevel( Configuration.get_loglevel( ) )

	logging_buffer = StringIO( )
	streamhandler = logging.StreamHandler( logging_buffer )
	streamhandler.setFormatter( formatter )
	streamhandler.setLevel( Configuration.get_loglevel( ) )
	logger.addHandler( streamhandler )

	message_buffer = StringIO( )
	streamhandler = logging.StreamHandler( message_buffer )
	streamhandler.setFormatter( formatter )
	streamhandler.setLevel( Configuration.get_loglevel( ) )
	streamhandler.addFilter( lambda r: r.getMessage( ).startswith( 'MBOX:' ) )
	logger.addHandler( streamhandler )

	desktop = XSCRIPTCONTEXT.getDesktop( )
	model = desktop.getCurrentComponent( )
	try:
		model.lockControllers( )
		model.addActionLock( )
		macro( )
	except:
		logging.exception( 'MBOX: {} failed because of an exception!'.format( title ) )

	finally:
		model.removeActionLock( )
		model.unlockControllers( )

	manager = getComponentContext( ).ServiceManager
	dispatcher = manager.createInstance( "com.sun.star.frame.DispatchHelper" )
	document = model.CurrentController.Frame
	dispatcher.executeDispatch( document, ".uno:CalculateHard", "", 0, ( [ ] ) )

	parentwin = model.CurrentController.Frame.ContainerWindow
	box = parentwin.getToolkit( ).createMessageBox( parentwin, MESSAGEBOX, BUTTONS_OK, title, message_buffer.getvalue( ) )
	result = box.execute( )

def QuestradeReconcile( *args ):
	RunMacro( 'QuestradeReconcile', lambda : Connection( ).update_from_server( ) )
	return None

def ReconcileAccounts( *args ):
	RunMacro( 'ReconcileAccounts', lambda : Accounts( ).reconcile( Connection( ) ) )
	return None

def ReconcileBalances( *args ):
	RunMacro( 'ReconcileBalances', lambda : Balances( ).reconcile( Connection( ) ) )
	return None

def ReconcileActivities( *args ):
	RunMacro( 'ReconcileActivities', lambda : Activities( ).reconcile( Connection( ) ) )
	return None

def ReconcileEquities( *args ):
	RunMacro( 'ReconcileEquities', lambda : Equities( ).reconcile( Connection( ) ) )
	return None

def ReconcilePositions( *args ):
	RunMacro( 'ReconcilePositions', lambda : Positions( ).reconcile( Connection( ) ) )
	return None

def TestCache( *args ):
	RunMacro( 'TestCache', lambda: Configuration.CACHE( ) )
	return None

g_exportedScripts = QuestradeReconcile,ReconcileAccounts,ReconcileBalances,ReconcileActivities,ReconcileEquities,ReconcilePositions,TestCache
