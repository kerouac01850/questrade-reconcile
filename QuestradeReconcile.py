'''	The QuestradeReconcile macro is python code that uses the Questrade application programming interface (API) to fetch
	account, position, balance, equity, 30 day activity, and dividend data into a LibreOffice spreadsheet file.

	The spreadsheet file must have seven sheets with the following names: Accounts, Positions, Balances, Equities, Activity,
	Dividends and Configuration. Except for the Configuration sheet any existing data on these sheets will be cleared when
	the QuestradeReconcile python macro is run.

	 Data from the Questrade on-line platform will be used to update these sheets. The sheets can be moved within the file:
	 sheet order does not affect anything. Any other sheets in the file are ignored.
	
	The Configuration sheet must have the following cells defined for the QuestradeReconcile macro to function correctly.
		Cell $Configuration.B1: Questrade authentication token_value text string.
		Cell $Configuration.B3: A text cell with a comma separated list of equity symbols.
		Cell $Configuration.B5: A date cell indicating the last time the macro was run.
		Cell $Configuration.B7: A text cell for logging macro status.

	To change these locations the python code must be edited. See config_token, config_equities, config_timestamp, and
	config_log, at the top of the file.
	
	This script must be copied into the following directory:
		C:\Program Files\LibreOffice\share\Scripts\python\QuestradeData.py

	TUTORIAL:
		Interface-oriented programming in OpenOffice / LibreOffice : automate your office tasks with Python Macros.
		http://christopher5106.github.io/office/2015/12/06/openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html
	
	TODO:

	Changes
		August 2019	: Format and sort values in Accounts, Positions, Balances, Equities, and Activities.
		July 2019	: Added Activity sheet ... goes back 30 days from today.
		June 2019	: Add numbers as numeric values and dates as date values instead of as text values.

	QuestradeReconcile is free software: you can redistribute and/or modify it under the terms of the GNU General Public
	License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later
	version.

	QuestradeReconcile is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
	warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU General Public License along with this program. If not, see
	https://www.gnu.org/licenses/.

	Copyright (C) 2019 kerouac01850
'''

config_token = '$Configuration.B1'
config_equities = '$Configuration.B3'
config_timestamp = '$Configuration.B5' 
config_log = '$Configuration.B7'

def desktop_model( ):
	desktop = XSCRIPTCONTEXT.getDesktop( )
	return desktop.getCurrentComponent( )

def sheet_by_name( name ):
	model = desktop_model( )
	return model.Sheets.getByName( name )

def cell_by_reference( reference ):
	values = reference.split( '.' )
	sheet = sheet_by_name( values[0].strip( '$' ) )
	return sheet.getCellRangeByName( values[1].strip( '$' ) )

def additional_equities( ):
	cell = cell_by_reference( config_equities )
	return cell.getString( )

def token_value( ):
	cell = cell_by_reference( config_token )
	return cell.getString( )

def log_cell( ):
	return cell_by_reference( config_log )

def log_clear( ):
	from datetime import datetime

	cell = log_cell( )
	today = datetime.today( )
	cell.setString( today.strftime( '%Y.%m.%d-%H:%M:%S : Reconciling from Questrade started!' ) )

def log_read( ):
	cell = log_cell( )
	return cell.getString( )

def log_write( message ):
	from datetime import datetime

	cell = log_cell( )
	today = datetime.today( )
	cell.setString( cell.getString( ) + '\n' + today.strftime( '%Y.%m.%d-%H:%M:%S : ' ) + message )

def log_traceback( ):
	from sys import exc_info
	from traceback import format_exception

	exc_type, exc_value, exc_traceback = exc_info( )
	lines = format_exception( exc_type, exc_value, exc_traceback )
	for line in lines:
		log_write( line )

def api_cache_connect( ):
	from questrade_api import Questrade

	try:
		q = Questrade( )
		timevalue = q.time
	except:
		q = None
		log_write( 'Failed to authenticate using cached token file.' )
	return q

def api_token_connect( ):
	from questrade_api import Questrade
	from os.path import expanduser, isfile
	from os import remove

	TOKEN_PATH = expanduser( '~/.questrade.json' )
	try:
		if isfile( TOKEN_PATH ):
			remove( TOKEN_PATH )
		q = Questrade( refresh_token = token_value( ) )
		timevalue = q.time
	except:
		q = None
		log_write( 'Failed to authenticate using refresh_token in cell {}'.format( config_token ) )
	return q

def api_connect( ):
	q = api_cache_connect( )
	if q is None:
		q = api_token_connect( )
	if q is None:
		raise RuntimeError( 'Failed to authenticate Questrade API.' )
	return q

def sortfield_by_index( index, ascending = True ):
	from com.sun.star.util import SortField

	sf = SortField( )
	sf.Field = index
	sf.SortAscending = ascending
	return sf

def property_value( name, value ):
	from com.sun.star.beans import PropertyValue

	pv = PropertyValue( )
	pv.Name = name
	pv.Value = value
	return pv

def format_cell( cell, format ):
	from com.sun.star.uno import RuntimeException

	document = XSCRIPTCONTEXT.getDocument( )
	try:
		cell.NumberFormat = document.NumberFormats.addNew( format, document.CharLocale )
	except RuntimeException:
		cell.NumberFormat = document.NumberFormats.queryKey( format, document.CharLocale, False )

def set_and_format_string( cell, value, format ):
	if value is None:
		set_and_format_string( cell, '', 'General' )
	else:
		cell.setString( value )
		format_cell( cell, format )

def string_to_date( value ):
	''' Sample date: '2019-07-02T00:00:00.000000-04:00'
	'''
	from datetime import date

	return date( int( value[0:4] ), int( value[5:7] ), int( value[8:10] ) )

def serial_date( value ):
	from datetime import date

	d = string_to_date( value )
	t = date( 1899, 12, 30 )
	delta = d - t

	return float( delta.days ) + float( delta.seconds ) / 86400

def within_monthly_span( value, span = 1 ):
	from datetime import date
	from calendar import monthrange

	d = string_to_date( value )
	t = date.today( )

	year = t.year + int( ( t.month + span - 1 ) / 12 )
	month = ( t.month + span - 1 ) % 12
	current, last = monthrange( year, month )

	t1 = date( t.year, t.month, 1 )
	t2 = date( year, month, last )

	return d >= t1 and d <= t2

def set_and_format_date( cell, value, format ):
	if value is None:
		set_and_format_string( cell, '', 'General' )
	elif isinstance( value, str ):
		cell.setValue( serial_date( value ) )
		format_cell( cell, format )
		if within_monthly_span( value ):
			cell.CellBackColor = 0xccffcc	# light green
	else:
		set_and_format_string( cell, value, 'General' )

def set_and_format_float( cell, value, format ):
	if value is None:
		set_and_format_string( cell, '', 'General' )
	else:
		cell.setValue( float( value ) )
		format_cell( cell, format )

def timestamp( ):
	from datetime import datetime

	now = datetime.now( )
	cell = cell_by_reference( config_timestamp )
	value = now.strftime( '%A, %d %b %Y %I:%M:%S %p' )
	set_and_format_string( cell, value, 'General' )

def message_box( message ):
	from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
	from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
	from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL

	model = desktop_model( )
	parentwin = model.CurrentController.Frame.ContainerWindow
	box = parentwin.getToolkit( ).createMessageBox( parentwin, MESSAGEBOX, BUTTONS_OK, 'QuestradeReconcile', message )
	result = box.execute( )

class Spreadsheet( ):
	setvalue = {
		's': lambda c, v: set_and_format_string( c, v, '@' ),
		'd': lambda c, v: set_and_format_date( c, v, 'MMM D, YYYY' ),
		'b': lambda c, v: set_and_format_float( c, v, 'BOOLEAN' ),
		'n': lambda c, v: set_and_format_float( c, v, 'General' ),
		'n0': lambda c, v: set_and_format_float( c, v, '#,##0;[RED]-#,##0' ),
		'n1': lambda c, v: set_and_format_float( c, v, '#,##0.0;[RED]-#,##0.0' ),
		'n2': lambda c, v: set_and_format_float( c, v, '#,##0.00;[RED]-#,##0.00' ),
		'n3': lambda c, v: set_and_format_float( c, v, '#,##0.000;[RED]-#,##0.000' ),
		'n4': lambda c, v: set_and_format_float( c, v, '#,##0.0000;[RED]-#,##0.0000' ),
		'n5': lambda c, v: set_and_format_float( c, v, '#,##0.00000;[RED]-#,##0.00000' ),
		'n6': lambda c, v: set_and_format_float( c, v, '#,##0.000000;[RED]-#,##0.000000' )
	}

	q = api_connect( )

	def __init__( self, name, fields ):
		self.name = name
		self.fields = fields
		self.sheet = sheet_by_name( self.name )
		self.range = self.sheet.getCellRangeByName( 'A1:AMJ1048576' )
		self.range.clearContents( 1 | 2 | 4 | 8 | 16 | 32 | 64 | 128 | 256 | 512 )
		self.rows = 0

	def add_row( self, data ):
		column = 0
		for field in self.fields:
			field_name = field[0]
			field_type = field[1]
			if self.rows == 0:
				name_cell = self.sheet.getCellByPosition( column, 0 )
				try:
					name_cell.setString( field_name )
				except:
					log_traceback( )
				name_cell.CellBackColor = 0xb4c7dc	# light blue / grey
			if field_name in data:
				value_cell = self.sheet.getCellByPosition( column, self.rows + 1 )
				try:
					self.setvalue[field_type]( value_cell, data[ field_name ] )
				except:
					log_write( '{}::add_row( n={} t={} r={} c={} v={} )'.format( self.name, field_name, field_type, self.rows, column, data[ field_name] ) )
					log_traceback( )
			column = column + 1
		self.rows = self.rows + 1

	def sort_by_indicies( self, indicies, ascending = True ):
		from uno import Any

		sort_fields = tuple( sortfield_by_index( index, ascending ) for index in indicies )
		property_fields = property_value( 'SortFields', Any( '[]com.sun.star.util.SortField', sort_fields ) )
		property_header = property_value( 'HasHeader', True )
		self.range.sort( [ property_fields, property_header ] )

class Accounts( Spreadsheet ):
	def __init__( self ):
		accounts_name = 'Accounts'
		accounts_fields = [
			( 'number', 'n', ),
			( 'type', 's', ),
			( 'clientAccountType', 's', ),
			( 'status', 's', ),
			( 'isPrimary', 'b', ),
			( 'isBilling', 'b', )
		]
		super( ).__init__( accounts_name, accounts_fields )

	def fetch( self ):
		try:
			quest_accounts = self.q.accounts
		except:
			log_write( 'Accounts::fetch( ) failed' )
			log_traceback( )
			return
		for quest_account in quest_accounts['accounts']:
			yield quest_account

class Balances( Spreadsheet ):
	CAD = 0
	USD = 1

	def __init__( self ):
		balances_name = 'Balances'
		balances_fields = [
			( 'balanceType', 's', ),
			( 'account', 'n', ),
			( 'accountType', 's', ),
			( 'currency', 's', ),
			( 'cash', 'n2', ),
			( 'marketValue', 'n2', ),
			( 'totalEquity', 'n2', ),
			( 'buyingPower', 'n2', ),
			( 'maintenanceExcess', 'n2', ),
			( 'isRealTime', 'b', )
		]
		super( ).__init__( balances_name, balances_fields )

	def fetch( self, account_number, account_type ):
		try:
			quest_balances = self.q.account_balances( account_number )
		except:
			log_write( 'Balances::fetch( {}, {} ) failed'.format( account_number, account_type ) )
			log_traceback( )
			return
		for balanceType in [ 'combinedBalances', 'perCurrencyBalances' ]:
			for currency in [ Balances.CAD, Balances.USD ]:
				quest_balances[balanceType][currency]['balanceType'] = balanceType
				quest_balances[balanceType][currency]['account'] = account_number
				quest_balances[balanceType][currency]['accountType'] = account_type
				yield quest_balances[balanceType][currency]

	def default_sort( self ):
		# balanceType (Column A = 0), account (Column B = 1), currency (Column D = 3)
		self.sort_by_indicies( ( 0, 1, 3, ) )

class Positions( Spreadsheet ):
	def __init__( self ):
		positions_name = 'Positions'
		positions_fields = [
			( 'account', 'n', ),
			( 'accountType', 's', ),
			( 'currency', 's', ),
			( 'symbol', 's', ),
			( 'symbolId', 'n', ),
			( 'openQuantity', 'n3', ),
			( 'currentPrice', 'n2', ),
			( 'currentMarketValue', 'n2', ),
			( 'averageEntryPrice', 'n3', ),
			( 'totalCost', 'n2', ),
			( 'openPnl', 'n2', ),
			( 'dayPnl', 'n2', ),
			( 'closedQuantity', 'n3', ),
			( 'closedPnl', 'n2', ),
			( 'isUnderReorg', 'b', ),
			( 'isRealTime', 'b', )
		]
		super( ).__init__( positions_name, positions_fields )

	def fetch( self, account_number, account_type ):
		try:
			quest_positions = self.q.account_positions( account_number )
		except:
			log_write( 'Positions::fetch( {}, {} ) failed'.format( account_number, account_type ) )
			log_traceback( )
			return
		for quest_position in quest_positions['positions']:
			quest_position['account'] = account_number
			quest_position['accountType'] = account_type
			yield quest_position

	def default_sort( self ):
		# account (column A = 0), currency (column C = 2), symbol (column D = 3)
		self.sort_by_indicies( ( 0, 2, 3, ) )

class Equities( Spreadsheet ):
	def __init__( self ):
		equities_name = 'Equities'
		equities_fields = [
			( 'account', 'n', ),
			( 'accountType', 's', ),
			( 'currency', 's', ),
			( 'symbol', 's', ),
			( 'symbolId', 'n', ),
			( 'description', 's', ),
			( 'listingExchange', 's', ),
			( 'securityType', 's', ),
			( 'prevDayClosePrice', 'n2', ),
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
			( 'lowPrice52', 'n2', ),
			( 'highPrice52', 'n2', ),
			( 'tradeUnit', 'b', )
		]
		super( ).__init__( equities_name, equities_fields )

	def fetch( self, symbol_names ):
		try:
			if len( symbol_names ) == 0:
				log_write( 'Equities::fetch( len( symbol_names ) is 0 )' )
				return
			quest_equities = self.q.symbols( names = symbol_names )
		except:
			log_write( 'Equities::fetch( {} ) failed'.format( symbol_names ) )
			log_traceback( )
			return
		for quest_equity in quest_equities['symbols']:
			yield quest_equity

	def fetch_unique( self, account_number, account_type, symbol_id ):
		try:
			quest_equities = self.q.symbol( symbol_id )
		except:
			log_write( 'Equities::fetch_unique( {}, {}, {} ) failed'.format( account_number, account_type, symbol_id ) )
			log_traceback( )
			return None
		quest_equity = quest_equities['symbols'][0]
		quest_equity['account'] = account_number
		quest_equity['accountType'] = account_type
		return quest_equity

	def default_sort( self ):
		# account (Column A = 0), currency (Column C = 2), symbol (Column C = 3)
		self.sort_by_indicies( ( 0, 2, 3, ) )

class Activities( Spreadsheet ):
	from datetime import date, timedelta

	startDate = ( date.today( ) - timedelta( days = 30 ) ).isoformat( ) + 'T00:00:00-04:00'
	endDate = date.today( ).isoformat( ) + 'T00:00:00-04:00'

	def __init__( self ):
		activities_name = 'Activities'
		activities_fields = [
			( 'account', 'n', ),
			( 'accountType', 's', ),
			( 'currency', 's', ),
			( 'transactionDate', 'd', ),
			( 'symbol', 's', ),
			( 'symbolId', 'n', ),
			( 'type', 's', ),
			( 'action', 's', ),
			( 'quantity', 'n3', ),
			( 'price', 'n4', ),
			( 'grossAmount', 'n2', ),
			( 'commission', 'n2', ),
			( 'netAmount', 'n2', ),
			( 'tradeDate', 'd', ),
			( 'settlementDate', 'd', ),
			( 'description', 's', )
		]
		super( ).__init__( activities_name, activities_fields )

	def fetch( self, account_number, account_type ):
		try:
			quest_activities = self.q.account_activities( account_number, startTime = Activities.startDate, endTime = Activities.endDate )
		except:
			log_write( 'Activities::fetch( {}, {} ) failed'.format( account_number, account_type ) )
			log_traceback( )
			return
		for quest_activity in quest_activities['activities']:
			quest_activity['account'] = account_number
			quest_activity['accountType'] = account_type
			yield quest_activity

	def default_sort( self ):
		# account (Column A = 0), currency (Column C = 2), transactionDate (Column D = 3), Symbol (Column E = 4)
		self.sort_by_indicies( ( 0, 2, 3, 4, ) )

class Dividends( Spreadsheet ):
	from re import compile, VERBOSE, MULTILINE

	tsx_frequency_pattern = compile( r"<p>Frequency:\s*(?P<frequency>[^<]*)</p>", VERBOSE | MULTILINE )
	tsx_dividend_patttern = compile(
		r'''<tr>\s*
			<td>\s*(?:<i>)?\s*(?P<dividend>\d{4}-\d{2}-\d{2})\s*(?:</i>)?\s*</td>\s*
			<td>\s*(?:<i>)?\s*(?P<payout>\d{4}-\d{2}-\d{2})\s*(?:</i>)?\s*</td>\s*
			<td>\s*(?:<i>)?\s*\$(?P<amount>\d+\.\d+)?(?P<note>\*\*)?\s*(?:</i>)?\s*</td>\s*
			<td>.*?</td>\s*
		</tr>''',
		VERBOSE | MULTILINE )
	nasdaq_dividend_pattern = compile(
		r'''<tr>\s*
			<td>\s*<span[^>]*>(?P<dividend>\d+/\d+/\d+)\s*</span>\s*</td>\s*
			<td>(?P<note>\w+)</td>\s*
			<td>\s*<span[^>]*>(?P<amount>\d+\.\d+)</span>\s*</td>\s*
			<td>\s*<span[^>]*>(?P<declaration>\d+/\d+/\d+)\s*</span>\s*</td>\s*
			<td>\s*<span[^>]*>(?P<record>\d+/\d+/\d+)\s*</span>\s*</td>\s*
			<td>\s*<span[^>]*>(?P<payout>\d+/\d+/\d+)\s*</span>\s*</td>\s*
		</tr>''',
		VERBOSE | MULTILINE )
	handler = {
		'TSX':    lambda o, e: o.fetch_tsx( e ),
		'NASDAQ': lambda o, e: o.fetch_nasdaq( e ),
		'BATS':   lambda o, e: o.fetch_nasdaq( e ),
		'ARCA':   lambda o, e: o.fetch_nasdaq( e ),
		'':       lambda o, e: o.fetch_null( e ),
		None:     lambda o, e: o.fetch_null( e )
	}

	def __init__( self ):
		dividends_name = "Dividends"
		dividends_fields = [
			( 'account', 'n' ),
			( 'accountType', 's' ),
			( 'symbol', 's' ),
			( 'symbolId', 'n' ),
			( 'currency', 's' ),
			( 'frequency', 's' ),
			( 'dividend', 'd' ),
			( 'payout', 'd' ),
			( 'amount', 'n6' ),
			( 'note', 's' ),
			( 'quantity', 'n3' ),
			( 'income', 'n2' )
		]
		super( ).__init__( dividends_name, dividends_fields )

	def cartesian_product( self, dividend, quest_equity ):
		dividend['account'] = quest_equity['account']
		dividend['accountType'] = quest_equity['accountType']
		dividend['symbol'] = quest_equity['symbol']
		dividend['symbolId'] = quest_equity['symbolId']
		dividend['currency'] = quest_equity['currency']
		dividend['quantity'] = None
		dividend['income'] = None
		return dividend

	def reformat_date( self, s ):
		''' Sample 12/24/2018 -> 2019-12-24
		'''
		from datetime import date
		u = s.split( '/' )
		d = date( int( u[2] ), int( u[0] ), int( u[1] ) )
		return d.strftime( '%Y-%m-%d' )

	def reformat_dates( self, dividend ):
		dividend['payout'] = self.reformat_date( dividend['payout'] )
		dividend['dividend'] = self.reformat_date( dividend['dividend'] )
		dividend['record'] = self.reformat_date( dividend['record'] )
		dividend['declaration'] = self.reformat_date( dividend['declaration'] )
		return dividend

	def url_open( self, url ):
		from urllib.request import urlopen
		try:
			response = urlopen( url )
			html = response.read( )
			return html.decode( 'utf-8' )
		except:
			log_write( 'Dividends::url_open( {} ) failed'.format( url ) )
			log_traceback( )
			return

	def fetch( self, quest_equity ):
		try:
			listing_exchange = quest_equity['listingExchange']
			generator = Dividends.handler[listing_exchange]( self, quest_equity )
		except:
			log_write( 'Dividends::fetch( {} ) failed'.format( quest_equity ) )
			log_traceback( )
			return
		return generator

	def fetch_tsx( self, quest_equity ):
		symbol = quest_equity['symbol'].replace( '.TO', '' )
		url = 'https://dividendhistory.org/payout/tsx/{}/'.format( symbol )
		data = self.url_open( url )

		frequency_match = Dividends.tsx_frequency_pattern.search( data )
		frequency = frequency_match.group( 'frequency' ) if frequency_match is not None else None

		for match in Dividends.tsx_dividend_patttern.finditer( data ):
			dividend = match.groupdict( )
			self.cartesian_product( dividend, quest_equity )
			dividend['frequency'] = frequency
			yield dividend

	def fetch_nasdaq( self, quest_equity ):
		symbol = quest_equity['symbol']
		url = 'https://www.nasdaq.com/symbol/{}/dividend-history'.format( symbol )
		data = self.url_open( url )

		for match in Dividends.nasdaq_dividend_pattern.finditer( data ):
			dividend = match.groupdict( )
			self.cartesian_product( dividend, quest_equity )
			self.reformat_dates( dividend )
			dividend['frequency'] = None
			yield dividend

	def fetch_null( self, quest_equity ):
		return { }

	def default_sort( self ):
		# dividend (Column G = 6), payout (Column H = 7)
		self.sort_by_indicies( ( 6, 7, ), ascending = False )

def QuestradeReconcile( *args ):
	timestamp( )
	log_clear( )

	try:
		desktop_model( ).lockControllers( )
		desktop_model( ).addActionLock( )

		accounts = Accounts( )
		balances = Balances( )
		activities = Activities( )
		positions = Positions( )
		equities = Equities( )
		dividends = Dividends( )

		for quest_account in accounts.fetch( ):
			accounts.add_row( quest_account )
			for quest_balance in balances.fetch( quest_account['number'], quest_account['type'] ):
				balances.add_row( quest_balance )
			for quest_activity in activities.fetch( quest_account['number'], quest_account['type'] ):
				activities.add_row( quest_activity )
			for quest_position in positions.fetch( quest_account['number'], quest_account['type'] ):
				quest_equity = equities.fetch_unique( quest_account['number'], quest_account['type'], quest_position['symbolId'] )
				if quest_equity is not None:
					quest_position['currency'] = quest_equity['currency']
					equities.add_row( quest_equity )
					for dividend in dividends.fetch( quest_equity ):
						if within_monthly_span( dividend['payout'], 3 ):
							dividend['quantity'] = quest_position['openQuantity']
							dividend['income'] = float( dividend['amount'] ) * float( quest_position['openQuantity'] )
						dividends.add_row( dividend )
				positions.add_row( quest_position )

		for quest_equity in equities.fetch( additional_equities( ) ):
			equities.add_row( quest_equity )
			for dividend in dividends.fetch( quest_equity ):
				dividends.add_row( dividend )

		activities.default_sort( )
		balances.default_sort( )
		positions.default_sort( )
		equities.default_sort( )
		dividends.default_sort( )

	except:
		log_traceback( )

	finally:
		desktop_model( ).removeActionLock( )
		desktop_model( ).unlockControllers( )

	log_write( 'Reconciling from Questrade completed!' )
	message_box( log_read( ) )

	return None

g_exportedScripts = QuestradeReconcile,
