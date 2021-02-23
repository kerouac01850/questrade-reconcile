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
	__context = None

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
	def set_context( cls, context ):
		Spreadsheet.__context = context

	@classmethod
	def get_context( cls ):
		return Spreadsheet.__context

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
		document = cls.get_context( ).getDocument( )
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
		desktop = cls.get_context( ).getDesktop( )
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

	def column_from( self, oCell ):
		return oCell.getCellAddress( ).Column

	def column_by_name( self, name ):
		return self.column_from( self.find_cell( "A1:AMJ1", name ) )

	def alias_for( self, name, row ):
		return chr( ord( 'A' ) + self.column_by_name( name ) ) + str( row + 1 )

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
