'''	The Configuration sheet must have the following cells defined for the QuestradeReconcile macro to function correctly.
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
from spreadsheet import Spreadsheet

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
	def get_token_ref( cls ):
		return cls.__token

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
	def get_cache_ref( cls ):
		return cls.__apicache

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
