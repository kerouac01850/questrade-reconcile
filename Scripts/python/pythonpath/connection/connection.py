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
from questrade_api import Questrade
from spreadsheet import Configuration

class Connection( ):
	def __init__( self, context ):
		Configuration.set_context( context )
		self.__questrade = self.__questrade_cache_connect( )
		if self.__questrade is None:
			self.__questrade = self.__questrade_token_connect( )
		if self.__questrade is None:
			ctrl = 'MBOX: Failed to authenticate. Generate new Questrade token at {}.'
			logging.critical( ctrl.format( Configuration.get_token_ref( ) ) )
			raise RuntimeError( 'Unable to authenticate connection with Questrade.' )

	def __questrade_cache_connect( self ):
		if len( Configuration.get_apicache( ) ) == 0:
			return None
		try:
			logging.info( 'MBOX: Authenticate using cached credentials at {}'.format( Configuration.get_cache_ref( ) ) )
			q = Questrade(
				logger = logging.debug,
				storage_adaptor = ( lambda : Configuration.get_apicache( ), lambda s: Configuration.set_apicache( s ) ) )
			Configuration.set_timestamp( q.time['time'] )
		except:
			logging.exception( 'MBOX: Authenticate using credentials FAILED!!!' )
			q = None
		return q

	def __questrade_token_connect( self ):
		if len( Configuration.get_token( ) ) == 0:
			return None
		try:
			logging.info( 'MBOX: Authenticate using Questrade token at {}'.format( Configuration.get_token_ref( ) ) )
			q = Questrade(
				logger = logging.debug,
				storage_adaptor = ( lambda : Configuration.get_apicache( ), lambda s: Configuration.set_apicache( s ) ),
				refresh_token = Configuration.get_token( ) )
			Configuration.set_timestamp( q.time['time'] )
		except:
			logging.exception( 'MBOX: Authenticate using Questrade token FAILED!!!' )
			q = None
		return q

	@property
	def questrade( self ):
		return self.__questrade

	def finalize_ratelimits( self ):
		remaining, reset = self.questrade.ratelimit
		Configuration.set_remaining( remaining )
		Configuration.set_reset( reset )
