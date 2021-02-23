''' QuestradeReconcile is free software: you can redistribute and/or modify it under the terms of the GNU General Public
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
