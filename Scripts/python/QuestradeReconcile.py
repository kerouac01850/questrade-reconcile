'''	The QuestradeReconcile macro is python code that uses the Questrade application programming interface (API) to fetch
	account, position, balance, equity, and 30 day activity into a LibreOffice spreadsheet file.

	The spreadsheet file must have six sheets with the following names: Accounts, Positions, Balances, Equities, Activity,
	and Configuration. Except for the Configuration sheet any existing data on these sheets will be cleared when the
	QuestradeReconcile python macro is run.

	Data from the Questrade on-line platform will be used to update these sheets. The sheets can be moved within the file:
	sheet order does not affect anything. Any other sheets in the file are ignored.
	
	QuestradeReconcile is free software: you can redistribute and/or modify it under the terms of the GNU General Public
	License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later
	version.

	QuestradeReconcile is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
	warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU General Public License along with this program. If not, see
	https://www.gnu.org/licenses/.

	Copyright (C) 2019 - 2021 kerouac01850
'''

import sys
from os.path import dirname
from unohelper import fileUrlToSystemPath

document = XSCRIPTCONTEXT.getDocument()
url = fileUrlToSystemPath( '{}/{}'.format( document.URL, 'Scripts/python/pythonpath' ) )
sys.path.insert( 0, url )

import logging
from spreadsheet import Spreadsheet, Configuration, Accounts, Activities, Balances, Equities, Positions
from connection import Connection

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
		logging.info( 'MBOX: {} from Questrade started!'.format( title ) )
		macro( )
		logging.info( 'MBOX: {} complete!'.format( title ) )
		Configuration.set_log( logging_buffer.getvalue( ) )

		manager = getComponentContext( ).ServiceManager
		dispatcher = manager.createInstance( "com.sun.star.frame.DispatchHelper" )
		document = model.CurrentController.Frame
		dispatcher.executeDispatch( document, ".uno:CalculateHard", "", 0, ( [ ] ) )

		parentwin = model.CurrentController.Frame.ContainerWindow
		box = parentwin.getToolkit( ).createMessageBox( parentwin, MESSAGEBOX, BUTTONS_OK, title, message_buffer.getvalue( ) )
		result = box.execute( )

	except:
		logging.exception( 'MBOX: {} failed because of an exception!'.format( title ) )

	finally:
		model.removeActionLock( )
		model.unlockControllers( )

def reconcile_everything( qt ):
	Accounts( ).reconcile( qt )
	Balances( ).reconcile( qt )
	Activities( ).reconcile( qt )
	Positions( ).reconcile( qt )
	Equities( ).reconcile( qt )

def Reconcile( *args ):
	qt = Connection( XSCRIPTCONTEXT )
	RunMacro( 'Reconcile', lambda: reconcile_everything( qt ) )
	qt.finalize_ratelimits( )
	return None

def ReconcileAccounts( *args ):
	qt = Connection( XSCRIPTCONTEXT )
	RunMacro( 'ReconcileAccounts', lambda: Accounts( ).reconcile( qt ) )
	qt.finalize_ratelimits( )
	return None

def ReconcileBalances( *args ):
	qt = Connection( XSCRIPTCONTEXT )
	RunMacro( 'ReconcileBalances', lambda: Balances( ).reconcile( qt ) )
	qt.finalize_ratelimits( )
	return None

def ReconcileActivities( *args ):
	qt = Connection( XSCRIPTCONTEXT )
	RunMacro( 'ReconcileActivities', lambda: Activities( ).reconcile( qt ) )
	qt.finalize_ratelimits( )
	return None

def ReconcilePositions( *args ):
	qt = Connection( XSCRIPTCONTEXT )
	RunMacro( 'ReconcilePositions', lambda: Positions( ).reconcile( qt ) )
	qt.finalize_ratelimits( )
	return None

def ReconcileEquities( *args ):
	qt = Connection( XSCRIPTCONTEXT )
	RunMacro( 'ReconcileEquities', lambda: Equities( ).reconcile( qt ) )
	qt.finalize_ratelimits( )
	return None

g_exportedScripts = Reconcile,ReconcileAccounts,ReconcileBalances,ReconcileActivities,ReconcileEquities,ReconcilePositions
