@echo off
REM https://git-scm.com/book/en/v2/Git-Tools-Submodules

if exist "..\questrade-api-master.zip" (
	if exist "questrade-api" (
		del /Q "questrade-api"
		for /D %%p in ( "questrade-api\*.*" ) do rmdir "%%p" /s /q
		rmdir "questrade-api"
	)
    powershell.exe -nologo -noprofile -command "& { Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('..\questrade-api-master.zip', '.'); }"
	ren questrade-api-master questrade-api
)

if exist "questrade-api-master.zip" (
	if exist "questrade-api" (
		del /Q "questrade-api"
		for /D %%p in ( "questrade-api\*.*" ) do rmdir "%%p" /s /q
		rmdir "questrade-api"
	)
    powershell.exe -nologo -noprofile -command "& { Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('questrade-api-master.zip', '.'); }"
	ren questrade-api-master questrade-api
)

if not exist "questrade-api\questrade_api" (
	echo FAILED: Missing questrade-api\questrade_api directory! Quitting!
	exit /B
)

if not exist "C:\Program Files\LibreOffice\share\Scripts\python" (
	echo FAILED: Missing "C:\Program Files\LibreOffice\share\Scripts\python". Is LibreOffice installed? Quitting!
	exit /B
)

if not exist "QuestradeReconcile.py" (
	echo FAILED: Missing "QuestradeReconcile.py" from current directory! Quitting!
	exit /B
)

if exist "C:\Program Files\LibreOffice\share\Scripts\python\QuestradeReconcile.py" (
	del /Q "C:\Program Files\LibreOffice\share\Scripts\python\QuestradeReconcile.py"
)

if exist "C:\Program Files\LibreOffice\share\Scripts\python\QuestradeReconcile.py" (
	echo FAILED: Could not remove old "C:\Program Files\LibreOffice\share\Scripts\python\QuestradeReconcile.py"! Not running as Administrator? Quitting!
	exit /B
)

copy "QuestradeReconcile.py" "C:\Program Files\LibreOffice\share\Scripts\python\QuestradeReconcile.py" /Y > NUL:

if not exist "C:\Program Files\LibreOffice\share\Scripts\python\QuestradeReconcile.py" (
	echo FAILED: Missing "C:\Program Files\LibreOffice\share\Scripts\python\QuestradeReconcile.py"! Copy didn't work! Quitting!
	exit /B
)

echo INFO: QuestradeReconcile.py copied to LibreOffice Scripts.

if not exist "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages" (
	echo FAILED: Missing "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages"! LibreOffice install not as expected! Quitting!
	exit /B
)

if exist "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages\questrade_api" (
	del /Q "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages\questrade_api"
	for /D %%p in ( "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages\questrade_api\*.*" ) do rmdir "%%p" /s /q
	rmdir "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages\questrade_api"
)

if exist "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages\questrade_api" (
	echo FAILED: Couldn't delete existing "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages\questrade_api" Quitting!
	exit /B
)

xcopy "questrade-api\questrade_api" "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages\questrade_api" /S /Y /I > NUL:

if not exist "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages\questrade_api" (
	echo FAILED: Couldn't copy "questrade-api\questrade_api" "C:\Program Files\LibreOffice\program\python-core-3.5.7\lib\site-packages\questrade_api" Quitting!
	exit /B
)

echo INFO: questrade_api copied to Python site-packages.

echo INFO: Success!!!
