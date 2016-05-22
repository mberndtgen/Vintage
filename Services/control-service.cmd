@echo off
:: call:
:: control-service.cmd [-super] [-start] [-stop] [-restart] [-killall] -service <servicename> [-account <accountname>]

:: parsing: see http://stackoverflow.com/questions/14286457/using-parameters-in-batch-files-at-dos-command-line
:: elevated mode: see http://stackoverflow.com/questions/1894967/how-to-request-administrator-access-inside-a-batch-file/10052222#10052222

:: save script name with path; necessary due to the shifts below (otherwise %~s0 becomes the last param)
set script=%~s0

:parse
if "%~1"=="" goto endparse
if "%~1"=="-super" set Mode=super
if "%~1"=="-start" set Start=yes
if "%~1"=="-clean" set Clean=yes
if "%~1"=="-stop" set Stop=yes
if "%~1"=="-restart" set Restart=yes
if "%~1"=="-killall" set Killall=yes
if "%~1"=="-service" set ServiceName=%2
if "%~1"=="-account" set Serviceaccount=%2
shift
goto parse
:endparse

if "%Clean%"=="yes" (
	goto clean
)

if "%Mode%"=="super" (
	goto elevate
) else (
	goto stopservice
)

:elevate
:: Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

:: If error flag set, we do not have admin privileges.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\get-admin.vbs"
    set params=%*
	echo %params%
    echo UAC.ShellExecute "cmd.exe", "/c %script% %params%", "", "runas", 1 >> "%temp%\get-admin.vbs"

    "%temp%\get-admin.vbs"
    del "%temp%\get-admin.vbs"
    exit /B

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"

:stopservice
if "%Start%"=="yes" (
	goto restartservice
)
sc config %ServiceName% start= demand
ping -n 10 localhost>nul

:LOOP1
echo on
	SET /A COUNTER+=1
	IF %COUNTER% GTR 3 GOTO :ERRORTEST99
	echo stopping service # %COUNTER% #, please wait
	echo.
	sc stop %ServiceName%
	ping -n 15 localhost>nul
	SC query %ServiceName% | FIND "STATE" | FIND "STOPPED"
echo off
	If %ERRORLEVEL% EQU 1 GOTO :LOOP1
GOTO :GOON9876
:ERRORTEST99
	SC query %ServiceName% | FIND "STATE" | FIND "STOP_PENDING"
	If %ERRORLEVEL% EQU 1 GOTO :ERROR
	if %Killall%=="yes" (
		echo Service hangs on "stopping", now killing %Serviceaccount% processes ...
		taskkill /f /fi "USERNAME eq %Serviceaccount%"
	) else (
		echo Service hangs on "stopping", now killing %ServiceName% 
		taskkill /f /fi "SERVICE eq %ServiceName%"
	)
	echo.
	ping -n 5 localhost>nul
	echo.
	SC query %ServiceName% | FIND "STATE" | FIND "STOPPED"
	If %ERRORLEVEL% EQU 1 GOTO :ERROR	
:GOON9876

if "%Restart%"=="yes" (
	goto restartservice
) else (
	exit /B
)

:RESTARTSERVICE
sc start %ServiceName%
ping -n 30 localhost>nul
sc config %ServiceName% start= auto

:CLEAN
set Mode=
set Clean=
set Start=
set Restart=
set Killall=
set ServiceName=
set Serviceaccount=
set COUNTER=0
exit /B

:ERROR
echo.
echo Error stopping service. Aborting.
echo.
goto :CLEAN
