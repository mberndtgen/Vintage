@echo off

rem params
set server=%1
set stage=%2
set user=%3
set pwd=%4

set SCOMserver=server
set SUFFIX=
if /i "%user%" != "" set suffix="-u %user% -p %pwd%"

if /i "%server%" == "" goto syntax
if /i "%stage%" == "" goto syntax


if "%stage%" == "stage-dev" goto action
if "%stage%" == "stage-prod" set SCOMserver=server

mmtool -s"%SCOMserver%" -a -m"on" -h"%server%" -c"Windows Hotfixes"
sleep 120

:action
rem planned, windows, hotfix
rem description see: http://msdn.microsoft.com/en-us/library/windows/desktop/aa376885(v=vs.85).aspx
psshutdown -e p:131072:18 -f -n 30 -r \\%server% %SUFFIX%
if "%stage%" == "stage-dev" goto end

sleep 120

mmtool -s"%SCOMserver%" -a -m"off" -h"%server%" 
goto end

:syntax
echo Syntax: "rreboot <server> <stage> [<user> <pwd>]"
echo <stage> := stage-prod | stage-dev

:end
echo Done.
