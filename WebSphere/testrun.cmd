@echo off

rem 
rem Test Stage
rem 
set username=user_id1
FOR  /F "usebackq delims==" %%i IN (`cscript //nologo setenv.vbs -u %username% -s`) DO set password=%%i
cscript listRunningServers.vbs -s server -e T -p 7101 -u "-user user_id1 -password %password%" -v -c .\wasnodes.conf
rem 
rem Approval Stage
rem 
set username=user_id2
FOR  /F "usebackq delims==" %%i IN (`cscript //nologo setenv.vbs -u %username% -s`) DO set password=%%i
cscript listRunningServers.vbs -s server -e A -port 7201 -u "-user user_id2 -password %password%" -v -c .\wasnodes.conf

rem 
rem Production Stage
rem 
set username=user_id3
FOR  /F "usebackq delims==" %%i IN (`cscript //nologo setenv.vbs -u %username% -s`) DO set password=%%i
cscript listRunningServers.vbs -s server -e P -port 7301 -u "-user user_id3 -password %password%" -v -c .\wasnodes.conf

rem
rem stop app servers
rem
rem cscript AppServerControl.vbs -m stop -v -t -f running-appservers-T.txt
rem
rem start app servers
rem
rem cscript AppServerControl.vbs -m start -v -t -f running-appservers-T.txt

set username=
set password=