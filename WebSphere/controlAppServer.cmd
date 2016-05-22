call \WebSphere\profiles\path\bin\setupCmdLine.bat
rem syntax: <batch> <node> <server> {start|stop}
\WebSphere\profiles\path\bin\wsadmin -lang jython -f crontrolAppServer.py -conntype SOAP -user user_id -password passw %2 %3 %4