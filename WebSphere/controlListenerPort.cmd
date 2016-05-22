call E:\Data\WebSphere\profiles\path\bin\setupCmdLine.bat
rem syntax: <batch> <node> <server> {start|stop}
E:\Data\WebSphere\profiles\path\bin\wsadmin -lang jython -f E:\Program\ShrdApps\WebSphereScripting\controlListenerPort.py -conntype SOAP -user user_id -password passw %2 %3 %4