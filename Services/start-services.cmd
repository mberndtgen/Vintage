echo Start Services
call control-service.cmd -start -service "WPCSvc"
call control-service.cmd -start -service "Fax"
call control-service.cmd -start -service "HttpAnalyzerV6 DllInjectService"