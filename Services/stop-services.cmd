echo Stop Services
call control-service.cmd -stop -service "WPCSvc"
call control-service.cmd -stop -service "Fax"
call control-service.cmd -stop -service "HttpAnalyzerV6 DllInjectService"