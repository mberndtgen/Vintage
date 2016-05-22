'
' AppServerControl
'
' PURPOSE: safely shutdown given WebSphere server by stopping all WAS Application servers  on that machine
' LAST EDITED: 11-08-10
' AUTHOR(S): Berndtgen, Düsseldorf, 2010
' SYNTAX: cscript AppServerControl.vbs -s[erver] <server> -p [dmgrportnumber] -u " -user <dmgruser> -password <dmgruserpwd>" [-v] [-t] -c [conffile] [[-]?] 
' VERSION: 2
'

dim aAppServer(100, 4) ' arry of was application servers
dim LogFilename ' name of logfile
dim LogFile ' as file object
dim InputFilename ' name of input file (containing appserver list)
dim Mode ' as string; 'start' or 'stop'
dim InputFile ' as file object
dim bVerbose ' as boolean
dim bTerminate ' as boolean
dim bLogging ' as boolean
dim sEnvironment ' as string

'
'main
'
Init
ReadConfigFile(InputFilename)
ControlAppServer(Mode)
'CleanUp

sub Init
	Set oShell = WScript.CreateObject("wscript.shell") ' give us a cmd shell
	bVerbose = false
	bLogging = false
	bTerminate = false
	sEnvironment = ""
	LogFilename = "logfile-shutdown.txt"
	Mode = ""
	
	on error resume next
	
	dim oArgumente
	set oArgumente = WScript.Arguments
	
	if oArgumente.Count < 2 then
	  	ShowUsage
	 	WScript.Quit
	else
	 	do 
	      	if UCase(oArgumente(i)) = "-F" or UCase(oArgumente(i)) = "-FILE" then
	  	  		' input file 
	  	  		i = i + 1
	  	  		InputFilename = oArgumente(i)
				if err.number<>0 then 
					WScript.Echo "ERROR: invalid input data (input file). Aborting."
					WScript.Quit
				end if  	  
			elseif UCase(oArgumente(i)) = "-M" or UCase(oArgumente(i)) = "-MODE" then
	  	  		' input file 
	  	  		i = i + 1
	  	  		Mode = oArgumente(i)
				if UCase(Mode)<>"STOP" and UCase(Mode)<>"START" then 
					WScript.Echo "ERROR: invalid input data (mode). Aborting."
					WScript.Quit
				end if  	
				Mode = LCase(Mode)
			elseif UCase(oArgumente(i)) = "-V" or UCase(oArgumente(i)) = "-VERBOSE" then
	  	  		' verbose mode
	  	  		bVerbose = true
				WScript.Echo "Using verbose mode."
			elseif UCase(oArgumente(i)) = "-T" or UCase(oArgumente(i)) = "-TERMINATE" then
	  	  		' terminate app servers (instead of just stopping them)
	  	  		bTerminate = true
				WScript.Echo "Using terminate mode."
			end if
	        i = i + 1    
	  	loop until i>=oArgumente.Count
		
		if InputFilename="" or Mode="" then 
			ShowUsage
			WScript.Quit
		end if
	end if
end sub

'	
' ReadConfigFile - read input file into array of ConfigDictionaries
'
sub ReadConfigFile(ByVal InputFilename)
	dim oFilesys, oFile ' FilesystemObject, Readfile
	dim Contents ' content of oFile
	dim InputData ' as array
	dim i ' Array counter
	dim cStage ' Stage identifier (T,A,P)
	
	i = 0
	set oFilesys = CreateObject("Scripting.FileSystemObject")
	if not oFilesys.FileExists(InputFilename) then 
		' no file? abort
		printf "ERROR: Input File does not exist. Aborting."
		WScript.Quit
	end if

	on error resume next

	set oFile = oFilesys.OpenTextFile(InputFilename, 1, false)
	if err.number<>0 then
		printf "ERROR: Can't read input file: " & Err.Message & ". Aborting."
		WScript.Quit
	end if

	on error goto 0
	
	dim LineCounter ' as integer
	LineCounter = 1
	
	do while not oFile.AtEndOfStream 
		Contents = oFile.ReadLine
		
		if left(Contents, 1) <> "#" and Trim(Contents)<>"" then ' "#" is a command symbol
			InputData = split(Contents, ";", -1) ' split input line 
			aAppServer(LineCounter, 0) = InputData(0) ' appserver name
			aAppServer(LineCounter, 1) = InputData(1) ' node name
			aAppServer(LineCounter, 2) = InputData(2) ' cell name
			if UCase(InputData(2)) = "WTCELL" then sStage = "T"
			if UCase(InputData(2)) = "WACELL" then sStage = "A"
			if UCase(InputData(2)) = "WPCELL" then sStage = "P"
			aAppServer(LineCounter, 3) = sStage
			
			if bVerbose then printf "ReadConfigFile: Reading " & InputData(0) & ", " & InputData(1) & ", " & InputData(2) & ", " & sStage
			LineCounter = LineCounter + 1
			' aAppServer now carries all appserver information from the line currently read in
		end if
	loop
	set oFilesys = nothing
end sub

sub ControlAppServer(ByVal Mode)
	dim ControlServerCommand
	dim Shell
	dim nCounter
	dim sT 
	
	ControlServerCommand = ""
	nCounter = 1
	set Shell = CreateObject("wscript.shell")
	
	if Mode="stop" and bTerminate then
	  sT = " terminate"
	else
	  sT = ""
	end if 
	
	while aAppServer(nCounter, 0) <> ""
		if Mode="stop" then
			ControlServerCommand = "stopAppServer.cmd" & " " & aAppServer(nCounter, 1) & " " & aAppServer(nCounter, 0) & " " & aAppServer(nCounter, 3) & sT
		else 
			ControlServerCommand = "startAppServer.cmd" & " " & aAppServer(nCounter, 1) & " " & aAppServer(nCounter, 0) & " " & aAppServer(nCounter, 3) 
		end if
		WScript.Echo "executing " & ControlServerCommand & vbCrLf
		
		'Shell.Run ControlServerCommand
		nCounter = nCounter + 1
	wend
	
	set Shell = nothing
end sub

' 
' printf - print on screen and into logfile (if wanted)
'
sub printf(Outputstr)
	WScript.Echo Outputstr
	if bLogging then LogFile.WriteLine Outputstr
end sub


'
' ShowUsage() - get some help
'
Sub ShowUsage()
    Wscript.echo "AppServerControl" & _
    "Stop or start WebSphere application servers given by appropriate information" & vbCRLF & vbCRLF & _
    "cscript AppServerControl.vbs -f[ile] <serverfile> [-v[erbose] [-t[erminate]] [[-]?]" & vbCRLF & _
    "  -f[ile] <serverfile>: treat appservers mentioned in <serverfile>." & vbCRLF & _
	"    <serverfile> is the output of listRunningServers.vbs" & vbCRLF & _
	"  -m[ode] {start|stop}: start or stop the appservers in <serverfile>" & vbCRLF & _
	"  -v[erbose]] (optional): be verbose" & vbCLRF & _
	"  -t[erminate] (optional): don't just stop but terminate all appservers" & vbCrLf & _
	"    will be ignored in 'start' mode" & vbCrLf & _
    "  -? (optional): this help" & vbCRLF & _
    "example: cscript AppServerControl.vbs -m start -v -t -f running-appservers-T.txt"
End Sub