'
' listRunningServers
'
' PURPOSE: enlist application servers with status "running" in a given websphere cell
' LAST EDITED: 11-08-10
' AUTHOR(S): Berndtgen, Düsseldorf, 2010
' SYNTAX: cscript listRunningServers.vbs -s[erver] <server> -p [dmgrportnumber] -u " -user <dmgruser> -password <dmgruserpwd>" -v -c [conffile] [[-]?] 
' VERSION: 2
'

' servername: wpnodea / wpnodeb

dim sServer ' server name
dim ConfigDict ' Dictionary object
dim HostsDict ' hosts table content
dim profileDict ' Dictionary object containing profile names and paths
dim Configfilename ' name of configuration file
dim LogFilename ' name of logfile
dim LogFile ' as file object
dim bVerbose ' as boolean
dim bLogging ' as boolean
dim sUserInfo ' as string
dim sDMGRPort ' as string
dim sEnvironment ' as string
dim aNodes(20) ' array containing WAS node names
dim aCells(20) ' array containing WAS cell names
dim DictAppSrv ' Dictionary containing WAS appserver names
dim nNodesCounter, nCellsCounter, nAppSrvCounter ' counter var for aNodes
dim oShell
dim sServersIPAddress, sServername, sServerLongname


' main
Init
ReadConfigFile(Configfilename) ' fill ConfigDict with values
GetProfiles(ConfigDict.Item("PROFILEREGISTRY"))
GetNodenames ConfigDict.Item("WASADMIN"), ConfigDict.Item("ACGETNODENAMES")
GetRunningAppServers ConfigDict.Item("WASADMIN"), ConfigDict.Item("NODEINFOSCRIPT")
PutResultsInFile
CleanUp

'
' Init - analyze command line
'
sub Init
	Set ConfigDict = CreateObject("Scripting.Dictionary") ' configuration container
	Set HostsDict = CreateObject("Scripting.Dictionary") ' hosts file content
	Set profileDict = CreateObject("Scripting.Dictionary") ' profile configuration container
	Set DictAppSrv = CreateObject("Scripting.Dictionary") ' profile configuration container
	Set nodesDict = CreateObject("Scripting.Dictionary") ' nodes configuration container
	Set oShell = WScript.CreateObject("wscript.shell") ' give us a cmd shell
	bVerbose = false
	bLogging = false
	sUserInfo = ""
	sDMGRPort = ""
	sEnvironment = ""
	LogFilename = "logfile.txt"
	nNodesCounter = 1
	nAppSrvCounter = 1
	nCellsCounter = 1
	
	dim oArgumente
	
	on error resume next
	
	set oArgumente = WScript.Arguments
	
	if oArgumente.Count < 2 then
	  	ShowUsage
	 	WScript.Quit
	else
	 	do 
	      	if UCase(oArgumente(i)) = "-S" or UCase(oArgumente(i)) = "-SERVER" then
	        	' server name
	          	i = i + 1
	          	sServername = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (server). Aborting."
	        		WScript.Quit
	      		end if    
	      	elseif UCase(oArgumente(i)) = "-E" or UCase(oArgumente(i)) = "-ENVIRONMENT" then
	        	' test, approval or productive stage
	          	i = i + 1
	          	sEnvironment = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (environment). Aborting."
	        		WScript.Quit
	      		end if    
			elseif UCase(oArgumente(i)) = "-C" or UCase(oArgumente(i)) = "-CONFIGURATION" or UCase(oArgumente(i)) = "-CONF" then
	  	  		' config file 
	  	  		i = i + 1
	  	  		Configfilename = oArgumente(i)
				if err.number<>0 then 
					WScript.Echo "ERROR: invalid input data (config file). Aborting."
					WScript.Quit
				end if  	  	
			elseif UCase(oArgumente(i)) = "-V" or UCase(oArgumente(i)) = "-VERBOSE" then
	  	  		' verbose mode
	  	  		bVerbose = true
				WScript.Echo "Using verbose mode."
	  	  	elseif UCase(oArgumente(i)) = "-L" or UCase(oArgumente(i)) = "-LOGFILE" then
	  	  		' logfile 
	  	  		i = i + 1
	  	  		LogFilename = oArgumente(i)
				if err.number<>0 then 
					WScript.Echo "ERROR: invalid input data (logfile). Aborting."
					WScript.Quit
				else 
					bLogging = true
					on error resume next
					Set LogFile = filesys.CreateTextFile(LogFilename, True) 
					if err.number<>0 then
						WScript.Echo "ERROR: can't write " & LogFilename & ". " & Err.Description & ". Aborting."
						WSCript.Quit
					end if
					on error goto 0
				end if  
			elseif UCase(oArgumente(i)) = "-P" or UCase(oArgumente(i)) = "-PORT" then
				' provide port number for DMGR
				i = i + 1
				on error resume next
				sDMGRPort = CInt(oArgumente(i))
				printf "DMGR Port: " & sDMGRPort & vbCrLf
				if isNumeric(sDMGRPort)=false or sDMGRPort<0 or sDMGRPort>65535 then
					WScript.Echo "ERROR: invalid input data (port). Aborting."
					WScript.Echo Err.Description
					WScript.Quit
				end if
				on error goto 0
			elseif UCase(oArgumente(i)) = "-U" or UCase(oArgumente(i)) = "-USERINFO" then
				' provide user info for enabling wasadmin to login
				i = i + 1
				sUserInfo = oArgumente(i)
				printf "Userinfo: " & sUserInfo & vbCrLf
				if err.number<>0 then
					WScript.Echo "ERROR: invalid input data (userinfo). Aborting."
					WScript.Quit
				end if
	        elseif UCase(oArgumente(i)) = "-H" or UCase(oArgumente(i)) = "-HELP" or oArgumente(i) = "?" or oArgumente(i) = "-?" then
	          	ShowUsage
	          	WScript.Quit
	        end if
	        i = i + 1    
	  	loop until i>=oArgumente.Count
		
		if sServername="" or Configfilename="" or sUserInfo="" then 
			ShowUsage
			WScript.Quit
		end if
	end if

	GetIpAndName sServername, sServersIPAddress, sServerLongname
	printf "Now going to shutdown all WebSphere processes on " & sServerLongname & " (" & sServersIPAddress & "), using Account <userinfo> with Port " & sDMGRPort & " in " & sEnvironment & " environment." & vbCRLF
	on error goto 0
	
end sub

'
' GetProfiles - Read profile data (name and path) from input file
'
sub GetProfiles(ByVal InputFilename)
	' read profileregistry.xml
	dim oFilesys, oFile ' FilesystemObject, Readfile
	dim Contents ' content of oFile
	dim InputData, KeyValArray ' as array
	dim IDElement ' string
	dim Lines() ' array of lines per InputFilename
	dim i ' Array counter
	
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
	
	do while not oFile.AtEndOfStream 
		Contents = oFile.ReadLine
		
		if bVerbose then printf "Reading " & Contents

		' extract profile names
		
		dim RetStr ' match string
		dim re ' regular expression for filtering out links
		Set re = New RegExp
		RetStr = ""
		re.Pattern = "name=""([^""]*)"" path=""([^""]*)"""
		re.Global = True
		re.IgnoreCase = True
		Set Matches = re.Execute(Contents)
		if (Matches.Count > 0) then 
			set Match = Matches(0)
			profileDict.Add Match.SubMatches(0), Match.SubMatches(1)
			' profileDict now carries profile name and profile path from the line currently read in
		end if
	loop
	set oFilesys = nothing
	
	dim aKeys
	
	aKeys = profileDict.Keys   ' Get the items.
	if bVerbose then 
		for i = 0 To profileDict.Count -1 ' Iterate the array.
			printf i & ", " & aKeys(i) & ", " & profileDict(aKeys(i))
		next
	end if
end sub


'
' CleanUp: tidy up objects
'
sub CleanUp
	if bVerbose then WScript.Echo "Tidying up."
	if bLogging then 
		if IsObject(Logfile) then 
			LogFile.Close   	
			set LogFile = nothing
		end if
	end if
	set nodesDict = nothing
	set profileDict = nothing
	set ConfigDict = nothing
	set oShell = nothing
end sub


'
' ReadConfigFile - read input file into array of ConfigDictionaries
'
sub ReadConfigFile(ByVal InputFilename)
	dim oFilesys, oFile ' FilesystemObject, Readfile
	dim Contents ' content of oFile
	dim InputData, KeyValArray ' as array
	dim IDElement ' string
	dim Lines() ' array of lines per InputFilename
	dim i ' Array counter
	
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
	
	do while not oFile.AtEndOfStream 
		Contents = oFile.ReadLine
		
		if bVerbose then printf "Reading " & Contents
		
		if left(Contents, 1) <> "#" and Trim(Contents)<>"" then ' "#" is a command symbol
			InputData = split(Contents, "=", 2) ' split input line 
			ConfigDict.Add InputData(0), InputData(1)
			' ConfigDict now carries all user information from the line currently read in
		end if
	loop
	set oFilesys = nothing
end sub


'
' ReadAndUnderstandHostsFile
'
sub ReadAndUnderstandHostsFile
	dim oFilesys, oFile ' FilesystemObject, Readfile
	dim Contents ' content of oFile
	dim InputData, KeyValArray ' as array
	dim IDElement ' string
	dim Lines() ' array of lines per InputFilename
	dim i ' Array counter
	Const HostsFile = "C:\WINDOWS\system32\Drivers\etc\hosts"
	
	i = 0
	set oFilesys = CreateObject("Scripting.FileSystemObject")
	if not oFilesys.FileExists(HostsFile) then 
		' no file? abort
		printf "ERROR: Hosts File does not exist. Aborting."
		WScript.Quit
	end if

	on error resume next

	set oFile = oFilesys.OpenTextFile(HostsFile, 1, false)
	if err.number<>0 then
		printf "ERROR: Can't read input file: " & Err.Message & ". Aborting."
		WScript.Quit
	end if

	on error goto 0
	
	do while not oFile.AtEndOfStream 
		Contents = oFile.ReadLine
		
		if bVerbose then printf "Reading " & Contents
		
		if left(Contents, 1) <> "#" and Trim(Contents)<>"" then ' "#" is a command symbol
			InputData = split(Contents, " ", 2) ' split input line 
			if HostsDict.Exists(InputData(0)) then 
				' this ip address is already known
				HostsDict(InputData(0)) = HostsDict(InputData(0)) & ";" & InputData(1)
			else
				HostsDict.Add InputData(0), InputData(1)
			end if
			' HostsDict now carries all IP addresses and their appropriate names
		end if
	loop
	set oFilesys = nothing
end sub

'
' GetNodenames - read all necessary profile data from USERPATH directory filesys
'
sub GetNodenames(ByVal sWasAdmin, ByVal sACGetNodenames)
	dim Matches, Match
	dim sReadLine
	
	if bVerbose then printf "GetNodenames: executing " & sWasAdmin & " <userinfo> -port " & sDMGRPort & " -c " & sACGetNodenames & vbCRLF
	
	Set objWshScriptExec = oShell.Exec(sWasAdmin & " " & sUserInfo & " -port " & sDMGRPort & " -c " & sACGetNodenames)

	Set objStdOut = objWshScriptExec.StdOut
 
	While Not objStdOut.AtEndOfStream
		sReadLine = objStdOut.ReadLine

		' get nodename
		Match = fGetMatches(sReadLine, "node=([^,]*),")
		if Match<>"" then 
			if bVerbose then 
				printf "GetNodenames: Found node " & Match & " in " & sReadLine & vbCRLF 
				printf "GetNodenames: put " & Match & " at aNodes(" & nNodesCounter & ")" & vbCRLF
			end if
			aNodes(nNodesCounter) = Match
			nNodesCounter = nNodesCounter + 1
		end if

		' get cellname
		Match = fGetMatches(sReadLine, "cell=([^,]*),")
		if Match<>"" then 
			if bVerbose then 
				printf "GetNodenames: Found cell " & Match & " in " & sReadLine & vbCRLF 
				printf "GetNodenames: put " & Match & " at aCells(" & nCellsCounter & ")" & vbCRLF
			end if
			aCells(nCellsCounter) = Match
			nCellsCounter = nCellsCounter + 1
		end if
		' each node name in aNodes now has its matching cell name in aCells
		Set objWshScriptExec = nothing
	Wend 

end sub


sub GetRunningAppServers(ByVal sWasAdmin, ByVal sNodeInfoScript)
	dim iReturn 
	dim Matches, Match 
	dim sReadLine
	
	for i = 1 To UBound(aNodes, 1) ' Iterate the aNodes array.
		if aNodes(i)<>"" then
			Set objWshScriptExec = oShell.Exec(sWasAdmin & " -port " & sDMGRPort & " " & sUserInfo & " -f " & sNodeInfoScript & " " & aNodes(i))
			Set objStdOut = objWshScriptExec.StdOut
			While Not objStdOut.AtEndOfStream
				sReadLine = objStdOut.ReadLine
				Match = fGetMatches(sReadLine, "name=([^,]*),")
				if Match<>"" and Match<>"nodeagent" then 
					if bVerbose then printf "GetRunningAppServers: Found appsrv " & Match & " in " & sReadLine & vbCRLF 
					if bVerbose then printf "GetRunningAppServers: put " & aNodes(i) & ";" & aCells(i) & " into DictAppSrv(" & Match & ")" & vbCRLF
					DictAppSrv(Match) = aNodes(i) & ";" & aCells(i)
				end if
			Wend
			Set objWshScriptExec = nothing
		end if
	next
end sub


'
' GetNsLookupAnswer - calls nslookup to obtain server's ip address
'
Function GetNsLookupAnswer(ByVal strAdr) 
   Dim objShell, objExec 
   Dim strPingResults 
   
   Set objShell = WScript.CreateObject("WScript.Shell")
   Set objExec = objShell.Exec("nslookup " & strAdr)
   GetNsLookupAnswer = objExec.StdOut.ReadAll
   Set objExec = Nothing
   Set objShell = Nothing
End Function


'
' GetIpAndName - parses nslookup's output
'
Function GetIpAndName(ByVal strIPOrName, ByRef strIp, ByRef strName) 
   Dim objRegExp, objMC 
   Set objRegExp = New RegExp
   
   objRegExp.Pattern = "^Name: +(.*?)$\s+^Address: +(.*?)$"
   objRegExp.MultiLine = True
   
   Set objMC = objRegExp.Execute(GetNsLookupAnswer(strIPOrName))
   If objMC.Count Then
      GetIpAndName = True
      strName = objMC(0).SubMatches(0)
      strIp = objMC(0).SubMatches(1)
   End If
   if bVerbose then WScript.Echo "Using Server: " & strName & "with IP: " & strIP
   Set objRegExp = Nothing
   Set objMC = Nothing
End Function


'
' GetDedicatedIps
'
Function GetDedicatedIps(ByVal strComputer, ByRef strFEIp, ByRef strBEIp) 
	dim objWMIService, colAdapters
	
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=True")
	For Each objAdapter In colAdapters
		' to obtain every ip address on each adapter:
		'for i=0 to UBound(objAdapter.IPAddress)
		'	Wscript.Echo i & "; " & objAdapter.IPAddress(i) & " : " & strComputer
		'next
		wscript.echo objAdapter.IPAddress(0) '## auf FE und BE verteilen
	Next
	Set objWMIService = nothing
	Set colAdapters = nothing
End Function

'
' fGetMatches
'
function fGetMatches(byVal sSource, byVal sPattern)
	dim RetStr ' match string
	dim Matches, Match
	dim re ' regular expression for filtering out links
	
	if (sSource="") then 
		fGetMatches = ""
	else
		Set re = New RegExp
		RetStr = ""
		re.Pattern = sPattern
		re.Global = True
		re.IgnoreCase = True

		if bVerbose then printf "fGetMatches: sSource = " & sSource & ", sPattern = " & sPattern & vbCRLF
		Set Matches = re.Execute(sSource)
		for each Match in Matches
			' if bVerbose then printf "fGetMatches: Match = " & Match
			RetStr = RetStr & Match.SubMatches(0) & " "
		next
		if (bVerbose and RetStr<>"") then printf "fGetMatches: found " & RetStr
		fGetMatches = Trim(RetStr)
	end if
end function


'
' PutResultsInFile - write list of nodes and their running appservers into a results file
'
sub PutResultsInFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	Dim fso, f, ts, i, keys
	Dim Filename
	
	'Filename = "running-appservers.txt"
	Filename = "running-appservers-" & sEnvironment & ".txt"
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(Filename) Then
		' delete old version of this file
		fso.DeleteFile(Filename) 
    end if

	fso.CreateTextFile Filename   ' Create a file.
	Set f = fso.GetFile(Filename)
	Set ts = f.OpenAsTextStream(ForWriting, TristateUseDefault)
	ts.Write "# running appservers" & vbCRLF
	
	keys = DictAppSrv.Keys
	for i=0 to DictAppSrv.Count-1
		ts.Write keys(i) & ";" & DictAppSrv(Keys(i)) & vbCRLF
	next
	ts.Close
end sub

' 
' printf - print on screen and into logfile (if wanted)
'
sub printf(Outputstr)
	WScript.Echo Outputstr
	if bLogging then LogFile.WriteLine Outputstr
end sub


'
' ShowUsage - display some help
'
Sub ShowUsage()
    Wscript.echo "listRunningServers" & _
    "List running WebSphere application servers of a given configuration" & vbCRLF & vbCRLF & _
    "cscript listRunningServers.vbs -s[erver] -e {T|A|P} -p <port> -u ""-user <username> -password <password>"" [-l[ogfile] <logfilename>] [-v[erbose]] [[-]?]" & vbCRLF & _
    "  -s[erver] <server>: server name" & vbCRLF & _
	"  -p[ort] <portnumber>: DMGR port number >0 and <64K)" & vbCLRF & _
	"  -u[serinfo] "" -port <dmgrport> -user <username> -password <password> """ & vbCrLf & _
	"  -e[nvironment] {T|A|P}: set option to Test, Approval or Productive environment" & vbCrLf & _
    "  -c[conf[iguration]] <configfile>: configuration file" & vbCRLF & _	
    "  -l[ogfile] <logfile> (optional): write various output into <logfile>" & vbCRLF & _
    "  -? (optional): this help" & vbCRLF & _
    "example: listRunningServers.vbs -s server -e T -p 7101 -u ""-user user_id -password ###"" -v -c .\wasnodes.conf"
End Sub
