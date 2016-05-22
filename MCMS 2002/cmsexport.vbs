'
' cmshelper
'
' PURPOSE: display channel structure of a MCMS site on a given start channel
' LAST EDITED: 05-10-12
' AUTHOR(S): Berndtgen, Düsseldorf, 2005
' SYNTAX: cscript cmshelper.vbs -c[hannel] <channel> -u[ser] <username> -p[assword] <password> [[-]?]
' VERSION: 1.0
'

dim RootChannelname ' as string ' root channel name to start export
dim User ' as string
dim Password ' as string
dim IncludeCreatedBy ' as string
dim IncludeRightsGroups ' as string
dim Targetfile ' as string

' default values
	' Include Rights Groups and their members
	' IncludeRightGroups: (1) Do not export rights groups or users (default)
	'                     (2) Export rights groups
	'                     (3) Export rights groups with members
	' IncludeCreatedBy:   (1) Do not include created by information
	'                     (2) Include created by information (default)
IncludeRightGroups = 1
IncludeCreatedBy = 2
RootChannelname = "/Channels"

' main
Init
ExportCMSObjects


'
' Init - analyze command line
'
sub Init
	dim oArgumente
	
	on error resume next
	
	set oArgumente = WScript.Arguments
	
	if oArgumente.Count < 1 then
	  	ShowUsage
	 	WScript.Quit
	else
	 	do 
	      	if UCase(oArgumente(i)) = "-C" or UCase(oArgumente(i)) = "-CHANNEL" then
	        	' start channel
	          	i = i + 1
	          	RootChannelname = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (channel). Aborting."
	        		WScript.Quit
	      		end if          
	        elseif UCase(oArgumente(i)) = "-U" or UCase(oArgumente(i)) = "-USER" then
	          	' destination file (text format)
	          	i = i + 1
	          	User = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (User). Aborting."
	        		WScript.Quit
	      		end if          
	        elseif UCase(oArgumente(i)) = "-P" or UCase(oArgumente(i)) = "-PASSWORD" then
	          	' start row
	          	i = i + 1
	          	Password = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (start row). Aborting."
	        		WScript.Quit
	      		end if          
	        elseif UCase(oArgumente(i)) = "-T" or UCase(oArgumente(i)) = "-TARGET" then
	          	' start row
	          	i = i + 1
	          	Targetfile = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (target file). Aborting."
	        		WScript.Quit
	      		end if          
	        elseif UCase(oArgumente(i)) = "-IRG" then
	          	' start row
	          	i = i + 1
	          	IncludeRightsGroups = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (irg). Aborting."
	        		WScript.Quit
	      		end if 
	        elseif UCase(oArgumente(i)) = "-ICB" then
	          	' start row
	          	i = i + 1
	          	IncludeCreatedBy = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (icb). Aborting."
	        		WScript.Quit
	      		end if 
	        elseif UCase(oArgumente(i)) = "-H" or UCase(oArgumente(i)) = "-HELP" or oArgumente = "?" or oArgumente = "-?" then
	          	ShowUsage
        		WScript.Quit
	        end if
	        i = i + 1    
	  	loop until i>=oArgumente.Count
	end if
	
	on error goto 0
	
end sub


'
' ExportCMSObjects
'
sub ExportCMSObjects
	dim MyDeployObject ' as object
	dim MyExportOptions ' as object
	dim PackageFilename, sdoDateQualifier, ReportUrl ' as string
	
	on error resume next
	
	set myDeployObject = WScript.CreateObject("CmsDeployServer.CmsDeployExport.1")
	if (Err.Number<>0) then
		WScript.Echo "ERROR: Export Problem Creating Deployment Object - " & err.description
		set myDeployObject = nothing
		exit sub
	end if

	' User "WinNT://server/mcmsmanage"
	' Password "mcmsmanage"
	call MyDeployObject.AuthenticateAsUser(User, Password)
	if (Err.Number<>0) then
		WScript.Echo "ERROR: Export Problem Authenticating the Admin User - " & Err.Description
		set myDeployObject = nothing
		exit sub
	end if
	
	set MyExportOptions = myDeployObject.Options
	if (Err.Number<>0) then
		WScript.Echo "ERROR: Export Problem Authenticating User - " & Err.Description
		set myDeployObject = nothing
		set MyExportOptions = nothing
		exit sub
	end if
	
	if IncludeRightsGroups < 1 or IncludeRightsGroups > 3 then 
		WScript.Echo "WARNING: Wrong parameter setting for IRG. Setting IRG to default value 1."
		IncludeRightsGroups = 1
	end if
	myExportOptions.IncludeRightsGroups = CInt(IncludeRightsGroups)

	if IncludeCreatedBy < 1 or IncludeCreatedBy > 2 then 
		WScript.Echo "WARNING: Wrong parameter setting for ICB. Setting ICB to default value 1."
		IncludeCreatedBy = 1
	end if
	myExportOptions.IncludeCreatedBy = CInt(IncludeCreatedBy)
	
	' create a date qualifier for creating unique export names
	if Targetfile="" then 
		sdoDateQualifier = day(now()) & month(now()) & year(now()) & hour(now()) & minute(now()) & second(now())
		PackageFilename = sdoDateQualifier & ".sdo"
	else 
		if right(Targetfile, 4) <> ".sdo" then Targetfile = Targetfile & ".sdo"
		PackageFilename = Targetfile
	end if
	
	' start deployment
	ReportUrl = MyDeployObject.Export(PackageFilename, 0, RootChannelname)
	if (Err.Number<>0) then
		WScript.Echo "ERROR: Export Problem creating sdo File - " & Err.description
		set myDeployObject = nothing
		set MyExportOptions = nothing
		exit sub
	end if		
	
	on error goto 0
	
	dim winShell
	set winShell = WScript.CreateObject("WScript.Shell")
	
	' show export report in a web browser
	call WScript.Echo("Export has completed successfully")
	Call WScript.Echo("Starting browser session to display the export report at " & ReportUrl)
	call winShell.run("iexplore.exe http://localhost" & ReportUrl, 1, False)
	
	' clean up
	set myDeployObject = nothing
	set MyExportOptions = nothing
	set winShell = nothing
end sub

'
' ShowUsage - display some help
'
Sub ShowUsage()
    Wscript.echo "cmsexport" & vbCRLF & _
    "Export recently changed MCMS data on a given channel" & vbCRLF & vbCRLF & _
    "cscript cmsexport.vbs -c[hannel] <channel> -u[ser] <user>] -p[assword] <password>] [-irg {1|2|3}] [-icb {1|2}] [[-]?]" & vbCRLF & _
    "  -c <channel>: channel to start export from" & vbCRLF & _
    "  -u <user>: user with appropriate channel manager or administrator rights. Syntax: WinNt://<server>/<user>" & vbCRLF & _
    "  -p <password>: password for <user>" & vbCRLF & _
    "  -irg {1|2|3} (optional): Import Right Groups" & vbCRLF & _
    "       (1 - Do not export rights groups or users (default))" & vbCRLF & _
	"       (2 - Export rights groups)" & vbCRLF & _
	"       (3 - Export rights groups with members)" & vbCRLF & _
    "  -icb {1|2}: (optional): Include Created By" & vbCRLF & _
    "       (1 - Do not include created by information)" & vbCRLF & _
	"       (2 - Include created by information (default))" & vbCRLF & _
	"  -t <targetfile> (optional): Target file name specification" & vbCRLF & _
    "  -? (optional): this help" & vbCRLF & _
    "example: cscript cmsexport.vbs -c /cms/ -u WinNt://xyz/user -p p@ssw0rd -irg 3 -icb 1" & vbCRLF & _
    "ATTENTION: only recently (since last export) changed data will be exported!"
End Sub
