'
' cmsimport
'
' PURPOSE: import .sdo file for MCMS
' LAST EDITED: 05-10-13
' AUTHOR(S): Berndtgen, Düsseldorf, 2005
' SYNTAX: cscript cmsimport.vbs -u[ser] <username> -p[assword] <password> [-irg {1|2|3}] [-icb {1|2}] [-roa {1|2|3}] [-ror {1|2}] [[-]?]
' VERSION: 1.0
'

Dim strROPPath
Dim shell
Dim fso
dim User ' as string
dim Password ' as string

dim IncludeRightsGroups 
dim RightsOnAdd 
dim RightsOnReplace 
dim IncludeCreatedBy 

' defaults
' IncludeRightGroups: (1) Do not export rights groups or users (default)
'                     (2) Export rights groups
'                     (3) Export rights groups with members
' IncludeCreatedBy:   (1) Do not include created by information
'                     (2) Include created by information (default)
' RightsOnAdd:        (1) Do not set any rights on newly added containers during an import operation. 
'                     (2) Set rights on newly added containers during an import operation by inheriting 
'						  the rights of the parent container. This value is the default. 
'					  (3) Set rights on newly added containers by using the rights specified in the 
'						  Site Deployment Object (SDO) file. 
' RightsOnReplace:    (1) Keep the existing rights on containers that are replaced during an import operation. 
'						  This value is the default. 
' 					  (2) Set rights on containers that are replaced during an import operation by using the 
'						  rights specified in the Site Deployment Object (SDO) file. 	
IncludeRightsGroups = 2
RightsOnAdd = 3
RightsOnReplace = 2
IncludeCreatedBy = 1


' main
Init
ImportCMSObjects

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
	      	if UCase(oArgumente(i)) = "-S" or UCase(oArgumente(i)) = "-SDO" then
	        	' start channel
	          	i = i + 1
	          	strROPPath = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (sdo file). Aborting."
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
	        elseif UCase(oArgumente(i)) = "-ROA" then
	          	' start row
	          	i = i + 1
	          	RightsOnAdd = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (roa). Aborting."
	        		WScript.Quit
	      		end if 
	        elseif UCase(oArgumente(i)) = "-ROR" then
	          	' start row
	          	i = i + 1
	          	RightsOnReplace  = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (ror). Aborting."
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
' ImportCMSObjects - Routine that does the import.
'
Sub ImportCMSObjects()
	Dim pCmsDeployImport
	On Error Resume Next

	Set fso = WScript.CreateObject("Scripting.FileSystemObject")
	If not(fso.FileExists(strROPPath) ) Then
	    WScript.Echo "ERROR: " & strROPPath & " does not exist. Aborting"
	    WScript.Quit
	End If
	
	Set pCmsDeployImport = WScript.CreateObject("CmsDeployServer.CmsDeployImport.1")
	If (Err.Number <> 0) Then
		WScript.Echo "ERROR getting Object: " & Err.Description
	    Set pCmsDeployImport = Nothing
	    WScript.Quit
	End If
	
	pCmsDeployImport.AuthenticateAsUser User, Password
	If (Err.Number <> 0) Then
		WScript.Echo "ERROR when authenticating: " & Err.Description
	    Set pCmsDeployImport = Nothing
	    WScript.Quit
	End If

	Dim pImportOptions 
	
	Set pImportOptions = pCmsDeployImport.Options
	If (Err.Number <> 0) Then
	    WScript.Echo "ERROR with Import Options: " & Err.Description
	    Set pCmsDeployImport = Nothing
	    Set pImportOptions = Nothing
	    WScript.Quit
	End If
	
	if IncludeRightsGroups < 1 or IncludeRightsGroups > 3 then 
		WScript.Echo "WARNING: Wrong parameter setting for IRG. Setting IRG to default value 1."
		IncludeRightsGroups = 1
	end if
	pImportOptions.IncludeRightsGroups = CInt(IncludeRightsGroups)

	if IncludeCreatedBy < 1 or IncludeCreatedBy > 2 then 
		WScript.Echo "WARNING: Wrong parameter setting for ICB. Setting ICB to default value 1."
		IncludeCreatedBy = 1
	end if
	pImportOptions.IncludeCreatedBy = CInt(IncludeCreatedBy)
	
	if CInt(RightsOnAdd) < 1 or CInt(RightsOnAdd) > 3 then 
		WScript.Echo "WARNING: Wrong parameter setting for ROA. Setting ROA to default value 3."
		RightsOnAdd = 3
	end if
	pImportOptions.RightsOnAdd = CInt(RightsOnAdd)

	if CInt(RightsOnReplace) < 1 or CInt(RightsOnReplace) > 2 then 
		WScript.Echo "WARNING: Wrong parameter setting for ROR. Setting ROR to default value 2."
		RightsOnReplace = 2
	end if
	pImportOptions.RightsOnReplace = CInt(RightsOnReplace)

	Dim strReportUrl 
	
	strReportUrl = pCmsDeployImport.Import(strROPPath) ' do it!
	If (Err.Number <> 0) Then
	    WScript.Echo "ERROR with Importing: " & Err.Description
	    Set pCmsDeployImport = Nothing
	    Set pImportOptions = Nothing
	    WScript.Quit
	End If
	
	' Display the report in the Web browser.
	Call WScript.Echo("Starting browser session to display the export report at " & ReportUrl)
	Set shell = WScript.CreateObject("WScript.Shell")
	call winShell.run("iexplore.exe http://localhost" & strReportUrl, 1, False)
	
	on error goto 0
	
	Set shell = Nothing
	Set pCmsDeployImport = Nothing
	Set pImportOptions = Nothing
End Sub


'
' ShowUsage - display some help
'
Sub ShowUsage()
    Wscript.echo "cmsimport" & vbCRLF & _
    "Import sdo file for MCMS" & vbCRLF & vbCRLF & _
    "cscript cmsimport.vbs -s[do] <sdofile> -u[ser] <username> -p[assword] <password> [-irg {1|2|3}] [-icb {1|2}] [-roa {1|2|3}] [-ror {1|2}] [[-]?]" & vbCRLF & _
    "  -u <user>: user with appropriate channel manager or administrator rights. Syntax: WinNt://<server>/<user>" & vbCRLF & _
    "  -s <sdofile>: .sdo file to import" & vbCRLF & _
    "  -p <password>: password for <user>" & vbCRLF & _
    "  -irg {1|2|3}: (optional) Import Right Groups" & vbCRLF & _
    "       (1 - Do not export rights groups or users (default))" & vbCRLF & _
	"       (2 - Export rights groups)" & vbCRLF & _
	"       (3 - Export rights groups with members)" & vbCRLF & _
    "  -icb {1|2}: (optional) Include Created By" & vbCRLF & _
    "       (1 - Do not include created by information)" & vbCRLF & _
	"       (2 - Include created by information (default))" & vbCRLF & _
	"  -roa {1|2|3}: (optional) Sets rights for containers added during an import operation" & vbCRLF & _
	"       (1 - Do not set any rights on newly added containers during an import operation.)" & vbCRLF & _
	"       (2 - Set rights on newly added containers during an import operation by inheriting" & vbCRLF & _
	"            the rights of the parent container. This value is the default.)" & vbCRLF & _
	"       (3 - Set rights on newly added containers by using the rights specified in the" & vbCRLF & _
	"            Site Deployment Object (SDO) file.)" & vbCRLF & _
	"  -ror {1|2}: (optional) Sets rights for containers replaced during an import operation"  & vbCRLF & _
	"       (1 - Keep the existing rights on containers that are replaced during an import operation. Default)" & vbCRLF & _
	"       (2 - Set rights on containers that are replaced during an import operation by using the" & vbCRLF & _
	"            rights specified in the Site Deployment Object (SDO) file.)" & vbCRLF & _
    "  -? (optional): this help" & vbCRLF & _
    "example: cscript cmsexport.vbs -c /cms/ -u WinNt://xyz/user -p p@ssw0rd -irg 3 -icb 1" & vbCRLF & _
    "ATTENTION: only recently (since last export) changed data will be exported!"
End Sub
