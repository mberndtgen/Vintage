'
' cmshelper
'
' PURPOSE: display channel structure of a MCMS site on a given start channel
' LAST EDITED: 08-08-01
' AUTHOR(S): Berndtgen, Düsseldorf, 2005-2008
' SYNTAX: cscript cmshelper.vbs -c[hannel] <channel> -u[ser] <username> -p[assword] <password> [[-]?]
' VERSION: 1.1
'


dim RootChannel ' as object ' root channel to start search
dim User ' as string
dim Password ' as string

' main
Init
CMSContextHelper

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
	          	RootChannel = oArgumente(i)
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
' CMSContextHelper - main routine
'
sub CMSContextHelper()
	dim MyContextObject
	dim offSet 
	
	on error resume next
	
	set MyContextObject = WScript.CreateObject("ResolutionObjectModel.CmsApplicationContext.1")
	if (Err.Number<>0) then
		WScript.Echo "ERROR: Export Problem Creating CmsApplicationContext Object - " & err.description
		set MyContextObject = nothing
		exit sub
	end if

	' "WinNT://server/mcmsmanage"
	' "mcmsmanage"
	call MyContextObject.AuthenticateAsUser(User, Password, 0, 0, 0, 0)
	if (Err.Number<>0) then
		WScript.Echo "ERROR: Export Problem Authenticating the Admin User - " & Err.Description
		set MyContextObject = nothing
		exit sub
	end if

	on error goto 0 
	
	Dim pResult, pSearches
	Set pSearches = MyContextObject.Searches
	Set pResult = pSearches.GetByPath(RootChannel)
	If not(pResult is nothing) Then
		'set RootChannel = MyContextObject.RootChannel
		set RootChannel = pResult
	else 
		WScript.Echo "ERROR: Root channel " & RootChannel & " not found. Aborting."
		set pSearches = nothing
		set pResult = nothing
		set MyContextObject = nothing
		exit sub
	end if
	
	WScript.Echo "DisplayPath" & chr(9) & "DisplayName" & chr(9) & "Name" & chr(9) & "CreatedDate" & chr(9) & "Author" & chr(9) & "LastModifiedDate" & chr(9) & "LastModifiedBy" & chr(9) & "URL" & chr(9) & "Templatename"
	call ExploreChannels(RootChannel)
	set MyContextObject = nothing
end sub

'
' ExploreChannels - walk down channel structure and print some tab-delimited output
'
sub ExploreChannels(byval Channel)
	dim Channels, pItem
	dim ChildChannel, Posting
	
	on error resume next
	Set Channels = Channel.Channels
	if err.number<>0 then
		WScript.Echo "ERROR: Could not get channel object (maybe wrong channel path) - " & err.description
		WScript.Quit
	end if
	on error goto 0
	
	WScript.Echo "Entering Channel " & Channel.Path & " (DisplayName: " & Channel.DisplayName & ")" 

	For Each pItem In Channels
    	' show links to all subscribed postings and channels
		if pItem.AllChildren.Count > 0 then call ExploreChannels(pItem)
  		  			
  		dim pTemplate
  		dim pCreator
		for each Posting in pItem.Postings
			Set pTemplate = Posting.Template
			Set pCreator =  Posting.CreatedBy
			WScript.Echo Posting.DisplayPath & chr(9) & Posting.DisplayName & chr(9) & Posting.Name & chr(9) & FormatDateTime(Posting.CreatedDate, vbGeneralDate) & chr(9) & pCreator.FullName & chr(9) & FormatDateTime(Posting.LastModifiedDate, vbGeneralDate) & chr(9) &  Posting.URL & chr(9) & pTemplate.Name
		next
    Next 
end sub


'
' ShowUsage - display some help
'
Sub ShowUsage()
    Wscript.echo "cmshelper" & _
    "Show information on MCMS Channels and Postings" & vbCRLF & vbCRLF & _
    "cscript cmshelper.vbs -c[hannel] <channel> -u[ser] <user>] -p[assword] <password>] [[-]?]" & vbCRLF & _
    "  -c <channel>: start output in this channel" & vbCRLF & _
    "  -u <user>: server account name with appropriate rights. Syntax: WinNt://<server>" & vbCRLF & _
    "  -p <password>: password for user <user>" & vbCRLF & _
    "  -? (optional): this help" & vbCRLF & _
    "example: cscript cmshelper.vbs -c /cms/company -u ""WinNt://server/admin"" -p ""p@ssw0rt"" "
End Sub
