'
' searchresource
'
' PURPOSE: searches MCMS objects by a given GUID, URL, or Channel Path
' LAST EDITED: 06-06-20
' AUTHOR(S): Berndtgen, Düsseldorf, 2006
' SYNTAX: cscript searchresource.vbs -c[hannel] <channel> -u[ser] <user> -p[assword] <password> {-g[uid] <guid> | -c[hanpath] <path> | -a[ddress] <url>}
' VERSION: 1.0
'

dim RootChannelname ' as string ' root channel name to start export
dim User ' as string
dim Password ' as string
dim Guid ' as string
dim ChanPath ' as string
dim Url ' as string

dim winShell
dim SearchFor ' as integer: 0: guid, 1: path, 2: url
SearchFor = -1

' default values
RootChannelname = "/Channels"

' main
Init
SearchResource

'
' Init - analyze command line
'
sub Init
	dim oArgumente
	
	on error resume next
	
	set oArgumente = WScript.Arguments
	
	if oArgumente.Count < 3 then
	  	ShowUsage
	 	WScript.Quit
	else
	 	do 
	        if UCase(oArgumente(i)) = "-U" or UCase(oArgumente(i)) = "-USER" then
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
	        elseif UCase(oArgumente(i)) = "-C" or UCase(oArgumente(i)) = "-CHANPATH" then
	          	' start row
	          	i = i + 1
	          	ChanPath = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (channel path). Aborting."
	        		WScript.Quit
	      		end if 
	      		if SearchFor <> -1 then 
	        		WScript.Echo "ERROR: only one search type allowed. Aborting."
	        		WScript.Quit
	        	else
	        		SearchFor = 1
	      		end if 
	        elseif UCase(oArgumente(i)) = "-A" or UCase(oArgumente(i)) = "-ADDRESS" then
	          	' start row
	          	i = i + 1
	          	Url = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (URL). Aborting."
	        		WScript.Quit
	      		end if          
	      		if SearchFor <> -1 then 
	        		WScript.Echo "ERROR: only one search type allowed. Aborting."
	        		WScript.Quit
	        	else
	        		SearchFor = 2
	      		end if 
	        elseif UCase(oArgumente(i)) = "-G" or UCase(oArgumente(i)) = "-GUID" then
	          	' start row
	          	i = i + 1
	          	Guid = "{" & oArgumente(i) & "}"
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (guid). Aborting."
	        		WScript.Quit
	      		end if          
	      		if SearchFor <> -1 then 
	        		WScript.Echo "ERROR: only one search type allowed. Aborting."
	        		WScript.Quit
	        	else
	        		SearchFor = 0
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
' WatchResources
'
sub SearchResource
	
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

	set winShell = WScript.CreateObject("WScript.Shell")	
	
	WScript.Echo "Searching..." 
	Set pSearches = MyContextObject.Searches
	
	select case SearchFor
		case 0
			' GUID search
			Set pResults = pSearches.GetByGUID(Guid)
		case 1
			' path search
			Set pResults = pSearches.GetByPath(ChanPath)
		case 2
			' url search
			Set pResults = pSearches.GetByURL(Url)
	end select
	
	dim kBytes 
	If pResults is Nothing Then
		WScript.Echo "Nothing found."
	else
		WScript.Echo "Name: " & chr(9) & chr(9) & pResults.Name 
		WScript.Echo "Creation Date: " & chr(9) & FormatDateTime(pResults.CreatedDate, 2) 
		if (pResults.Type) = "Resource" then 
			kBytes = Int(pResults.Size / 1024)
			WScript.Echo "KB: "  & chr(9) & chr(9) & kBytes 
		end if
		WScript.Echo "Displaypath: "  & chr(9) & pResults.DisplayPath
		WScript.Echo "Path: " & chr(9) & chr(9) & pResults.Path 
		WScript.Echo "Type: " & chr(9) & chr(9) & pResults.Type
		WScript.Echo "GUID: " & chr(9) & chr(9) & pResults.GUID
	End If
	set winShell = nothing	
	set MyContextObject = nothing
end sub

'
' ShowUsage - display some help
'
Sub ShowUsage()
    Wscript.echo "searchresource" & vbCRLF & _
    "Search for a resource / posting / channel given on a GUID, Url, or Channel Path" & vbCRLF & vbCRLF & _
    "cscript searchresource.vbs -c[hannel] <channel> -u[ser] <user> -p[assword] <password> {-g[uid] <guid> | -c[hanpath] <path> | -a[ddress] <url>}" & vbCRLF & _
    "  -u <user>: user with appropriate channel manager or administrator rights. Syntax: WinNt://<server>/<user>" & vbCRLF & _
    "  -p <password>: password for <user>" & vbCRLF & _
    "  -g <guid>: GUID of resource to search for. No curly brackets!" & vbCRLF & _
    "  -c <pathname>: full qualified path name for a channel, posting, template, TemplateGallery, Resource, or ResourceGallery object" & vbCrLf & _
    "  -a <url>: URL that uniquely identifies a posting or channel within the channel hierarchy. My be hierarchy-based or GUID-based." & vbCrLf & _
    "  -? (optional): this help" & vbCRLF & _
    "examples: cscript searchresource.vbs -u WinNt://xyz/user -p p@ssw0rd -g 44DF6E97-297C-4D59-A1B1-6D56D3728F77 " & vbCrLf & _
    "          cscript searchresource.vbs -u WinNt://xyz/user -p p@ssw0rd -a ""/Channel/Channel"" (omit /Channels!)" & vbCrLf & _
    "          cscript searchresource.vbs -u WinNt://xyz/user -p p@ssw0rd -c ""/Channels/Channel/Channel2"""
End Sub
