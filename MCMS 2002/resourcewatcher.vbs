'
' resourcewatcher
'
' PURPOSE: Show resources of all MCMS resource galleries exceeding a given threshold
' LAST EDITED: 06-06-20
' AUTHOR(S): Berndtgen, Düsseldorf, 2006
' SYNTAX: cscript resourcewatcher.vbs -c[hannel] <channel> -u[ser] <user>] -p[assword] <password>] [-irg {1|2|3}] [-icb {1|2}] [[-]?]
' VERSION: 1.0
'

dim RootChannelname ' as string ' root channel name to start export
dim User ' as string
dim Password ' as string
dim Threshold ' as integer

dim winShell

' default values
RootChannelname = "/Channels"

' main
Init
WatchResources

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
	        elseif UCase(oArgumente(i)) = "-T" or UCase(oArgumente(i)) = "-THRESHOLD" then
	          	' start row
	          	i = i + 1
	          	Threshold = Int(oArgumente(i))
	      		if err.number<>0 or IsNumeric(Threshold)=False then 
	        		WScript.Echo "ERROR: invalid input data (threshold). Aborting."
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


sub ExamineResources(ByVal StartGallery, ByVal Indentation)
 	Dim pThisGallery, pGallery, pGalleryCollection
  	Dim pResource, pResources
  	Dim pCreator
  	Dim dtmCreated
  	dim kBytes
  	
  	' display the current name of the gallery
    Set pThisGallery = StartGallery
  	rem WScript.Echo pThisGallery.Name

    ' display name of resources in the current gallery
  	Set pResources = pThisGallery.Resources
 	For Each pResource In pResources
 		Set pCreator = pResource.CreatedBy
 		dtmCreated = pResource.CreatedDate
 		kBytes = Int(pResource.Size / 1024)
 		if (kBytes > Threshold) then 
   			WScript.Echo pResource.Path & chr(9) & pCreator.FullName & chr(9) & FormatDateTime(dtmCreated, 2) & chr(9) & kBytes 
   		end if

	 Next
   ' recursively down the ResourceGallery hierarchy
    Set pGalleryCollection = pThisGallery.ResourceGalleries
 	For Each pGallery In pGalleryCollection
     	Call ExamineResources(pGallery, Indentation & "")
 	Next
end sub



'
' WatchResources
'
sub WatchResources
	
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
	
	Dim pGallery, pResult, pSearches
	Set pSearches = MyContextObject.Searches
	Set pResults = pSearches.UserResourceGalleries
	If pResults.Count > 0 Then
		WScript.Echo "Available Resource Galleries:"
	
	 	For Each pGallery In pResults
	    	call ExamineResources(pGallery, "-")
	   	Next
	End If
	
	set winShell = nothing	
	set MyContextObject = nothing
end sub

'
' ShowUsage - display some help
'
Sub ShowUsage()
    Wscript.echo "resourcewatcher" & vbCRLF & _
    "Show resources of all MCMS resource galleries exceeding a given threshold" & vbCRLF & vbCRLF & _
    "cscript resourcewatcher.vbs -c[hannel] <channel> -u[ser] <user>] -p[assword] <password>] [-irg {1|2|3}] [-icb {1|2}] [[-]?]" & vbCRLF & _
    "  -u <user>: user with appropriate channel manager or administrator rights. Syntax: WinNt://<server>/<user>" & vbCRLF & _
    "  -p <password>: password for <user>" & vbCRLF & _
    "  -t <threshold>: max kBytes filter <user>" & vbCRLF & _
    "  -? (optional): this help" & vbCRLF & _
    "example: cscript resourcewatcher.vbs -u WinNt://xyz/user -p p@ssw0rd -t 1000 " 
End Sub
