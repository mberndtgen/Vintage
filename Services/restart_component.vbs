'
' restart_component
'
' PURPOSE: restarts com+ component (specified by parameter)
' LAST EDITED: 07-10-14
' AUTHOR(S): Berndtgen, Düsseldorf, 2007
' SYNTAX: cscript restart_component.vbs -c[omponent] <componentname> [[-]?]
' VERSION: 1.0
'

Dim catalog 'Enter name of component to restart
Dim ComponentName ' name of COM+ component

' main
Init
RestartComponent


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
	      	if UCase(oArgumente(i)) = "-N" or UCase(oArgumente(i)) = "-NAME" then
	        	' start channel
	          	i = i + 1
	          	ComponentName = oArgumente(i)
	      		if err.number<>0 then 
	        		WScript.Echo "ERROR: invalid input data (name). Aborting."
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
' ShowUsage - display some help
'
Sub ShowUsage()
    Wscript.echo "restart_component" & vbCRLF & _
    "Stop and (Re-)Start a COM+ component" & vbCRLF & vbCRLF & _
    "cscript restart_component.vbs -n[ame] <componentname> [[-]?]" & vbCRLF & _
    "  -n <componentname>: name of COM+ component name" & vbCRLF & _
    "  -? (optional): this help" & vbCRLF & _
    "exmple: cscript restart_component.vbs -n /cms/ -u WinNt://xyz/user -p p@ssw0rd -irg 3 -icb 1" 
End Sub


Sub RestartComponent
	wscript.stdout.writeline("Restarting Com+ Component " + ComponentName)
	
	Dim catalog
	Set catalog = CreateObject("COMAdmin.COMAdminCatalog.1")
	
	If Err.Number <> 0 Then
		wscript.stdout.writeline "Error accessing COM+ Catalog: " & Err.Description
		wscript.Quit(-1)
	End If
	
	wscript.stdout.writeline("Stopping:" & Time)
	catalog.ShutdownApplication(ComponentName)

	wscript.stdout.writeline("Starting:" & Time)
	catalog.StartApplication(ComponentName)
	
	wscript.stdout.writeline "Execution successful"
End Sub