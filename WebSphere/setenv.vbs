const pwdfile = ".\pwd.tmp"

dim Username ' as string
dim bSet ' as boolean
dim bFile ' as boolean

Init
if bSet or bFile then 
	if bSet then SetVar2 "password", Decrypt(ReadIni(".\pwd.ini", "passwords", Username), "huasHIYhkasdho1")
	if bFile then WriteIntoTextFile(Decrypt(ReadIni(".\pwd.ini", "passwords", Username), "huasHIYhkasdho1"))
	Wscript.Echo(Decrypt(ReadIni(".\pwd.ini", "passwords", Username), "huasHIYhkasdho1"))
else 
	UnsetVar("password")
end if
	
sub Init
	dim oArgumente ' as Array
	dim i ' as integer
  
	on error resume next
	
	set oArgumente = WScript.Arguments
  
	if oArgumente.Count < 2 then 
		ShowUsage
		WScript.Quit
	else
		do
			if UCase(oArgumente(i)) = "-U" or UCase(oArgumente(i)) = "-USER" then
				' username
				i = i + 1
				Username = oArgumente(i)
				if err.number<>0 then 
					WScript.Echo "ERROR: invalid input data (username). Aborting."
					WScript.Quit
				end if    
			elseif UCase(oArgumente(i)) ="-S" or UCase(oArgumente(i)) = "-SET" then
				' set variable
				bSet = true
				if err.number<>0 then 
					WScript.Echo "ERROR: invalid input data (set). Aborting."
					WScript.Quit
				end if    
			elseif UCase(oArgumente(i)) ="-D" or UCase(oArgumente(i)) = "-DELETE" then
				' unset variable
				bSet = false
				if err.number<>0 then 
					WScript.Echo "ERROR: invalid input data (delete). Aborting."
					WScript.Quit
				end if    
			elseif UCase(oArgumente(i)) ="-F" or UCase(oArgumente(i)) = "-FILE" then
				' unset variable
				bFile = true
				if err.number<>0 then 
					WScript.Echo "ERROR: invalid input data (file). Aborting."
					WScript.Quit
				end if    
			elseif UCase(oArgumente(i)) = "-H" or UCase(oArgumente(i)) = "-HELP" or oArgumente(i) = "-?" then
				ShowUsage
				WScript.Quit
			end if
			i = i + 1
		loop until i >= oArgumente.Count	
	end if
	on error goto 0
end sub


sub SetVar(ByVal Varname, ByVal Varvalue)
	strVarName = Varname
	strVarValue = Varvalue
	strComputer = "."
	
	set objVarClass = GetObject("winmgmts://" & strComputer & "/root/cimv2:Win32_Environment")
	set objVar = objVarClass.SpawnInstance_
	objVar.Name = strVarName
	objVar.VariableValue = strVarValue
	objVar.UserName = "<SYSTEM>"
	objVar.Put_ 
	Set objVar = nothing
	set objVarClass = nothing
end sub

sub SetVar2(ByVal Varname, ByVal Varvalue)
	set wshShell = CreateObject("WScript.Shell")
	set wshSystemEnv = wshShell.Environment("USER")
	' Display the current value
	' Set the environment variable
	wshSystemEnv(Varname) = Varvalue
	Set wshSystemEnv = Nothing
	Set wshShell = Nothing
end sub


sub UnsetVar(ByVal Varname)
	strVarName = Varname
	strComputer = "."

	set objVarClass = GetObject("winmgmts://" & strComputer & "/root/cimv2:Win32_Environment")
	set objVar = objVarClass.SpawnInstance_
	objVar.Name = strVarName
	objVar.VariableValue = ""
	objVar.UserName = "<SYSTEM>"
	objVar.Put_
	set objVar = nothing
	set objVarClass = nothing
end sub
	
function Decrypt(str,key)
	dim lenKey, KeyPos, LenStr, x, Newstr, DecCharNum

	Newstr = ""
	lenKey = Len(key)
	KeyPos = 1
	LenStr = Len(Str)

	str=StrReverse(str)
	for x = LenStr to 1 step -1
		DecCharNum = asc (mid (str, x, 1)) - asc (mid (key,KeyPos, 1)) + 256
		Newstr = Newstr & chr(DecCharNum mod 256)
		KeyPos = KeyPos+1
		if KeyPos > lenKey then KeyPos = 1
    next
    Newstr=StrReverse(Newstr)
    Decrypt = Newstr
end function

Function ReadIni(myFilePath, mySection, myKey)
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude
	' taken from http://www.robvanderwoude.com/vbstech_files_ini.php

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ReadIni     = ""
    strFilePath = Trim(myFilePath)
    strSection  = Trim(mySection)
    strKey      = Trim(myKey)

    If objFSO.FileExists(strFilePath) Then
        Set objIniFile = objFSO.OpenTextFile(strFilePath, ForReading, False)
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim(objIniFile.ReadLine)

            ' Check if section is found in the current line
            If LCase(strLine) = "[" & LCase(strSection) & "]" Then
                strLine = Trim(objIniFile.ReadLine)

                ' Parse lines until the next section is reached
                Do While Left(strLine, 1) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr(1, strLine, "=", 1)
                    If intEqualPos > 0 Then
                        strLeftString = Trim(Left(strLine, intEqualPos - 1))
                        ' Check if item is found in the current line
                        If LCase(strLeftString) = LCase(strKey) Then
                            ReadIni = Trim(Mid(strLine, intEqualPos + 1))
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim(objIniFile.ReadLine)
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exist. Exiting..."
        Wscript.Quit 1
    End If
End Function

sub WriteIntoTextFile(ByVal password)
	const ForReading = 1
	const ForWriting = 2
	const ForAppending = 8

	set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(pwdfile) then objFSO.DeleteFile(pwdfile)

	set objTextFile = objFSO.OpenTextFile(pwdfile, ForWriting, True)
	objTextFile.WriteLine(password)
	objTextFile.Close
end sub

sub ShowUsage()
	WScript.Echo "SetEnv" & vbCRLF & _
	"bla bla." & vbCRLF & vbCRLF & _
	"cscript SetEnv.vbs -u[sername] <username> {-s[et] | -d[elete]} [-h]" & vbCRLF & _
	"  -u <username>: retrieve/unset ""password"" environment var for user <username>" & vbCRLF & _
	"  -s: set the ""password"" environment variable " & vbCRLF & _
	"  -d: delete the ""password"" environment variable" & vbCRLF & _
	"  -f: write output into a file ~pwd.tmp" & vbCRLF & _
	"  -? (optional): this help"
end sub	