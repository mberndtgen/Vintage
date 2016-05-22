'
' excel2ldap
'
' PURPOSE: read excel table 
'   (fields: Nachname, Vorname, Land/Region, Firma, Abteilung, Position, 
' 	 E-Mail Adresse, Rufnummer, Fax, CMS Groups Reader, CMS Groups Editor, 
' 	 WSS Teamroom Reader, WSS Teamroom Editor, CMS Channel, WSS Teamroom, Typ, UserID, DeleteFromGroups )
'   and transform it into a flat file of ldap-style user records. Change settings 
'   in the few lines below
' LAST EDITED: 09-04-07
' AUTHOR(S): Berndtgen, 2005
' SYNTAX: cscript excel2ldap.vbs -s[ource] <sourcefile> [-t[arget] <targetfile>] [-r[ow] <startrow>] [-dn <dn>] [-do[main] <domain>] [-u[pdate]] [[-]?]
' VERSION: 1.2
'
' ---------------------------------
' LATEST ADDITIONS AND BUGFIXES
'   - handles now filenames without full path correctly (ActiveX Control needs full path)
'	- more replacement rules in Asciify
' 	- more replacement rules in GetMemberOfEntries
' 	- modifications due to design of Excel table

dim DN, Domain ' as string

'
' change settings section start 
'
DN = "OU=Groups,OU=COMPANY,OU=Customers,DC=user,DC=domain,DC=company"
Domain = "user.domain.company"
'
' change settings section end
'


Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
Const RowStandard = 2

Dim intRow, i, j ' as int
Dim inputfile, outputfile ' as string
Dim oArgumente ' as object
Dim Output, WholeOutput ' as string
Dim StdIn, StdOut

Set StdIn = WScript.StdIn
Set StdOut = WScript.StdOut

intRow = RowStandard
i = 0
j = 0
Output = ""
WholeOutput = ""

on error resume next

set oArgumente = WScript.Arguments

if oArgumente.Count < 1 then
	ShowUsage
	WScript.Quit
else
	do 
  		if UCase(oArgumente(i)) = "-S" or UCase(oArgumente(i)) = "-SOURCE" then
  			' source file in excel format
  	  		i = i + 1
  	  		inputfile = oArgumente(i)
			Dim objFSO
			Dim objFile
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.GetFile(inputfile)
			if (objFSO.GetFileName(objFile) = inputfile) then 
				' was the file given with complete pathname?
				' if so, add the full path to the file name
				inputfile = objFSO.GetAbsolutePathName(objFile)
			end if
			if err.number<>0 then 
				WScript.Echo "ERROR: invalid input data (sourcefile). Aborting."
				WScript.Quit
			end if  	  	
  	  	elseif UCase(oArgumente(i)) = "-T" or UCase(oArgumente(i)) = "-TARGET" then
  	  		' destination file (text format)
  	  		i = i + 1
  	  		outputfile = oArgumente(i)
			if err.number<>0 then 
				WScript.Echo "ERROR: invalid input data (output file). Aborting."
				WScript.Quit
			end if  	  		
  	  	elseif UCase(oArgumente(i)) = "-R" or UCase(oArgumente(i)) = "-ROW" then
  	  		' start row
  	  		i = i + 1
  	  		intRow = oArgumente(i)
			if err.number<>0 then 
				WScript.Echo "ERROR: invalid input data (start row). Aborting."
				WScript.Quit
			end if  	  		
  	  	elseif UCase(oArgumente(i)) = "-DN" then
  	  		' set DN
  	  		i = i + 1
  	  		DN = oArgumente(i)
			if err.number<>0 then 
				WScript.Echo "ERROR: invalid input data (DN). Aborting."
				WScript.Quit
			end if  	  		
  	  	elseif UCase(oArgumente(i)) = "-DO" or UCase(oArgumente(i)) = "-DOMAIN" then 
  	  		' set domain
  	  		i = i + 1
  	  		Domain = oArgumente(i)
			if err.number<>0 then 
				WScript.Echo "ERROR: invalid input data (Domain). Aborting."
				WScript.Quit
			end if  	  		
  	  	elseif UCase(oArgumente(i)) = "-H" or UCase(oArgumente(i)) = "-HELP" or oArgumente = "?" or oArgumente = "-?" then
  	  		ShowUsage
  	  	end if
  	  	i = i + 1		
  loop until i>=oArgumente.Count
end if

' at least there has to be an input file
if inputfile = "" then 
	WScript.Echo "ERROR: missing sourcefile. Aborting."
	ShowUsage
	WScript.Quit
end if
' does the input file exist?
Set fs = CreateObject("Scripting.FileSystemObject")
if not fs.FileExists(inputfile) then
	WScript.Echo "ERROR: input file does not exist. Aborting."
	WScript.Quit
end if
Set objExcel = CreateObject("Excel.Application")
if err.number<>0 then 
	WScript.Echo "ERROR: No Excel present here. Aborting."
	WScript.Quit
end if

' everything okay? then let's roll

Set objWorkbook = objExcel.Workbooks.Open(inputfile)

' write output to output file if output file name is given, else to normal stdout
if outputfile <> "" then 
	set StdOut = fs.OpenTextFile(outputfile, ForWriting, True)
	If Err.number<>0 Then
		WScript.Echo "ERROR: Can't write into " & outputfile & ": " & Err.Message & ". Aborting."
		WScript.Quit
	end if
end if

on error goto 0

' normal mode: prepare records for making new users
Do while objExcel.Cells(intRow, 1).Value <> "" or objExcel.Cells(intRow, 13).Value <> ""
	Output = ""
	wscript.echo "reading line " & cstr(intRow) 
	Output = Output &  "displayName:" & trim(objExcel.Cells(intRow, 1).Value) & "##"
	dim Entry ' as string
	if UCase(trim(objExcel.Cells(intRow, 13))) = "N" then 
		Entry = Left(Asciify(trim(objExcel.Cells(intRow, 1).Value)), 20) ' max 20 chars because this is the maximal length of userPrincipalName
		' new user -> compute defaults
		Output = Output &  "cn:" & Entry & "##" 
		Output = Output &  "userPrincipalName:" & Entry & "@" & trim(Domain) & "##"
		Output = Output &  "sAMAccountName:" & Entry & "##"
	else
		' update/delete mode
		Entry = Asciify(trim(objExcel.Cells(intRow, 12).Value))
		Output = Output &  "cn:" & Entry & "##" 
		Output = Output &  "userPrincipalName:" & Entry & "@" & trim(Domain) & "##"
		Output = Output &  "sAMAccountName:" & Entry & "##"
	end if
	Output = Output &  "sn:" & trim(objExcel.Cells(intRow, 1).Value) & "##"
	Output = Output &  "givenName:" & trim(objExcel.Cells(intRow, 2).Value) & "##"
	Output = Output &  "co:" & trim(objExcel.Cells(intRow, 3).Value) & "##" ' country
	Output = Output &  "company:" & trim(objExcel.Cells(intRow, 4).Value) & "##" ' country
	Output = Output &  "department:" & trim(objExcel.Cells(intRow, 5).Value) & "##" ' department / division
	Output = Output &  "title:" & trim(objExcel.Cells(intRow, 6).Value) & "##" ' department / division
	Output = Output &  "mail:" & trim(objExcel.Cells(intRow, 7).Value) & "##"
	Output = Output &  "telephoneNumber:" & trim(objExcel.Cells(intRow, 8).Value) & "##"
	Output = Output &  "facsimileTelephoneNumber:" & trim(objExcel.Cells(intRow, 9).Value) & "##"
	Output = Output &  "memberOf:" & GetMemberOfEntries(intRow, 10) & "##"
	Output = Output &  "noMemberOf:" & DeleteGroupEntries(intRow, 11) & "##"
	Output = Output &  "action:" & trim(objExcel.Cells(intRow, 13)) & "##"
	Output = Output &  "password: ##"
	Output = Output &  "description:na" 
	Output = sReplace(Output, chr(13), "")
	Output = sReplace(Output, chr(10), "")
	intRow = intRow + 1
	WholeOutput = WholeOutput & Output & vbCrLf
loop

StdOut.Write WholeOutput
objExcel.Quit	
	
Wscript.Echo "done."
WScript.Quit

'
' sReplace - replaces a substring within a string using REs
'
function sReplace(byVal InputString, byVal Pattern, byval ReplaceString) ' as string
	' because there may be some CRs in the output string Output, here's to get rid of them
	dim regEx ' as object
	Set regEx = New RegExp 
	regEx.Pattern = Pattern 
	regEx.IgnoreCase = True 
	regEx.Global = True  
	'wscript.echo("is: " + InputString + ", rs: " + ReplaceString)
	sReplace = regEx.Replace(InputString, ReplaceString) 
	set regEx = nothing
end function

'
' Asciify - make a string ascii-conform 
'
function Asciify(byval InputString) ' as string
	dim TmpStr ' as string
	TmpStr = sReplace(InputString, "á", "a") 	
	TmpStr = sReplace(TmpStr, "â", "a") 
	TmpStr = sReplace(TmpStr, "à", "a") 
	TmpStr = sReplace(TmpStr, "å", "a") 
	TmpStr = sReplace(TmpStr, "Å", "A")
	TmpStr = sReplace(TmpStr, "ä", "ae") 
	TmpStr = sReplace(TmpStr, "Ä", "Ae")
	TmpStr = sReplace(TmpStr, "æ", "ae")
	TmpStr = sReplace(TmpStr, "Æ", "ae")
	TmpStr = sReplace(TmpStr, "é", "e") 
	TmpStr = sReplace(TmpStr, "É", "E") 
	TmpStr = sReplace(TmpStr, "è", "e") 
	TmpStr = sReplace(TmpStr, "ê", "e") 
	TmpStr = sReplace(TmpStr, "í", "i") 
	TmpStr = sReplace(TmpStr, "ì", "i") 
	TmpStr = sReplace(TmpStr, "î", "i") 
	TmpStr = sReplace(TmpStr, "ï", "i") 
	TmpStr = sReplace(TmpStr, "ó", "o") 
	TmpStr = sReplace(TmpStr, "ò", "o") 
	TmpStr = sReplace(TmpStr, "ô", "o") 
	TmpStr = sReplace(TmpStr, "ö", "oe") 
	TmpStr = sReplace(TmpStr, "Ö", "Oe") 
	TmpStr = sReplace(TmpStr, "ü", "ue") 
	TmpStr = sReplace(TmpStr, "Ü", "Ue") 
	TmpStr = sReplace(TmpStr, "ú", "u") 
	TmpStr = sReplace(TmpStr, "ù", "u") 
	TmpStr = sReplace(TmpStr, "û", "u") 
	TmpStr = sReplace(TmpStr, "ß", "ss") 	
	TmpStr = sReplace(TmpStr, "Ç", "C") 	
	TmpStr = sReplace(TmpStr, "ç", "c")
	TmpStr = sReplace(TmpStr, "ñ", "n")
	TmpStr = sReplace(TmpStr, "Ñ", "N")
	TmpStr = sReplace(TmpStr, "-", "_") 	
	TmpStr = sReplace(TmpStr, " ", "_")
	TmpStr = sReplace(TmpStr, "\\", "_")
	TmpStr = sReplace(TmpStr, "\/", "_")
	TmpStr = sReplace(TmpStr, "\[", "_")
	TmpStr = sReplace(TmpStr, "\]", "_")
	TmpStr = sReplace(TmpStr, "\:", "_")
	TmpStr = sReplace(TmpStr, "\;", "_")
	TmpStr = sReplace(TmpStr, "\|", "_")
	TmpStr = sReplace(TmpStr, "\=", "_")
	TmpStr = sReplace(TmpStr, ",", "_")
	TmpStr = sReplace(TmpStr, "\+", "_")
	TmpStr = sReplace(TmpStr, "\*", "_")
	TmpStr = sReplace(TmpStr, "\?", "_")
	TmpStr = sReplace(TmpStr, "<", "_")
	TmpStr = sReplace(TmpStr, ">", "_")
	TmpStr = sReplace(TmpStr, "@", "_")
	TmpStr = sReplace(TmpStr, """", "")
	Asciify = TmpStr
end function

'
' GetMemberOfEntries - compose "memberOf"-Attribute string built from excel input data
'
function GetMemberOfEntries(byVal Row, byval Col)
	dim Result
	dim ReadFromExcel
	dim TmpArray, TmpStr
	
	Result = ""
	ReadFromExcel = trim(objExcel.Cells(Row, Col).Value) ' may be a list of multiple groups
	ReadFromExcel = sReplace(ReadFromExcel, """", "")
	if ReadFromExcel <> "" then
		if InStr(ReadFromExcel, ";") <> 0 then
			' there are multiple group entries in this data field
			TmpArray = Split(ReadFromExcel, ";")
			for each TmpStr in TmpArray 
				if (TmpStr <> "") then Result = Result & "CN=" & trim(TmpStr) & "," & DN & ";"
			next 
		else
			if (trim(ReadFromExcel) <> "") then Result = Result & "CN=" & trim(ReadFromExcel) & "," & DN & ";"
		end if 
	end if
	if Right(Result, 1) = ";" then Result = Left(Result,Len(Result) - 1)
	GetMemberOfEntries = Result
end function

'
' DeleteGroupEntries - compose "noMemberOf"-Attribute string using GetMemberOfEntries
'
function DeleteGroupEntries(byval Row, byval Col)
	DeleteGroupEntries = GetMemberOfEntries(Row, Col)
end function 

'
' ShowUsage - display some help
'
Sub ShowUsage()
    Wscript.echo "excel2ldap" & _
    "Read user account data from an excelsheet" & vbCRLF & vbCRLF & _
    "cscript excel2ldap.vbs -s[ource] <sourcefile> [-t[arget] <targetfile>] [-r[ow] <startrow>] [-dn <dn>] [-do[main] <domain>] [[-]?]" & vbCRLF & _
    "  -s <sourcefile>: input file (MS Excel format)" & vbCRLF & _
    "  -t <targetfile> (optional): output in ldap-like syntax of user account data for further processing with ldap2user.vbs" & vbCRLF & _
    "  -r <startrow> (optional): first row of Excel sheet to read from" & vbCRLF & _
    "  -dn <dn> (optional): organizational DN, e.g. OU=Groups,OU=COMPANY,OU=Customers,DC=user,DC=domain,DC=company" & vbCRLF & _
    "  -do <domain> (optional): Domain name, e.g. user.domain.company" & vbCRLF & _
    "  -? (optional): this help" & vbCRLF & _
    "example: cscript excel2ldap.vbs -s e:\Exceldatei.xls -t e:\Useranlage_JJJJ-MM-TT.txt -3"
End Sub
