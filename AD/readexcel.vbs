'Declare variables
Dim strArgumentArray()
ReDim strArgumentArray(0)
Dim intRow
Dim ok
Dim filename
Dim j

'Declare constantes
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
Const RowStandard = 3
const CN = "OU=Groups,OU=COMPANY,OU=Customers,DC=user,DC=domain,DC=company"
const FirstGroupCol = 10
const LastGroupCol = 15

'Initialize variables
strArgumentArray(0) = ""
intRow = RowStandard
ok = ""
filename = ""
j = 0

Select Case Wscript.arguments.item(i)
   Case "h"
      Call ShowUsage()
   Case "-h"
      Call ShowUsage()
   Case "--h"
      Call ShowUsage()
   Case "help"
      Call ShowUsage()
   Case "-help"
      Call ShowUsage()
   Case "--help"
      Call ShowUsage()
   Case "?"
      Call ShowUsage()
   Case Else
      
      'Get the command line arguments
      For i = 0 to Wscript.arguments.count - 1
          ReDim Preserve strArgumentArray(i)
          strArgumentArray(i) = Wscript.arguments.item(i)
          j = j + 1
          'Wscript.echo "j:" & j & ", i:" & i & ", " & strArgumentArray(i)
      Next
      
      Set objExcel = CreateObject("Excel.Application")
      Set objWorkbook = objExcel.Workbooks.Open(strArgumentArray(0))
      filename = strArgumentArray(1)
      If j > 2 then intRow = strArgumentArray(2)
      
      Set fs = CreateObject("Scripting.FileSystemObject")
      Set textstream = fs.OpenTextFile(filename, ForWriting, True)
      ok = (Err.number = 0)
      
      If ok Then
         On Error GoTo 0
         
         Do Until objExcel.Cells(intRow,1).Value = ""
            textstream.Write "displayName:" & objExcel.Cells(intRow, 1).Value & "##"
            textstream.Write "cn:" & objExcel.Cells(intRow, 1).Value & "##"
            textstream.Write "sn:" & objExcel.Cells(intRow, 1).Value & "##"
            textstream.Write "userPrincipalName:" & objExcel.Cells(intRow, 1).Value & "@user.domain.company##"
            textstream.Write "sAMAccountName:" & objExcel.Cells(intRow, 1).Value & "##"
            textstream.Write "givenName:" & objExcel.Cells(intRow, 2).Value & "##"
            textstream.Write "co:" & objExcel.Cells(intRow, 3).Value & "##" ' country
            textstream.Write "company:" & objExcel.Cells(intRow, 4).Value & "##" ' country
            textstream.Write "department:" & objExcel.Cells(intRow, 5).Value & "##" ' department / division
            textstream.Write "title:" & objExcel.Cells(intRow, 6).Value & "##" ' department / division
            textstream.Write "mail:" & objExcel.Cells(intRow, 7).Value & "##"
            textstream.Write "telephoneNumber:" & objExcel.Cells(intRow, 8).Value & "##"
            textstream.Write "facsimileTelephoneNumber:" & objExcel.Cells(intRow, 9).Value & "##"
            textstream.Write "memberOf:" & GetMemberOfEntries(intRow) & "##"
            textstream.Write "action:" & objExcel.Cells(intRow, LastGroupCol + 1) & "##"
            textstream.Write "password: ##"
            textstream.Write "description:na" & vbCrLf
            
            intRow = intRow + 1
         Loop
      Else
         MsgBox "Fehler: " & Err.Description
      End If
      
      objExcel.Quit
end Select
Wscript.Echo "done."
WScript.Quit


function GetMemberOfEntries(byVal Row)
	dim i
	dim Result
	dim ReadFromExcel
	
	for i = FirstGroupCol to LastGroupCol
		ReadFromExcel = trim(objExcel.Cells(Row, i).Value)
		if ReadFromExcel <> "" then Result = Result & "CN=" & ReadFromExcel & "," & CN & ";"
	next
	if Right(Result, 1) = ";" then Result = Left(Result,Len(Result) - 1)
	Wscript.Echo "memberOf: " & Result
	GetMemberOfEntries = Result
end function

'********************************************************************
'*
'* Sub ShowUsage()
'* Purpose: Shows the correct usage to the user.
'* Input:   None
'* Output:  Help messages are displayed on screen.
'*
'********************************************************************

Private Sub ShowUsage()
    Wscript.echo " " & _
    "Read user account data out of an excelsheet." & vbCRLF & vbCRLF & _
    "READEXCEL.VBS <SOURCEFILE> <TARGETFILE> [<COLUMNNUMBER>]" & vbCRLF & _
    "           SOURCEFILE    An file in Excel format." & vbCRLF & _
    "           TARGETFILE    The output of the user account data for the AD script." & vbCRLF & vbCRLF & _
    "(optional) COLUMNNUMBER  In wich row start to get the data." & vbCRLF & _
    "EXAMPLE:" & vbCRLF & _
    "READEXCEL.VBS e:\Exceldatei.xls e:\Useranlage_JJJJ-MM-TT.txt 3"

End Sub