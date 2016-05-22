'
' TidyUp
'
' PURPOSE: erase files (of given type(s)) (recursively) in a given folder that are older than the given number of days
' LAST EDITED: 07-04-11
' AUTHOR(S): M. Berndtgen, Düsseldorf, 2007
' SYNTAX: cscript -f[older] <folder>
'                 -d[ays] <days>
'                [-r[ecursive]]
'                [-[f]i[lter] "<regex-type1>||<regex-type2>||...||<regex-typen>]"
'                [-m <targetfolder>]
'                [-[f]a[ke]]
' VERSION: 1.2
'

dim FolderName ' as Object
Dim TargetFolderName ' as Object
dim NumberOfDays ' as integer
dim RecursiveFlag ' as boolean
dim Filterflag ' as boolean
dim FakeAction ' as boolean
Dim MoveFiles ' as boolean
dim objFSO ' as Object
dim FileTypes ' as Array

' main
Init
TidyUp

'
' Init - analyze command line
'
sub Init
  dim oArgumente ' as Array
  dim i ' as integer
  
  RecursiveFlag = false
  Filterflag = false
  FakeAction = False
  MoveFiles = false
  set oArgumente = WScript.Arguments
  
  if oArgumente.Count < 2 then 
     ShowUsage
     WScript.Quit
  else
     do
        if UCase(oArgumente(i)) = "-F" or UCase(oArgumente(i)) = "-FOLDER" then
           ' folder to operate in
           i = i + 1
           FolderName = oArgumente(i)
        elseif UCase(oArgumente(i)) = "-D" or UCase(oArgumente(i)) = "-DAYS" then
           ' number of days
           i = i + 1
           NumberOfDays = CInt(oArgumente(i))
           if NumberOfDays <= 0 then
              WScript.Echo "ERROR: days must be > 0"
              WScript.Quit
           end if
        elseif UCase(oArgumente(i)) = "-R" or UCase(oArgumente(i)) = "-RECURSIVE" then
           ' number of days
           RecursiveFlag = true
        elseif UCase(oArgumente(i)) = "-A" or UCase(oArgumente(i)) = "-FAKE" then
           ' just pretend deleting, don't do anything that would harm our precious data
           FakeAction = True
        Elseif UCase(oArgumente(i)) = "-M" Or UCase(oArgumente(i)) = "-MOVE" Then
           ' move files to destination folder
           i = i + 1
           TargetFolderName = oArgumente(i)
           MoveFiles = true
        elseif UCase(oArgumente(i)) = "-I" or UCase(oArgumente(i)) = "-FILTER" then
           ' set file filter
           i = i + 1
           Filterflag = true
           FileTypes = Split(Replace(oArgumente(i), """", ""), "||", -1, 1)
        elseif UCase(oArgumente(i)) = "-H" or UCase(oArgumente(i)) = "-HELP" or oArgumente(i) = "-?" then
           ShowUsage
           WScript.Quit
        end if
        i = i + 1
     loop until i >= oArgumente.Count	
  end if
end sub

'
' TidyUp - main routine
'
sub TidyUp()
   dim objFolder ' as Object
   dim colFiles ' as Collection
   
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   
   If objFSO.FolderExists(FolderName) Then
      Set objFolder = objFSO.GetFolder(FolderName)
   Else
      Wscript.Echo "ERROR: Folder does not exist."
      WScript.Quit
   End If
   
   Set colFiles = objFolder.Files
   Call FileAction(objFolder, colFiles)
      
   if RecursiveFlag then TraverseFolders(objFolder)
   
end sub

'
' TraverseFolders - traverse Folders recursively
'
sub TraverseFolders(Folderobj)
   dim objFolder ' as object
   dim Subfolder ' as object
   dim colFiles ' as Collection
   
   for each Subfolder in Folderobj.SubFolders
      set objFolder = objFSO.GetFolder(Subfolder.Path)
      set colFiles = objFolder.Files

      Call FileAction(objFolder, colFiles)
      Call TraverseFolders(objFolder)
   next
end Sub


'
' Do File Action
'
Sub FileAction(objFolder, colFiles)
   dim filetype ' as string
   dim objFile ' as object

   for each objFile in colFiles
      'WScript.Echo objFile.Name & "/" & objFile.DateCreated & "/" & Abs(DateDiff("d", Now, objFile.DateCreated))
      if Abs(DateDiff("d", objFile.DateCreated, Now)) >= NumberOfDays then
         if Filterflag then
            for each filetype in FileTypes
'               if filetype = Right(objFile.Name, Len(filetype)) Then
               if MatchFile(filetype, objFile.Name) Then
                  If Not MoveFiles Then ' move or delete
                     call DeleteFile(objFolder.Path, objFile.Name)
                  Else
                     Call MoveFile(objFolder.Path, objFile.Name, TargetFolderName)
                  End If
                  exit for
               end if
            next
         else 
            If Not MoveFiles Then ' move or delete
               call DeleteFile(objFolder.Path, objFile.Name)
            Else
               Call MoveFile(objFolder.Path, objFile.Name, TargetFolderName)
            End If
         end if
      end if
   next
End Sub

'
' MatchFile - Regex search for files
'
Function MatchFile(Pattern, Filename)
   Dim regEx, Match, Matches
   Set regEx = New RegExp

   regEx.Pattern = Pattern
   regEx.IgnoreCase = True ' case sensitivity not req'd for file name search
   regEx.Global = True

   MatchFile = regEx.Test(Filename)
End Function 


'
' DeleteFile - do the dirty work
'
sub DeleteFile(FolderPath, Filename)
   Wscript.Echo "Deleting " & FolderPath & "\" & Filename
   on error resume next 
   if not FakeAction then objFSO.DeleteFile(FolderPath & "\" & Filename)
   if Err.Number<>0 then WScript.Echo "Error: " & CStr(Err.Number) & ": " & Error.Desription
   on error goto 0
end Sub

'
' MoveFile - do not delete but move files to another destination
'
Sub MoveFile(FolderPath, Filename, TargetFolder)
   Wscript.Echo "Moving " & FolderPath & "\" & Filename & " to " & TargetFolder
   On Error Resume Next
   ' does target path already exist?
   If Not objFSO.FolderExists(TargetFolder) Then
      ' if not, try to create it
      WScript.Echo "Target Folder " & TargetFolder & " does not exist. Trying to create it..."
      If Not FakeAction Then objFSO.CreateFolder(TargetFolder)
      If Err.Number<>0 Then WScript.Echo "Error: " & CStr(Err.Number) & ": " & Error.Description
   End If
   If Not FakeAction Then objFSO.MoveFile(TargetFolder & "\" & Filename)
   If Err.Number<>0 Then WScript.Echo "Error: " & CStr(Err.Number) & ": " & Error.Description
   On Error GoTo 0
End Sub


'
' ShowUsage - display some help
'
sub ShowUsage()
   WScript.Echo "tidyup" & vbCRLF & _
          "Delete (recursively) old files (of given type(s)) in a given folder that are at least <n> day(s) old." & vbCRLF & vbCRLF & _
          "cscript tidyup.vbs -f[older] <folder> -d[ays] <days> [-r[ecursive]] [-[f]i[lter] <filterstring>] [-m[ove] <targetfolder>]" & vbCRLF & _
          "  -f <folder>: folder to erase old files from" & vbCRLF & _
          "  -d <days>: number of days files must exists to get deleted." & vbCRLF & _
          "  -r (optional): recursive deletion" & vbCRLF & _
          "  -i <filterlist> (optional): file filter (in format ""<regextype1>||<regextype2>||...||<regextypen>""; use || to separate regexes" & vbCRLF & _
          "  -m <targetfolder> (optional) move files into target folder <targetfolder> (do not delete them). If <targetfolder> does not exist, Tidy Up will try to mkdir it." & vbCRLF & _
          "  -a (optional): fake action: don't delete but show what would be done" & vbCRLF & _
          "  -? (optional): this help" & vbCRLF & vbCRLF & _
          "examples: cscript tidyup.vbs -f c:\temp -d 20" & vbCRLF & _ 
          "          cscript tidyup.vbs -f c:\temp -d 20 -r" & vbCRLF & _
          "          cscript tidyup.vbs -f c:\temp -d 20 -r -m d:\temp" & vbCRLF & _ 
          "          cscript tidyup.vbs -f c:\temp -d 20 -i ""log\.\d{4}-\d{2}-\d{2}||\.bak||\.te?mp""" 
end sub	
