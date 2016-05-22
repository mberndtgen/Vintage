'
' ldap2user
'
' PURPOSE: create/delete users in AD by retrieving data from excel2ldap 
' LAST EDITED: 05-09-27
' AUTHOR(S): Berndtgen, Düsseldorf, 2005
' SYNTAX: cscript 
' VERSION: 1.0
' TODOS: support for "Update" of user data
'

'Option Explicit

dim Domain, OUDN, InputFilename, LogFilename, Outputfilename ' as string
dim bVerbose ' as boolean

'
' change settings section start
'
Domain = "LDAP://server/" 
OUDN = "ou=Users,OU=COMPANY,OU=Customers,dc=user,dc=domain,dc=company"
bVerbose = true 
InputFilename = "input.txt"
Outputfilename = "output.txt"
LogFilename = "logfile.txt"
'
' change settings section end
'

Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_DELETE = 4 
const ADS_UF_NORMAL_ACCOUNT = 512 

dim StdIn, StdOut, oArgumente ' as object
dim dso 'As IADsOpenDSObject
dim ClassArray 'As Variant
dim InputLine 'as String
dim Dict, DictUsers ' as dictionary
dim DictElement ' as string
dim Password ' as string (password return value)
Dim filesys, LogFile, OutputFile ' as object
dim bLogging ' as boolean

Init
HandleUser(InputFilename)
CleanUp

'
' subroutines
'

'
' Init: do some initialization stuff
'
sub Init
	' examine command line arguments
	bLogging = false

	set StdIn = Wscript.StdIn
	set StdOut = Wscript.StdOut
	set Dict = CreateObject("Scripting.Dictionary")
	set DictUsers = CreateObject("Scripting.Dictionary")
	Set filesys = CreateObject("Scripting.FileSystemObject") 

	set oArgumente = WScript.Arguments
	if oArgumente.Count > 0 then 
		do 
	  		if UCase(oArgumente(i)) = "-V" or UCase(oArgumente(i)) = "-VERBOSE" then
	  			' verbose output
	  			bVerbose = true
	  	  	elseif UCase(oArgumente(i)) = "-S" or UCase(oArgumente(i)) = "-SOURCE" then
	  	  		' source file (output from excel2ldap)
	  	  		i = i + 1
	  	  		InputFilename = oArgumente(i)
				if err.number<>0 then 
					WScript.Echo "ERROR: invalid input data (source file). Aborting."
					WScript.Quit
				end if  	  		
	  	  	elseif UCase(oArgumente(i)) = "-L" or UCase(oArgumente(i)) = "-LOGFILE" then
	  	  		' logfile 
	  	  		i = i + 1
	  	  		LogFilename = oArgumente(i)
				if err.number<>0 then 
					WScript.Echo "ERROR: invalid input data (logfile). Aborting."
					WScript.Quit
				else 
					bLogging = true
					on error resume next
					Set LogFile = filesys.CreateTextFile(LogFilename, True) 
					if err.number<>0 then
						WScript.Echo "ERROR: can't write " & LogFilename & ". " & Err.Description & ". Aborting."
						WSCript.Quit
					end if
					on error goto 0
				end if  	  		
	  	  	elseif UCase(oArgumente(i)) = "-O" or UCase(oArgumente(i)) = "-OUTPUT" then
	  	  		' output file (contains user account name and user password)
	  	  		i = i + 1
	  	  		if oArgumente(i) <> "output.txt" then 
	  	  			' if output file name isn't changed then initialize it later
		  	  		Outputfilename = oArgumente(i)
					if err.number<>0 then 
						WScript.Echo "ERROR: invalid input data (outputfile). Aborting."
						WScript.Quit
					else 
						on error resume next
						Set OutputFile = filesys.CreateTextFile(OutputFilename, True) 
						if err.number<>0 then
							WScript.Echo "ERROR: can't write " & OutputFilename & ". " & Err.Description & " .Aborting."
							WSCript.Quit
						end if
						on error goto 0
					end if  	  		
				end if
	  	  	elseif UCase(oArgumente(i)) = "-DN" then
	  	  		' set UODN
	  	  		i = i + 1
	  	  		OUDN = oArgumente(i)
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
	  	  	elseif UCase(oArgumente(i)) = "-H" or UCase(oArgumente(i)) = "-HELP" or oArgumente(i) = "?" or oArgumente(i) = "-?" then
	  	  		ShowUsage
	  	  		WScript.Quit
	  	  	end if
			i = i + 1		
		loop until i>=oArgumente.Count
	end if
end sub

'
' CleanUp: tidy up objects
'
sub CleanUp
	if bVerbose then WScript.Echo "Tidying up."
	set Dict = nothing
	set DictUsers = nothing
	if bLogging then 
		if IsObject(Logfile) then 
			LogFile.Close   	
			set LogFile = nothing
		end if
	end if
	if IsObject(OutputFile) then 
		OutputFile.Close
		set OutputFile = nothing
	end if
	set filesys = nothing
end sub

'
' HandleUser - read input file into array of dictionaries
'
sub HandleUser(ByVal InputFilename)
	dim oFilesys, oFile ' FilesystemObject, Readfile
	dim Contents ' content of oFile
	dim InputData, KeyValArray ' as array
	dim IDElement ' string
	dim Lines() ' array of lines per InputFilename
	dim i ' Array counter
	
	i = 0
	set oFilesys = CreateObject("Scripting.FileSystemObject")
	if not oFilesys.FileExists(InputFilename) then 
		' no file? abort
		printf "ERROR: Input File does not exist. Aborting."
		WScript.Quit
	end if

	on error resume next

	set oFile = oFilesys.OpenTextFile(InputFilename, 1, false)
	if err.number<>0 then
		printf "ERROR: Can't read input file: " & Err.Message & ". Aborting."
		WScript.Quit
	end if

	if (OutputFilename = "output.txt") then ' name not changed and thus no output file initialized
		Set OutputFile = filesys.CreateTextFile(OutputFilename, True) 
		if err.number<>0 then
			WScript.Echo "ERROR: can't write " & OutputFilename & ". " & Err.Description & " .Aborting."
			WSCript.Quit
		end if
	end if
	
	on error goto 0
	
	do while not oFile.AtEndOfStream 
		Contents = oFile.ReadLine
		
		if bVerbose then printf "Reading " & Contents
		
		if left(Contents, 1) <> "#" then ' "#" is a command symbol
			InputData = split(Contents, "##") ' split input line 
			for each IDElement in InputData 
				' fill dictionary obj with key/value pairs from input line
				KeyValArray = split(IDElement, ":")
				Dict.Add KeyValArray(0), KeyValArray(1)
			next
			'
			' Dict carries all user information from the line currently read in
			'
			select case UCase(Dict.Item("action"))
				case "N"
					' new user
					printf "Creating new Account"
					
					if FindSingleUser(Dict.Item("sAMAccountName")) then
						printf Dict.Item("userPrincipalName") & " is already in the directory!"
					else 
						printf "User " & Dict.Item("userPrincipalName") & " not found. Will create user."
					end if
					CreateUser ' make user record
				case "C"
					' update user
					printf "Updating User " & Dict.Item("cn")
					if FindSingleUser(Dict.Item("sAMAccountName")) then
						UpdateUser
					else 
						printf "User " & Dict.Item("userPrincipalName") & " not found. Cannot update."
					end if
				case "D"
					' delete user
					printf "Deleting User " & Dict.Item("cn")
					on error resume next
					Set objOU = GetObject(Domain & OUDN)					
					objOU.Delete "user", "cn=" & Dict.Item("cn")
					if Err.Number <> 0 then 
						printf "ERROR: Could not delete user"
					else
						printf "Done."
					end if
					on error goto 0
			end select		
			' Work is done for this line, dispose data
			Dict.RemoveAll
		end if
	loop

	if bVerbose then PrintDict(Dict)
end sub



'
' find a single user (using LDAP directly)
'
function FindSingleUser(byVal cn) ' as boolean
	dim UserDN ' as string
	dim User ' as variant
	
	on error resume next
	
	cn = trim(cn)
	UserDN = Domain & "cn=" & cn & "," & OUDN 
	wscript.echo "FindSingleUser: " & UserDN
	set User = GetObject(UserDN)
	if IsObject(User) then
		if bVerbose then WScript.Echo "User found."
		FindSingleUser = true
	else 
		if bVerbose then WScript.Echo "User not found."
		FindSingleUser = false
	end if

	on error goto 0
	if err.number<>0 then printf "ERROR searching for " & UserDN
end function

'
' find a single user on DictUser
'
function FindUser(byVal UserEmailAddress)
	if DictUsers.Exists(UserEmailAddress) then 
		FindUser = true
	else 
		FindUser = false
	end if
end function


'
' UpdateUser - update an already present record
'
sub UpdateUser
	dim User ' as object
	dim Group 'As IADs 
	dim grp ' as object
	
	if AccountNameAlreadyPresent(Dict.Item("sAMAccountName")) then 
		on error resume next

		set User = GetObject(Domain & "cn=" & Dict.Item("cn") & "," & OUDN)

		if trim(Dict.Item("givenName"))<>"" then User.Put "givenName", trim(Dict.Item("givenName"))
		if trim(Dict.Item("sn"))<>"" then User.Put "sn", trim(Dict.Item("sn"))
		if trim(Dict.Item("co"))<>"" then 
			User.Put "co", trim(Dict.Item("co"))
			if GetCountryCode(Dict.Item("co"))<>"" then User.Put "c", GetCountryCode(Dict.Item("co")) ' "c" is a special 2-digit-code
		end if
		if trim(Dict.Item("title"))<>"" then User.Put "title", trim(Dict.Item("title"))
		if trim(Dict.Item("department"))<>"" then User.Put "department", trim(Dict.Item("department"))
		if trim(Dict.Item("displayName"))<> "" then User.Put "displayName", trim(Dict.Item("displayName"))
		if trim(Dict.Item("company"))<>"" then User.Put "company", trim(Dict.Item("company"))
		if trim(Dict.Item("telephoneNumber"))<>"" then User.Put "telephoneNumber", trim(Dict.Item("telephoneNumber"))
		if trim(Dict.Item("facsimileTelephoneNumber"))<>"" then User.Put "facsimileTelephoneNumber", trim(Dict.Item("facsimileTelephoneNumber"))
		if trim(Dict.Item("mail"))<>"" then User.Put "mail", trim(Dict.Item("mail")) 
		
		' try
		User.SetInfo
		' catch
		if Err.Number <> 0 then
			WScript.Echo "ERROR updating user: " & Err.Number & ": " & Err.Description
			set User = nothing
			exit sub
		end if

		' put user into group(s)
		aGroups = split(Dict.Item("memberOf"), ";")
		for each grp in aGroups
			if bVerbose then WScript.Echo "Putting " & Dict.Item("cn") & " into " & grp & " ## " & Domain & grp
			set Group = GetObject(Domain & grp)
			if IsObject(Group) then
				' try
				Group.Add(Domain & "cn=" & Dict.Item("cn") & "," & OUDN)
				' catch
				if Err.Number <> 0 then WScript.Echo "ERROR putting user into group " & grp & ": " & CStr(Err.Number) & ": " & Err.Description
			else
				WScript.Echo "ERROR: group " & grp & " does not exist: " & CStr(Err.Number) & ": " & Err.Description
				exit for
			end if
			set Group = nothing
		next
		
		' delete user from group(s)
		aGroups = split(Dict.Item("noMemberOf"), ";")
		for each grp in aGroups
			if bVerbose then WScript.Echo "Deleting " & Dict.Item("cn") & " from " & grp & " ## " & Domain & grp
			set Group = GetObject(Domain & grp)
			if IsObject(Group) then
				' try
				Group.Remove(Domain & "cn=" & Dict.Item("cn") & "," & OUDN)
				' catch
				if Err.Number <> 0 then WScript.Echo "ERROR deleting user from group " & grp & ": " & CStr(Err.Number) & ": " & Err.Description
			else
				WScript.Echo "ERROR: group " & grp & " does not exist: " & CStr(Err.Number) & ": " & Err.Description
				exit for
			end if
			set Group = nothing
		next
		
		on error goto 0
	end if
end sub

'
' CreateUser - create a user and do some validity tests
'
sub CreateUser
	dim Container 'As IADsContainer
	dim User 'As IADs
	dim Group 'As IADs
	dim objRegEx, Match, Matches, StrReturnStr 
	dim Counter, sName
	dim OldsAMAccount

	Counter = 2
	on error resume next
	
	OldsAMAccount = Dict.Item("sAMAccountName")

	'
	' test validity of e-mail address
	' 
	Set objRegEx = New RegExp 
	objRegEx.Global = True
	objRegEx.IgnoreCase = True 'set case insensitive
	objRegEx.Pattern = "\b[A-Z0-9._%-]+@[A-Z0-9._%-]+\.[A-Z]{2,4}\b" 'set pattern
	Set Matches = objRegEx.Execute(Dict.Item("mail")) 
	if Matches.Count = 0 then
		printf "ERROR: User " & Dict.Item("sAMAccountName") & " has incorrect E-mail format " & Dict.Item("mail")
		if not (OutputFile = nothing) then OutputFile.WriteLine Dict.Item("sAMAccountName") & " has an incorrect e-mail format, User will not be installed."
		exit sub
	end if
	
	'
	' sAMAccountName already present; try to concat 1st letter of the givenName
	'
	if AccountNameAlreadyPresent(Dict.Item("sAMAccountName")) then 
		WScript.Echo "sAMAccountName " & Dict.Item("sAMAccountName") & " already present."
		if trim(Dict.Item("givenName"))<>"" then 
			Dict.Item("sAMAccountName") = Dict.Item("sAMAccountName") & left(trim(Dict.Item("givenName")), 1)
			printf "sAMAccountName changed to " & Dict.Item("sAMAccountName")
		else
			printf "Could not change sAMAccountName because givenName is empty"
		end if
	end if

	'
	' sAMAccountName still present; attach a number 
	'
	if AccountNameAlreadyPresent(Dict.Item("sAMAccountName")) then 
		sName = Dict.Item("sAMAccountName")
		printf "sAMAccountName " & sName & " already / still present. Will attach a number."
		do 
			sName = Dict.Item("sAMAccountName") & CStr(Counter)
			if AccountNameAlreadyPresent(sName) = false then exit do
			Counter = Counter + 1
		loop
		Dict.Item("sAMAccountName") = sName
		printf "sAMAccountName changed to " & Dict.Item("sAMAccountName")
	end if
	
	call SynchronizeNames(OldsAMAccount) ' update cn and userPrincipalName with sAMAccountName 
	
	set Container = GetObject(Domain & OUDN)
	set User = Container.Create("user", "cn=" & Dict.Item("cn"))
	
	' tell us what is happening now
	if bVerbose then call PrintRecord

	' fill user data with dictionary content
	if trim(Dict.Item("sAMAccountName"))<>"" then User.Put "sAMAccountName", trim(Dict.Item("sAMAccountName"))
	if trim(Dict.Item("userPrincipalName"))<>"" then User.Put "userPrincipalName", trim(Dict.Item("userPrincipalName"))
	if trim(Dict.Item("givenName"))<>"" then User.Put "givenName", trim(Dict.Item("givenName"))
	if trim(Dict.Item("sn"))<>"" then User.Put "sn", trim(Dict.Item("sn"))
	if trim(Dict.Item("cn"))<>"" then User.Put "cn", trim(Dict.Item("cn"))
	if trim(Dict.Item("co"))<>"" then 
		User.Put "co", trim(Dict.Item("co"))
		if GetCountryCode(Dict.Item("co"))<>"" then User.Put "c", GetCountryCode(Dict.Item("co")) ' "c" is a special 2-digit-code
	end if
	if trim(Dict.Item("title"))<>"" then User.Put "title", trim(Dict.Item("title"))
	if trim(Dict.Item("department"))<>"" then User.Put "department", trim(Dict.Item("department"))
	if trim(Dict.Item("displayName"))<> "" then User.Put "displayName", trim(Dict.Item("displayName"))
	if trim(Dict.Item("company"))<>"" then User.Put "company", trim(Dict.Item("company"))
	if trim(Dict.Item("telephoneNumber"))<>"" then User.Put "telephoneNumber", trim(Dict.Item("telephoneNumber"))
	if trim(Dict.Item("facsimileTelephoneNumber"))<>"" then User.Put "facsimileTelephoneNumber", trim(Dict.Item("facsimileTelephoneNumber"))
	if trim(Dict.Item("mail"))<>"" then User.Put "mail", trim(Dict.Item("mail")) 
	if trim(Dict.Item("description"))<>"" then User.Put "description", trim(Dict.Item("description"))
	
	' try
	User.SetInfo
	' catch
	if Err.Number <> 0 then
		WScript.Echo "ERROR creating user: " & Err.Number & ": " & Err.Description
		set User = nothing
		exit sub
	end if
	
	' give user a password
	Password = RandomPW(10)
	printf "User gets Password: " & Password 
	' try
	User.SetPassword(Password)
	' catch
	if Err.Number <> 0 then 
		WScript.Echo "ERROR setting password: " & CStr(Err.Number) & ": " & Err.Description
		set User = nothing
		exit sub
	end if
	
	User.AccountDisabled = false
	User.SetInfo
	
	' put user into group(s)
	aGroups = split(Dict.Item("memberOf"), ";")
	for each grp in aGroups
		if bVerbose then WScript.Echo "Putting " & Dict.Item("cn") & " into " & grp & " ## " & Domain & grp
		set Group = GetObject(Domain & grp)
		if IsObject(Group) then
			Group.Add(Domain & "cn=" & Dict.Item("cn") & "," & OUDN)
			if Err.Number <> 0 then 
				WScript.Echo "ERROR putting user into group " & grp & ": " & CStr(Err.Number) & ": " & Err.Description
				'set Group = nothing
				'exit for
			end if
		else
			WScript.Echo "ERROR: group " & grp & " does not exist: " & CStr(Err.Number) & ": " & Err.Description
			exit for
		end if
		set Group = nothing
	next
	
	Outputfile.WriteLine Dict.Item("sAMAccountName") & chr(9) & chr(9) & Password
	
	on error goto 0	
	
end sub

'
' AccountNameAlreadyPresent: detect if sAMAccountName is already being used
'
function AccountNameAlreadyPresent(byVal Account) ' as boolean
	dim objConnection ' as object
	dim objCommand ' as object
	dim objRecordSet ' as object
	dim LDAPString ' as string
	
	on error resume next
	
	Set objConnection = CreateObject("ADODB.Connection")
	if Err.Number <> 0 then 
		WScript.Echo "AccountNameAlreadyPresent: ERROR getting ABODB.Connection : " & CStr(Err.Number) & ": " & Err.Description
		WScript.Quit
	end if
	
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	 
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	if Err.Number <> 0 then 
		WScript.Echo "AccountNameAlreadyPresent: ERROR getting ADODB.Command : " & CStr(Err.Number) & ": " & Err.Description
		WScript.Quit
	end if

	LDAPString = "<" & Domain & OUDN & ">;"		 
	objCommand.CommandText = LDAPString & "(&(objectCategory=User)(sAMAccountName=" & Account & "));sAMAccountName;subtree"
	'StdOut.WriteLine ">>" & objCommand.CommandText
	Set objRecordSet = objCommand.Execute
	
	if Err.Number <> 0 then 
		WScript.Echo "AccountNameAlreadyPresent: ERROR getting executing ADODB command : " & CStr(Err.Number) & ": " & Err.Description
		WScript.Quit
	end if
	 
	If objRecordset.RecordCount = 0 Then
		WScript.Echo "Account not present"
	    AccountNameAlreadyPresent = false
	Else
		WScript.Echo "Account already present"
	    AccountNameAlreadyPresent = true
	End If
	 
	objConnection.Close
	
	on error goto 0
end function

'
' SynchronizeNames - copy sAMAccountName to cn and userPrincipalName
'
sub SynchronizeNames(OldsAMAccount)
	dim Parts ' as array
	
	' sync cn with sAMAccountName
	if OldsAMAccount = Dict.Item("cn") then Dict.Item("cn") = Dict.Item("sAMAccountName")
	
	' sync userPrincipalname with sAMAccountName
	Parts = split(Dict.Item("userPrincipalName"), "@")
	if Parts(0) = OldsAMAccount then 
		Parts(0) = trim(Dict.Item("sAMAccountName")) & "@"
		Dict.Item("userPrincipalName") = join(Parts)
		Dict.Item("userPrincipalName") = Replace(Dict.Item("userPrincipalName"), " ", "") ' eliminate " " after "@" in userPrincipalName
	end if
	
end sub

'
' RandomPW - password generator
'   taken from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=6824&lngWId=4
'
Function RandomPW(myLength)
	'These constant are the minimum and maximum length for random
	'length passwords.  Adjust these values to your needs.
	Const minLength = 6
	Const maxLength = 20
	
	Dim X, Y, strPW
	
	If myLength = 0 Then
		Randomize
		myLength = Int((maxLength * Rnd) + minLength)
	End If
	
	For X = 1 To myLength
		'Randomize the type of this character
		Y = Int((3 * Rnd) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase
		
		Select Case Y
			Case 1
				'Numeric character
				Randomize
				strPW = strPW & CHR(Int((9 * Rnd) + 48))
			Case 2
				'Uppercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 65))
			Case 3
				'Lowercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 97))
		End Select
	Next
	RandomPW = strPW
end Function


'
' GetCountryCode: return country code suitable to (english) country name
'
function GetCountryCode(byval Country) 
	dim ID ' as string
	
	Select case Country
	    case "Andorra" ID = "AD"
	    case "Lesotho" ID = "LS"
	    case "Afghanistan"	ID = "AF"
	    case "Lithuania" ID = "LT"
	    case "United Arab Emirates" ID = "AE"
	    case "Luxembourg" ID = "LU"
	    case "Antigua and Barbuda" ID = "AG"
	    case "Latvia" ID = "LV"
	    case "Anguilla" ID = "AI"
	    case "Libya" ID = "LY"
	    case "Albania" ID = "AL"
	    case "Morocco" ID = "MA"
	    case "Armenia" ID = "AM"
		case "Monaco" ID = "MC"
		case "Netherlands Antilles" ID = "AN"
		case "Moldova" ID = "MD"
		case "Angola" ID = "AO"
		case "Madagascar" ID = "MG"
		case "Antarctica" ID = "AQ"
		case "Marshall Islands" ID = "MH"
		case "Argentina" ID = "AR"
		case "Macedonia" ID = "MK"
		case "American Samoa" ID = "AS"
		case "Mali" ID = "ML"
		case "Austria" ID = "AT"
		case "Myanmar" ID = "MM"
		case "Australia" ID = "AU"
		case "Mongolia" ID = "MN"
		case "Aruba" ID = "AW"
		case "Macau" ID = "MO"
 		case "Azerbaijan" ID = "AZ"
 		case "Northern Mariana Islands" ID = "MP"
		case "Bosnia and Herzegovina" ID = "BA"
		case "Martinique" ID = "MQ"
		case "Barbados" ID = "BB"
		case "Mauritania" ID = "MR"
		case "Bangladesh" ID = "BD"
		case "Montserrat" ID = "MS"
		case "Belgium" ID = "BE"
		case "Malta" ID = "MT"
		case "Burkina Faso" ID = "BF"
		case "Mauritius" ID = "MU"
		case "Bulgaria" ID = "BG"
		case "Maldives" ID = "MV"
		case "Bahrain" ID = "BH"
		case "Malawi" ID = "MW"
		case "Burundi" ID = "BI"
		case "Mexico" ID = "MX"
		case "Benin" ID = "BJ"
		case "Malaysia" ID = "MY"
		case "Bermuda" ID = "BM"
		case "Mozambique" ID = "MZ"
		case "Brunei Darussalam" ID = "BN"
		case "Namibia" ID = "NA"
		case "Bolivia" ID = "BO"
		case "New Caledonia" ID = "NC"
		case "Brazil" ID = "BR"
		case "Niger" ID = "NE"
		case "Bahamas" ID = "BS" 
		case "Norfolk Island" ID = "NF"
		case "Bhutan" ID = "BT"
		case "Nigeria" ID = "NG"
		case "Bouvet Island" ID = "BV"
		case "Nicaragua" ID = "NI"
		case "Botswana" ID = "BW"
		case "Netherlands" ID = "NL"
		case "Belarus" ID = "BY"
		case "Norway" ID = "NO"
		case "Belize" ID = "BZ"
		case "Nepal" ID = "NP"
 		case "Canada" ID = "CD"
		case "Nauru" ID = "NR"
		case "Cocos (Keeling) Islands" ID = "CC"
		case "Neutral Zone" ID = "NT"
		case "Central African Republic" ID = "CF"
		case "Niue"  ID = "NU"
		case "Congo" ID = "CG"
		case "New Zealand" ID = "NZ" ' (Aotearoa) 
		case "Switzerland" ID = "CH"
		case "Oman" ID = "OM"
		case "Cote D'Ivoire (Ivory Coast)" ID = "CI"
		case "Panama" ID = "PA"
		case "Cook Islands" ID = "CK"
		case "Peru" ID = "PE"
 		case "Chile" ID = "CL"
 		case "French Polynesia" ID = "PF"
		case "Cameroon" ID = "CM"
		case "Papua New Guinea" ID = "PG"
		case "China" ID = "CN"
		case "Philippines" ID = "PH"
		case "Colombia" ID = "CO"
		case "Pakistan" ID = "PK"
		case "Costa Rica" ID = "CR"
		case "Poland" ID = "PL"
		case "Czechoslovakia (former)" ID = "CS"
		case "St. Pierre and Miquelon" ID = "PM"
		case "Cuba" ID = "CU"
		case "Pitcairn" ID = "PN"
		case "Cape Verde" ID = "CV"
		case "Puerto Rico" ID = "PR"
		case "Christmas Island" ID = "CX"
		case "Portugal" ID = "PT"
		case "Cyprus" ID = "CY"
		case "Palau" ID = "PW"
		case "Czech Republic" ID = "CZ"
		case "Paraguay" ID = "PY"
		case "Germany" ID = "DE"
		case "Deutschland" ID = "DE"
		case "Qatar" ID = "QA"
		case "Djibouti" ID = "DJ"
		case "Reunion" ID = "RE"
		case "Denmark" ID = "DK"
		case "Romania" ID = "RO"
		case "Dominica" ID = "DM"
		case "Russian Federation" ID = "RU"
		case "Dominican Republic" ID = "DO"
		case "Rwanda" ID = "RW"
		case "Algeria" ID = "DZ"
		case "Saudi Arabia" ID = "SA"
		case "Ecuador" ID = "EC"
		case "Solomon Islands" ID = "SB"
		case "Estonia" ID = "EE"
		case "Seychelles" ID = "SC"
		case "Egypt" ID = "EG"
		case "Sudan" ID = "SD"
		case "Western Sahara" ID = "EH"
		case "Sweden" ID = "SE"
		case "Eritrea" ID = "ER"
		case "Singapore" ID = "SG"
		case "Spain" ID = "ES"
		case "St. Helena" ID = "SH"
		case "Ethiopia" ID = "ET"
		case "Slovenia" ID = "SI"
		case "Finland" ID = "FI"
		case "Svalbard and Jan Mayen Islands" ID = "SJ"
		case "Fiji" ID = "FJ"
		case "Slovak Republic" ID = "SK"
		case "Falkland Islands (Malvinas)" ID = "FK"
		case "Sierra Leone" ID = "SL"
		case "Micronesia" ID = "FM"
		case "San Marino" ID = "SM"
		case "Faroe Islands" ID = "FO"
		case "Senegal" ID = "SN"
		case "France" ID = "FR"
		case "Somalia" ID = "SO"
		case "France, Metropolitan" ID = "FX"
		case "Suriname" ID = "SR"
		case "Gabon" ID = "GA"
		case "Sao Tome and Principe" ID = "ST"
		case "Great Britain (UK)" ID = "GB"
		case "USSR (former)" ID = "SU"
		case "Grenada" ID = "GD"
		case "El Salvador" ID = "SV"
		case "Georgia" ID = "GE"
		case "Syria" ID = "SY"
		case "French Guiana" ID = "GF"
		case "Swaziland" ID = "SZ"
		case "Ghana" ID = "GH"
		case "Turks and Caicos Islands" ID = "TC"
		case "Gibraltar" ID = "GI"
		case "Chad" ID = "TD"
		case "Greenland" ID = "GL"
		case "French Southern Territories" ID = "TF"
		case "Gambia" ID = "GM"
		case "Togo" ID = "TG"
		case "Guinea" ID = "GN"
		case "Thailand" ID = "TH"
		case "Guadeloupe" ID = "GP"
		case "Tajikistan" ID = "TJ"
		case "Equatorial Guinea" ID = "GQ"
		case "Tokelau" ID = "TK"
		case "Greece" ID = "GR"
		case "Turkmenistan" ID = "TM"
		case "S. Georgia and S. Sandwich Islands" ID = "GS" 
		case "Tunisia" ID = "TN"
		case "Guatemala" ID = "GT"
		case "Tonga" ID = "TO"
		case "Guam" ID = "GU"
		case "East Timor" ID = "TP"
		case "Guinea-Bissau" ID "GW"
		case "Turkey" ID = "TR"
		case "Guyana" ID = "GY"
		case "Trinidad and Tobago" ID = "TT"
		case "Hong Kong" ID = "HK"
		case "Tuvalu" ID = "TV"
		case "Heard and McDonald Islands" ID = "HM"
		case "Taiwan" ID = "TW"
		case "Honduras" ID = "HN"
		case "Tanzania" ID = "TZ"
		case "Croatia (Hrvatska)" ID = "HR"
		case "Ukraine" ID = "UA"
		case "Haiti" ID = "HT"
		case "Uganda" ID = "UG"
		case "Hungary" ID = "HU"
		case "United Kingdom" ID = "UK"
		case "Indonesia" ID = "ID"
		case "US Minor Outlying Islands" ID = "UM"
		case "Ireland" ID = "IE"
		case "United States" ID = "US"
		case "Israel" ID = "IL"
		case "Uruguay" ID = "UY"
		case "India" ID = "IN"
		case "Uzbekistan" ID = "UZ"
		case "British Indian Ocean Territory" ID = "IO"
		case "Vatican City State" ID = "VA"
		case "Iraq" ID = "IQ"
		case "Saint Vincent and the Grenadines" ID = "VC"
		case "Iran" ID = "IR"
		case "Venezuela" ID = "VE"
		case "Iceland" ID = "IS"
		case "Virgin Islands (British)" ID = "VG"
		case "Italy" ID = "IT"
		case "Virgin Islands (U.S.)" ID = "VI"
		case "Jamaica" ID = "JM"
		case "Vietnam" ID = "VN"
		case "Jordan" ID = "JO"
		case "Vanuatu" ID = "VU"
		case "Japan" ID = "JP"
		case "Wallis and Futuna Islands" ID = "WF"
		case "Kenya" ID = "KE"
		case "Samoa" ID = "WS"
		case "Kyrgyzstan" ID = "KG"
		case "Yemen" ID = "YE"
		case "Cambodia" ID = "KH"
		case "Mayotte" ID = "YT"
		case "Kiribati" ID = "KI"
		case "Yugoslavia" ID = "YU"
		case "Comoros" ID = "KM"
		case "South Africa" ID = "ZA"
		case "Saint Kitts and Nevis" ID = "KN"
		case "Zambia" ID = "ZM"
		case "Korea (North)" ID = "KP"
		case "Zaire" ID = "ZR"
		case "Korea (South)" ID = "KR"
		case "Zimbabwe" ID = "ZW"
		case "Kuwait" ID = "KW"
		case "US Commercial" ID = "COM"
		case "Cayman Islands" ID = "KY"
		case "US Educational" ID = "EDU"
		case "Kazakhstan" ID = "KZ"
		case "US Government" ID = "GOV"
		case "Laos" ID = "LA"
		case "International" ID = "INT"
		case "Lebanon" ID = "LB" 
		case "US Military" ID = "MIL"
		case "Saint Lucia" ID = "LC" 
		case "Network" ID = "NET"
		case "Liechtenstein" ID = "LI"
		case "Non-Profit Organization" ID = "ORG"
		case "Sri Lanka" ID = "LK"
		case "Old Style Arpanet" ID = "ARPA"
		case "Liberia" ID = "LR"
		case "NATO Field" ID = "NATO"
		case else ID = "N/A"
	end select

	GetCountryCode = ID
end function 

'
' PrintRecord - verbose print dictionary's content
'
sub PrintRecord
	printf "sAMAccountName:" & Dict.Item("sAMAccountName") 
	printf "userPrincipalName:" & Dict.Item("userPrincipalName")
	printf "givenName:" & Dict.Item("givenName")
	printf "sn:" & Dict.Item("sn")
	printf "cn:" & Dict.Item("cn")
	printf "co:" & Dict.Item("co")
	printf "c:" & GetCountryCode(Dict.Item("co"))
	printf "title:" & Dict.Item("title")
	printf "department:" & Dict.Item("department")
	printf "displayName:" & Dict.Item("displayName")
	printf "company:" & Dict.Item("company")
	printf "telephoneNumber:" & Dict.Item("telephoneNumber")
	printf "facsimileTelephoneNumber:" & Dict.Item("facsimileTelephoneNumber")
	printf "memberOf" & Dict.Item("memberOf")
	printf "mail:" & Dict.Item("mail")
	printf "description:" & Dict.Item("description")
end sub

'
' PrintDict - print content of a dictionary
'
sub PrintDict(byval Dic)
	dim i, a 

	a = Dic.Keys
	for i = 0 To Dic.Count -1 
   		WScript.Echo a(i) & " = " & Dic.Item(a(i)) 
	next 	
end sub

'
' GetUserlist - get filtered user list
'
sub GetUserlist
	dim strFilter, strAttrs, strScope ' as string
	dim objConn, objRS ' as variant
	dim LDAPString ' as string
	
	LDAPString = "<" & Domain & OUDN & ">;" ' for one special user, add a "cn=Testuser"
	
	strFilter = "(&(objectclass=user)(objectcategory=person));" 
	strAttrs  = "displayName, mail;"
	strScope  = "subtree"
	
	WScript.Echo "Retrieving Users..."
	
	set objConn = CreateObject("ADODB.Connection")
	objConn.Provider = "ADsDSOObject"
	objConn.Open "Active Directory Provider"
	
	set objRS = objConn.Execute(LDAPString & strFilter & strAttrs & strScope)
	objRS.MoveFirst
	while Not objRS.EOF
		if objRS.Fields(1).Value <> "" then 
			DictUsers.Add objRS.Fields(1).Value, objRS.Fields(0).Value
	   		'StdOut.WriteLine "Adding " & objRS.Fields(1).Value & " / " & objRS.Fields(0).Value ' display Attrs as in strAttrs set
	   	else 
	   		printf "WARNING: Problem with User " & objRS.Fields(1).Value & " / " & objRS.Fields(0).Value & " (possibly missing values)"
	   	end if
	   	objRS.MoveNext
	wend
	set objRS = nothing
end sub

' 
' printf - print on screen and into logfile (if wanted)
'
sub printf(Outputstr)
	WScript.Echo Outputstr
	if bLogging then LogFile.WriteLine Outputstr
end sub


'
' ShowUsage - display some help
'
Sub ShowUsage()
    Wscript.echo "ldap2user" & _
    "create/delete/update users in AD by retrieving data from excel2ldap" & vbCRLF & vbCRLF & _
    "cscript ldap2user.vbs [-v[erbose]] [-s[ource] <sourcefile>] [-o[utputfile]] [-l[ogfile] <logfilename>] [-dn <oudn>] [-do[main] <domaindn>] [[-]?]" & vbCRLF & _
    "  -v (optional): verbose mode (default: off)" & vbCRLF & _
    "  -s <sourcefile> (optional): input file (as given from excel2ldap, default: input.txt)" & vbCRLF & _
    "  -o <outputfile> (optional): output file (default: output.txt)" & vbCRLF & _
    "  -l <logfile> (optional): write various output into <logfile>" & vbCRLF & _
    "  -dn <oudn> (optional): organizational DN, e.g. ou=Users,OU=COMPANY,OU=Customers,dc=user,dc=domain,dc=company" & vbCRLF & _
    "  -do <domain> (optional): Domain name, e.g. LDAP://server/" & vbCRLF & _
    "  -? (optional): this help" & vbCRLF & _
    "example: cscript ldap2user.vbs " & vbCRLF & _
    "additional information: any user is identified by his sAMAccountName. Any new " & vbCRLF & _
    "user will get a sAMAccountName (and userPrincipalName) as defined in the " & vbCRLF & _
    "<sourcefile>. If there is a homonymous sAMAccountName ldap2user will attach " & vbCRLF & _
    "the 1st letter of the givenName (if given). " & vbCRLF & _
    "If there is still an identic sAMAccountName a number will be attached to the " & vbCRLF & _
    "new user's sAMAccountName (this happens as often as sAMAccountName eventually" & vbCRLF & _
    "becomes unique. In the end, cn and userPrincipalName will be synchronized" & vbCRLF & _
    "with sAMAccountName."
End Sub
