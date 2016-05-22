const DomainDN = "LDAP://server/" 
const DC = "ou=Users,OU=COMPANY,OU=Customers,dc=user,dc=domain,dc=company"
LDAPString = "<" & DomainDN & DC & ">;" 

Set objConnection = CreateObject("ADODB.Connection")

objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
 
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection

objCommand.CommandText = LDAPString & "(&(objectclass=user)(objectcategory=person));displayName, c;subtree"
  
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Dim counter 
counter = 0

Do Until objRecordSet.EOF
	if objRecordSet.Fields(1).Value = "DE" then 
		'Wscript.Echo objRecordSet.Fields(1).Value & " / " & objRecordSet.Fields(0).Value ' display Attrs as in strAttrs set
		counter = counter + 1
	end if
		
	objRecordSet.MoveNext
Loop 
objConnection.Close

WScript.Echo "Es gibt " & counter & " Accounts aus Deutschland."