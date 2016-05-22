const DomainDN = "LDAP://server/" 
const DC = "ou=Users,OU=COMPANY,OU=Customers,dc=user,dc=domain,dc=company"

if AccountNameAlreadyPresent("username") then
	WScript.Echo "ok"
else
	Wscript.Echo "ko"
end if


function AccountNameAlreadyPresent(byVal Account) ' as boolean
	
	LDAPString = "<" & DomainDN & DC & ">;" 
	
	Set objConnection = CreateObject("ADODB.Connection")

	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	 
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	 
	objCommand.CommandText = LDAPString & "(&(objectCategory=User)(samAccountName=" & Account & "));samAccountName;subtree"
	  
	Set objRecordSet = objCommand.Execute
	 
	If objRecordset.RecordCount = 0 Then
	    AccountNameAlreadyPresent = false
	Else
	    AccountNameAlreadyPresent = true
	End If
	 
	objConnection.Close
end function
