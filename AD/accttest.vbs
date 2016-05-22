Set objUser = GetObject ("LDAP://server/cn=user_id,ou=Users,dc=corp,dc=company")
 
If objUser.AccountDisabled = FALSE Then
      WScript.Echo "The account is enabled."
Else
      WScript.Echo "The account is disabled."
End If
	