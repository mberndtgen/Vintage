Dim dso 'As IADsOpenDSObject
Dim Container 'As IADsContainer
Dim User 'As IADs
Dim ClassArray 'As Variant

Const ADS_UF_NORMAL_ACCOUNT = 512 

set Container = GetObject("LDAP://server/ou=Users,OU=Projects,dc=corp,dc=company")

Set User = Container.Create("user", "cn=Testuser-script")

User.Put "sAMAccountName", "Testuser-script"   ' e.g. joes
User.Put "userPrincipalName", "Testuser-script@company.xyz" ' e.g. joes@rallencorp.com
User.Put "givenName", "Test"   ' e.g. Joe
User.Put "sn", "User"           ' e.g. Smith
User.Put "displayName", "Test User" ' e.g. Joe Smith
User.Put "userAccountControl", ADS_UF_NORMAL_ACCOUNT
User.SetInfo
User.SetPassword("password")
User.AccountDisabled = FALSE
User.SetInfo
