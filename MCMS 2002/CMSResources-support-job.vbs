'
' this script looks for two scripts generated by User2Mailer.vbs and Eds2Mailer.vbs
' running periodically on server and sends them to user@company.de
'

const TARGET = "user@company.com"
const SMARTHOST = "server"
const FILE1 = "path\results.txt"

Set objEmail = CreateObject("CDO.Message")

objEmail.From = "user@company.de"
objEmail.To = TARGET
objEmail.Subject = "[INFO] CMS-Ressourcen > 5 MB" 
objEmail.Textbody = "Anbei eine Liste aller CMS-Ressourcen von server mit einer Gr��e von mindestens 5 MBytes."
objEmail.AddAttachment FILE1

objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMARTHOST 
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send
	