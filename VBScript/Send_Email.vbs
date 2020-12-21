Dim objEmail
Dim fromEmail, toEmail, subject, body,strSMTP

' **** Configuration ****

fromEmail = "youremail@domain.com"
toEmail = "recipientemail@domain.com"
subject = "My Subject"
body = "My Mesage..."
strSMTP = "smtpserverhere"

' ***********************

Set objEmail = CreateObject("CDO.Message")

with objEmail
	.From = fromEmail
	.To = toEmail
	.Subject = subject
	.Textbody = body
	.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTP
	.Configuration.Fields.Update

	.Send
end with

set objEmail = nothing
