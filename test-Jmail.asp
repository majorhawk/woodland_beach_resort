<%
Text="This test mail was sended with use JMail-function from ASP-script."
 
'send email
 
Set JMail = Server.CreateObject("JMail.SMTPMail") 
JMail.ServerAddress = "127.0.0.1"
JMail.Sender = "test@woodlandbeachresort.com"
JMail.AddRecipient "kevin@woodlandbeachresort.com"
JMail.Subject = "Test ASP-JMail from woodlandbeachresort.com"
JMail.ContentType = "text/html"
 
JMail.Body = Text
JMail.Priority = 1
 
JMail.Execute
 
Response.Write Text
%>