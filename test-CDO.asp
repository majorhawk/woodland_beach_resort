<%
Set objCDOSYSMail = Server.CreateObject("CDO.Message") 
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 
 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1" 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 
objCDOSYSCon.Fields.Update 
 
Set objCDOSYSMail.Configuration = objCDOSYSCon 
objCDOSYSMail.From = "test@woodlandbeachresort.com"
objCDOSYSMail.To = "kevin@woodlandbeachresort.com"
objCDOSYSMail.Subject = "Test ASP-CDO from woodlandbeachresort.com"
objCDOSYSMail.HTMLBody = "This test mail was sended with use CDO-function from ASP-script."
objCDOSYSMail.Send 
 
Set objCDOSYSMail = Nothing 
Set objCDOSYSCon = Nothing 
%>