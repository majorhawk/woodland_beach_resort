
<% 
If Request.QueryString("process") = "1" then

dim objRSComments, About, AboutEmail


About = "Questions/Comments from Website"
AboutEmail = "kpmwalleye@aol.com"

If  instr(Request.Form("Comments"),"http" ) >0 Then

	response.Redirect("/contact/?email=1")

else

Set objRSComments= Server.CreateObject ("ADODB.Recordset")
objRSComments.Open "tblContact", objConn, , adLockOptimistic, adCmdTable

objRSComments.AddNew

objRSComments("FNme") = Request.Form("FNme")
objRSComments("LNme") = Request.Form("LNme")
objRSComments("Email") = Request.Form("Email")
objRSComments("Comments") = Request.Form("Comments")
objRSComments("Added_By") = "Website"
objRSComments("Added_On") = NOW
objRSComments("Edit_By") = "Website"
objRSComments("Edit_On") = NOW

objRSComments.Update

objRSComments.Close
Set objRSComments = Nothing


Dim strBody, emailBody, objMail, subjectInfo

subjectInfo = About & " - from woodlandbeachresort.com"
strBody = "<html>" &_
"</head>" &_
"<body>" &_
"<table cellpadding='0' cellspacing='0' border='1' bordercolor='#CCCCCC' width='640' align='left'>" &_
"<tr>" &_
"<td colspan='2' bgcolor='#FFFFFF' valign='top' align='left' width='630' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#FFFFFF; padding:10px 0px 2px 10px;'><img src='http://www.woodlandbeachresort.com/images/contactEmail.gif' title='Nonin Medical, Inc.'></td>" &_
"</tr>" &_
"<tr>" &_
"<td colspan='2' bgcolor='#6F0000' valign='top' align='left' width='630' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#FFFFFF; padding:2px 0px 2px 10px;'><strong>Contact Information</strong></td>" &_
"</tr>" &_
"<tr>" &_
"<td valign='top' align='left' width='160' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'><strong>First Name:</strong></td>" &_
"<td valign='top' align='left' width='460' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'>" & Request.Form("FNme") & "&nbsp;</td>" &_
"</tr>" &_
"<tr>" &_
"<td valign='top' align='left' width='160' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'><strong>Last Name:</strong></td>" &_
"<td valign='top' align='left' width='460' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'>" & Request.Form("LNme") & "&nbsp;</td>" &_
"</tr>" &_
"<tr>" &_
"<td valign='top' align='left' width='160' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'><strong>E-mail Address:</strong></td>" &_
"<td valign='top' align='left' width='460' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'><a href='mailto:" & Request.Form("Email") & "'>" & Request.Form("Email") & "</a>&nbsp;</td>" &_
"</tr>" &_
"<tr>" &_
"<td valign='top' align='left' width='160' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'><strong>Phone Number:</strong></td>" &_
"<td valign='top' align='left' width='460' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'>" & Request.Form("Phone") & "&nbsp;</td>" &_
"</tr>" &_
"<tr>" &_
"<td valign='top' align='left' width='160' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'><strong>Questions/Comments:</strong></td>" &_
"<td valign='top' align='left' width='460' style='font-family: verdana, arial, sans-serif; font-size: 11px; color:#000000; padding:2px 0px 5px 10px;'>" & Request.Form("Comments") & "&nbsp;</td>" &_
"</tr>" &_
"</table>" &_
"<br>" &_
"</body>" &_
"</html>"


Set objMail = Server.CreateObject("CDO.message")

objMail.From = Request.Form("Email")
objMail.To = AboutEmail
'objMail.CC = emailList2
objMail.BCC = "dan@thehochstaetters.com"
objMail.Subject = subjectInfo
objMail.HTMLBody = strBody

objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'Name or IP of remote SMTP server
objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="localhost"

'Server port
 objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25

objMail.Configuration.Fields.update

objMail.Send

Set objMail = Nothing
					
response.Redirect("/contact/?email=1")

End If

End If
 %>
<form action="/contact/?process=1" method="post" name="contact" onsubmit="return checkFields(contact)">
<table cellpadding="0" cellspacing="0" border="0" width="276" style="font-family:Arial, Helvetica, sans-serif; ">
	<tr>
		<td valign="top" width="276" height="26" style="background-image:url(/images/form-first.jpg); background-repeat:no-repeat; padding:4px 0 10px 72px;">
		<input type="text" name="FNme" style="width:194px; background-image:url(/images/form-pattern.jpg); background-color:none; border:0px; color:#000000;"></td>
	</tr>
	<tr>
		<td valign="top" width="276" height="26" style="background-image:url(/images/form-last.jpg); background-repeat:no-repeat; padding:4px 0 10px 72px;">
		<input type="text" name="LNme" style="width:194px; background-image:url(/images/form-pattern.jpg); background-color:none; border:0px; color:#000000;"></td>
	</tr>
	<tr>
		<td valign="top" width="276" height="26" style="background-image:url(/images/form-email.jpg); background-repeat:no-repeat; padding:4px 0 10px 92px;">
		<input type="text" name="Email" style="width:174px; background-image:url(/images/form-pattern.jpg); background-color:none; border:0px; color:#000000;"></td>
	</tr>
	<tr>
		<td valign="top" width="276" height="26" style="background-image:url(/images/form-phone.jpg); background-repeat:no-repeat; padding:4px 0 10px 50px;">
		<input type="text" name="Phone" style="width:174px; background-image:url(/images/form-pattern.jpg); background-color:none; border:0px; color:#000000;"></td>
	</tr>
	<tr>
		<td width="276" height="123" style="background-image:url(/images/form-comments.jpg); background-repeat:no-repeat; padding:12px 4px 4px 8px;"><textarea name="Comments" style="width:255px; height:88px; background-image:url(/images/form-pattern2.jpg); background-color:none; border:0px; color:#000000;" wrap="hard"></textarea></td> 
	</tr>
	<tr>
	  <td align="center" style="padding-top:5px;"><input type="image" src="/images/submitBtn.gif" name="submit"></td>
	</tr>
</table>
</form>