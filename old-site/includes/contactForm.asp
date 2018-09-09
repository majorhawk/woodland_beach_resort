
<% 
If Request.QueryString("process") = "1" then

dim objRSComments, About, AboutEmail


About = "Questions/Comments from Website"
AboutEmail = "don@woodlandbeachresort.com"

If  instr(Request.Form("Comments"),"http" ) >0 Then

	response.Redirect("contact.asp?PageID=6000&email=1")

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
objMail.BCC = "dan@bigfishdesignsllc.com"
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
					
response.Redirect("contact.asp?PageID=6000&email=1")

End If

End If
 %>

<form action="contact.asp?PageID=6000&process=1" method="post" name="contact" onsubmit="return checkFields(contact)">
<table cellpadding="0" cellspacing="0" border="0" width="450" height="24" style="padding-bottom:1px;">
	<tr>
		<td width="450" height="24" class="formHeader">Contact Us Form</td>
	</tr>
</table>
<table cellpadding="4" cellspacing="0" border="0" width="450" bgcolor="#FBF7EF" style="border: 1px solid #BFB18E;">
	<tr>
		<td width="150px" class="inputBorder1">First Name:</td>
		<td class="inputBorder2"><input type="text" name="FNme" style="width:250px;" class="input1"></td>
	</tr>
	<tr>
		<td class="inputBorder1">Last Name:</td>
		<td class="inputBorder2"><input type="text" name="LNme" style="width:250px;" class="input1"></td>
	</tr>
	<tr>
		<td class="inputBorder1">Email Address:</td>
		<td class="inputBorder2"><input type="text" name="Email" style="width:250px;" class="input1"></td>
	</tr>
	<tr>
		<td colspan="2" class="inputBorder1">Question/Comments:<br>
		<textarea name="Comments" style="width:400px; height:150px; margin-top:3px;" wrap="hard" class="input1"></textarea></td>
		<!-- <td></td> -->
	</tr>
	<tr>
	  <td align="center" colspan="2" class="inputBorderSubmit"><input type="image" src="images/submitBtn.gif" name="submit"><img src="images/BtnSpacer.gif" width="18" height="23" border="0"><a href="<%= Request.QueryString("") %>" onClick="document.contact.reset()"><img src="images/resetBtn.gif" width="74" height="23" border="0"></a></td>
	</tr>
</table>

</form>