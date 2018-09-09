<%@ Language=VBScript %>
<% Option Explicit %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>

<!--#include file="includes/DBconnect.asp"-->
<!--#include file="includes/adovbs.inc"-->
<!--#include file="includes/iif.asp"-->

<%
Dim SQLString, objRS
SQLString = "Select * FROM tblGeneral where PageID = '9001'"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

%>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Woodland Beach Resort Email</title>

	
</head>

<body>
<table cellpadding="0" cellspacing="0" border="0" height="473">
	<tr>
		<td rowspan="3" valign="top"><a href="http://www.woodlandbeachresort.com"><img src="http://www.woodlandbeachresort.com/images/email-wbr-main.jpg" height="473" width="401" border="0" alt=""></a></td>
		<td width="311" valign="top"><img src="http://www.woodlandbeachresort.com/images/email-wbr-top.jpg" height="90" width="311" border="0" alt=""></td>
		<td valign="top" rowspan="3"><img src="http://www.woodlandbeachresort.com/images/email-wbr-side.jpg" height="473" width="51" border="0" alt=""></td>
	</tr>
	<tr>
		<td valign="top" height="332" width="311" bgcolor="#f6f3ec" style="font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#333333;"><div style="margin-top:8px; margin-bottom:8px; font-size:16px; color:#7a2321; font-weight:bold;"><%= objRS("Title") %></div>
<%= objRS("Body") %></td>
	</tr>
	<tr>
		<td valign="top" height="51" width="311"><img src="http://www.woodlandbeachresort.com/images/email-wbr-bottom.jpg" height="51" width="311" border="0" alt=""></td>
	</tr>
</table>
<br><br><br>


</body>
</html>

