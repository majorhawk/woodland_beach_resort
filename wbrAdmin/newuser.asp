

<%@ Language=VBScript %>
<%  Option Explicit%>


<!--#include file="connect.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<!--#include file="security.asp"-->

<%
Dim objRS, SQLString, SQLString2, objRS2
%>

<html>
<head>
	<title>.:: <%= Session("ComNme") %> ::.</title>
</head>
<LINK href="css/main.css" type="text/css" rel="stylesheet">
<meta name="robots" content="noindex, nofollow" />
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0"  bgcolor="#F7F7F8" background="images/bg.gif">


<!--- home page text starts here --->

<!--#include file="includes/header.inc"-->


<!--- home page text starts here --->



<form action="updateuser.asp?new=1" method="POST">
<table cellpadding="0" cellspacing="1" width="500" style="border: 1px solid #841D17;" align="center">
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" width="500">
				<tr>
					<td height="23" class="surHeader" align="left" valign="middle" bgcolor="#FFFFFF"><img src="images/header_logo.gif" title="Woodland Beach Resort"></td>
					<td height="23" width="100%" class="surHeader" align="left" background="images/header_fill.gif">Add A New User</td>
				</tr>
			</table>
			<table cellpadding="0" cellspacing="0" width="500" class="surBody">
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>First Name:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;"><input type="Text" name="nme" style="width:100%;" class="input1"></td>
				</tr>
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>Last Name:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;"><input type="Text" name="lnme" style="width:100%;" class="input1"></td>
				</tr>
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>User Name:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;"><input type="Text" name="login" style="width:100%;" class="input1"></td>
				</tr>
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>Password:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;"><input type="Text" name="password" style="width:100%;" class="input1"></td>
				</tr>
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>User Level:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;"><input type="Radio" name="Rights" value="1">&nbsp;Admin&nbsp;&nbsp;&nbsp;&nbsp;<input type="Radio" name="Rights" value="2">&nbsp;User 2&nbsp;&nbsp;&nbsp;&nbsp;<input type="Radio" name="Rights" value="3">&nbsp;User 3&nbsp;&nbsp;&nbsp;&nbsp;<input type="Radio" name="Rights" value="4">&nbsp;User 4&nbsp;&nbsp;&nbsp;&nbsp;<input type="Radio" name="Rights" value="5">&nbsp;User 5</td>
				</tr>
				<tr>
					<td colspan="2" height="23" width="500" align="center" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><br><input type="Submit" value="Add User" class="inputSubmit"> <input type="submit" value="Cancel" name="Cancel" class="inputSubmit"> <input type="reset" value="Reset" class="inputSubmit"></td>
			</table>
		</td>
	</tr>
</table>
</form>


<% 
objConn.Close
Set objConn = Nothing
%>

<!--#include file="includes/footer.inc"-->


