

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


<%
Dim objRS3, SQLStrng3, ID

SQLStrng3 = "Select * From tblUsers Where ID = " + request.QueryString("ID")



Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLStrng3, objConn, AdOpenKeySet



%>	

<form action="updateuser.asp" method="POST">
<table cellpadding="0" cellspacing="1" width="500" style="border: 1px solid #841D17;" align="center">
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" width="500">
				<tr>
					<td height="23" class="surHeader" align="left" valign="middle" bgcolor="#FFFFFF"><img src="images/header_logo.gif" title="Woodland Beach Resort"></td>
					<td height="23" width="100%" class="surHeader" align="left" background="images/header_fill.gif">Edit User</td>
				</tr>
			</table>
			<table cellpadding="0" cellspacing="0" width="500" class="surBody">
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>First Name:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;"><input type="Text" name="nme" value="<%=objRS3("FNme")%>" style="width:395px;" class="input1"></td>
				</tr>
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>Last Name:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;"><input type="Text" name="lnme" value="<%=objRS3("LNme")%>" style="width:395px;" class="input1"></td>
				</tr>
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>User Name:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;"><input type="Text" name="login" value="<%=objRS3("UserNme")%>" style="width:395px;" class="input1"></td>
				</tr>
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>Password:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;"><input type="text" name="password" value="<%=objRS3("Password")%>" style="width:395px;" class="input1"></td>
				</tr>
				<tr>
					<td height="23" width="115" align="right" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><strong>User Level:</strong></td>
					<td height="23" width="455" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB;">
					<%if cint(objRS3("Admin")) = 1 then %>
					<input type="Radio" name="Rights" value="1" checked>&nbsp;Admin
					<%else %>
					<input type="Radio" name="Rights" value="1">&nbsp;Admin
					<%end if %>
					&nbsp;&nbsp;&nbsp;&nbsp;
					<%if cint(objRS3("Admin")) = 2 then %>
					<input type="Radio" name="Rights" value="2" checked>&nbsp;User 2
					<%else %>
					<input type="Radio" name="Rights" value="2">&nbsp;User 2
					<%end if %>
					&nbsp;&nbsp;&nbsp;&nbsp;
					<%if cint(objRS3("Admin")) = 3 then %>
					<input type="Radio" name="Rights" value="3" checked>&nbsp;User 3
					<%else %>
					<input type="Radio" name="Rights" value="3">&nbsp;User 3
					<%end if %>
					&nbsp;&nbsp;&nbsp;&nbsp;
					<%if cint(objRS3("Admin")) = 4 then %>
					<input type="Radio" name="Rights" value="4" checked>&nbsp;User 4
					<%else %>
					<input type="Radio" name="Rights" value="4">&nbsp;User 4
					<%end if %>
					&nbsp;&nbsp;&nbsp;&nbsp;
					<%if cint(objRS3("Admin")) = 5 then %>
					<input type="Radio" name="Rights" value="5" checked>&nbsp;User 5
					<%else %>
					<input type="Radio" name="Rights" value="5">&nbsp;User 5
					<%end if %>
					</td>
				</tr>
				<tr>
					<td colspan="2" height="23" width="500" align="center" valign="middle" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 0px; border-bottom:1px solid #ABABAB;"><input type="Submit" value="Update User" class="inputSubmit"> <input type="submit" value="Cancel" name="Cancel" class="inputSubmit"> <input type="reset" value="Reset" class="inputSubmit"> <form action="deleteU.asp" method="POST" class="inputSubmit"><input type="Submit" value="Delete User" name="Delete" class="inputSubmit"></td>
			</table>
		</td>
	</tr>
</table>
<input type="hidden" name="ID" value="<%=objRS3("ID")%>">
</form>


<%
objRs3.Close
Set objRS3 = Nothing
%>



<!--#include file="includes/footer.inc"-->

