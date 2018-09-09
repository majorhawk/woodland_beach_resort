

<%@ Language=VBScript %>
<%  Option Explicit%>

<!--#include file="connect.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<!--#include file="security.asp"-->

<% Dim objRS, SQLString, SQLString2, objRS2 %>

<html>
<head>
	<title>.:: <%= Session("ComNme") %> ::.</title>
</head>

<LINK href="css/main.css" type="text/css" rel="stylesheet">
<meta name="robots" content="noindex, nofollow" />
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0"  bgcolor="#010066" background="../images/bg.gif">


<!--- home page text starts here --->

<!--#include file="includes/header.inc"-->

<!--- home page text starts here --->



<form action="updategeneral.asp?new=1" method="POST">
<table border="0" width="500" cellpadding="3" cellspacing="2" align="center" class="body">
	<tr>
		<td colspan="5" bgcolor="#ffffff"><font size="+1">Add New Page:</font></td>
	</tr>
</table>
<table border="0" width="500" cellpadding="5" cellspacing="1" align="center" class="body" bgcolor="C0C0C0">
	<tr>
		<td width="100" bgcolor="#ffffff">Edited On:</td>
		<td width="400" valign="top" bgcolor="#ffffff"><input type="hidden" name="Edit_On" value="<%= NOW%>" size="40"><%= NOW%></td>
	</tr>
	<tr>
		<td width="100" bgcolor="#ffffff">Edited By:</td>
		<td width="400" valign="top" bgcolor="#ffffff"><input type="hidden" name="Edit_By" value="<%=Session("FirstName")%>" size="50"><%=Session("FirstName")%></td>
	</tr>
	<tr>
		<td width="100" bgcolor="#ffffff">Page ID:</td> 
		<td width="400" valign="top" bgcolor="#ffffff"><input type="Text" name="PageID" style="WIDTH: 395px;"></td>
	</tr>
	<tr>
		<td width="110" bgcolor="#ffffff">Searchable:</td> 
		<td width="390" valign="top" bgcolor="#ffffff"><input type="radio" name="Searchable" value="1">&nbsp;Yes&nbsp;&nbsp;&nbsp;<input type="radio" name="Searchable" value="0"<%=iif(objRS3("Searchable") = 0," checked","")%>>&nbsp;No</td>
	</tr>
	<tr>
		<td width="100" bgcolor="#ffffff">Section:</td>
		<td width="400" valign="top" bgcolor="#ffffff">
		<select name="BodySection" style="WIDTH: 395px;">
			<option value="1">Our Company</option>
			<option value="2">Products</option>
			<option value="3">Our Technolgy</option>
			<option value="4">News & Events</option>
			<option value="5">Customer Support</option>
			<option value="6">Contact Us</option>
		</select>
		</td>
	</tr>
	<tr>
		<td width="100" bgcolor="#ffffff">Title:</td>
		<td width="400" valign="top" bgcolor="#ffffff"><input type="Text" name="Title" style="WIDTH: 395px;"></td>
	</tr>
	<tr>
		<td width="100" valign="top" bgcolor="#ffffff">Body:</td>
		<td width="400" valign="top" bgcolor="#ffffff"><textarea style="WIDTH: 395px; HEIGHT: 300px;" name="HTML" wrap="off" class="body2a"></textarea></td>
	</tr>
		<td colspan="2" bgcolor="#ffffff" align="center"><input type="Submit" value="Add Page"> <input type="submit" value="Cancel" name="Cancel"></td>
	</tr>
</table>
</form>


<%
objConn.Close
Set objConn = Nothing
%>

<!--#include file="includes/footer.inc"-->


