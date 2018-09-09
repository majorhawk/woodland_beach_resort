

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

<%
Dim objRS3,SQLStrng3,ID

SQLStrng3 = "Select * From tblPage Where ID = " + request.QueryString("ID")

Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLStrng3, objConn, AdOpenKeySet


'************************************
'* BELOW IS THE INLINE IF STATEMENT *
'************************************
Function iif(trueFalse, a, b)
	If cbool(trueFalse) Then
		iif = a
	Else
		iif = b
	End If
End Function

%>	

<form action="updategeneral.asp" method="POST">
<table border="0" width="500" cellpadding="3" cellspacing="2" align="center" class="body">
	<tr>
		<td colspan="5" bgcolor="#ffffff"><font size="+1">Edit Page:</font></td>
	</tr>
</table>
<table border="0" width="500" cellpadding="5" cellspacing="1" align="center" class="body" bgcolor="C0C0C0">
	<tr>
		<td width="100" bgcolor="#ffffff">Edited On:</td>
		<td valign="top" class="body2a" bgcolor="#ffffff"><input type="hidden" name="Edit_On" value="<%=NOW%>"><%=NOW%></td>
	</tr>
	<tr>
		<td width="100" bgcolor="#ffffff">Edited By:</td>
		<td valign="top" class="body2a" bgcolor="#ffffff"><input type="hidden" name="Edit_By" value="<%=Session("FirstName")%>" style="WIDTH: 395px;"><%=Session("FirstName")%></td>
	</tr>
	<tr>
		<td width="100" bgcolor="#ffffff">Page ID:</td> 
		<td width="400" valign="top" bgcolor="#ffffff"><input type="Text" name="PageID" style="WIDTH: 395px;" value="<%=objRS3("PageID")%>"></td>
	</tr>
	<tr>
		<td width="110" bgcolor="#ffffff">Searchable:</td> 
		<td width="390" valign="top" bgcolor="#ffffff"><input type="radio" name="Searchable" value="1" <%=iif(objRS3("Searchable") = 1," checked","")%>>&nbsp;Yes&nbsp;&nbsp;&nbsp;<input type="radio" name="Searchable" value="0"<%=iif(objRS3("Searchable") = 0," checked","")%>>&nbsp;No</td>
	</tr>
	<tr>
		<td width="100" bgcolor="#ffffff">Section:</td>
		<td width="400" valign="top" bgcolor="#ffffff">
		<select name="BodySection" style="WIDTH: 395px;">
			<option value="1"<%=iif(objRS3("BodySection") = 1," selected","")%>>Our Company</option>
			<option value="2"<%=iif(objRS3("BodySection") = 2," selected","")%>>Products</option>
			<option value="3"<%=iif(objRS3("BodySection") = 3," selected","")%>>Our Technolgy</option>
			<option value="4"<%=iif(objRS3("BodySection") = 4," selected","")%>>News & Events</option>
			<option value="5"<%=iif(objRS3("BodySection") = 5," selected","")%>>Customer Support</option>
			<option value="6"<%=iif(objRS3("BodySection") = 6," selected","")%>>Contact Us</option>
		</select>
		</td>
	</tr>
	<tr>
		<td width="100" bgcolor="#ffffff">Title:</td>
		<td valign="top" bgcolor="#ffffff"><input type="Text" name="Title" value="<%=objRS3("Title")%>" style="WIDTH: 395px;"></td>
	</tr>
	<tr>
		<td width="100" valign="top" bgcolor="#ffffff">Body:</td>
		<td width="400" valign="top" bgcolor="#ffffff"><textarea style="WIDTH: 395px; HEIGHT: 300px;" name="HTML" wrap="off" class="body2a"><%=objRS3("HTML")%></textarea></td>
	</tr><tr>
		<td colspan="2" bgcolor="#ffffff" align="center"><input type="Submit" value="Update Page" name="update"> <input type="submit" value="Cancel" name="Cancel"> <input type="Submit" value="Delete" name="Delete"></td>
</table>
<input type="Hidden" name="ID" value="<%=objRS3("ID")%>">
</form>


<% 
objRs3.Close
Set objRS3 = Nothing

objConn.Close
Set objConn = Nothing
%>

<!--#include file="includes/footer.inc"-->


