

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
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0"  bgcolor="#F7F7F8" background="images/bg.gif">


<!--- home page text starts here --->

<!--#include file="includes/header.inc"-->

<!--- home page text starts here --->

<%
Dim objRS3,SQLStrng3,ID

SQLStrng3 = "Select * From tblBoat Where ID = " + request.QueryString("ID")

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

<form action="updateBoats.asp" method="POST">
<table cellpadding="0" cellspacing="1" width="500" style="border: 1px solid #841D17;" align="center">
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" width="500">
				<tr>
					<td height="23" class="surHeader" align="left" valign="middle" bgcolor="#FFFFFF"><img src="images/header_logo.gif" title="Woodland Beach Resort"></td>
					<td height="23" width="100%" class="surHeader" align="left" background="images/header_fill.gif">Edit Boat -  <%=objRS3("BoatNme")%></td>
				</tr>
			</table>
			<table border="0" width="500" cellpadding="5" cellspacing="1" align="center" bgcolor="C0C0C0" class="surBody">
				<tr>
					<td width="100" bgcolor="#ffffff">Edited On:</td>
					<td valign="top" class="body2a" bgcolor="#ffffff"><input type="hidden" name="Edit_On" value="<%=NOW%>"><%=NOW%></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Edited By:</td>
					<td valign="top" class="body2a" bgcolor="#ffffff"><input type="hidden" name="Edit_By" value="<%=Session("FirstName")%>" style="WIDTH: 395px;"><%=Session("FirstName")%></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Viewable:</td> 
					<td width="400" valign="top" bgcolor="#ffffff"><input type="radio" name="viewable" value="1"<%=iif( objRS3("Display") = 1," checked","")%>>&nbsp;Yes&nbsp;&nbsp;&nbsp;<input type="radio" name="viewable" value="0"<%=iif( objRS3("Display") = 0," checked","")%>>&nbsp;No</td>
				</tr>
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">Main Photo:<br></td>
					<td width="400" valign="top" bgcolor="#ffffff"><input type="Text" name="MainPhoto" style="WIDTH: 395px;" value="<%= objRS3("MainPhoto") %>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Boat Name:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="BoatNme" style="WIDTH: 395px;" value="<%= objRS3("BoatNme") %>"></td>
				</tr>
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">Long Description:</td>
					<td width="400" valign="top" bgcolor="#ffffff"><textarea style="WIDTH: 395px; HEIGHT: 150px;" name="LongDes" wrap="off" class="body2a"><%= objRS3("LongDes") %></textarea></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Day Rate:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="BoatRateD" style="WIDTH: 395px;" value="<%= objRS3("BoatRateD") %>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Week Rate:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="BoatRateW" style="WIDTH: 395px;" value="<%= objRS3("BoatRateW") %>"></td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="#ffffff" align="center"><input type="Submit" value="Update Boat" name="update" class="inputSubmit"> <input type="button" value="Cancel" name="Cancel" class="inputSubmit" onClick="parent.location='dates.asp'"> <% If (objRS2("Admin") = "1") Then %><input type="Submit" value="Delete" name="Delete" class="inputSubmit"><% End If %></td>
				</tr>
			</table>
			
			
			
			
		</td>
	</tr>
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


