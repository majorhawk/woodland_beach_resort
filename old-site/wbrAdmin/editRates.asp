

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
Dim objRS3, SQLStrng3, objRS4, SQLString4, ID

SQLStrng3 = "Select * From tblRates Where ID = " + request.QueryString("ID")

Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLStrng3, objConn, AdOpenKeySet


SQLString4 = "Select * From tblCabins Order by CabinNme"

Set objRS4 = Server.CreateObject("ADODB.Recordset")
objRS4.Open SQLString4, objConn, AdOpenKeySet



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

<form action="updateRates.asp" method="POST">
<table cellpadding="0" cellspacing="1" width="500" style="border: 1px solid #841D17;" align="center">
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" width="500">
				<tr>
					<td height="23" class="surHeader" align="left" valign="middle" bgcolor="#FFFFFF"><img src="images/header_logo.gif" title="Woodland Beach Resort"></td>
					<td height="23" width="100%" class="surHeader" align="left" background="images/header_fill.gif">Edit Rates for -  <%=objRS3("Cabin")%></td>
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
					<td width="100" bgcolor="#ffffff">Cabin Name:</td>
					<td valign="top" bgcolor="#ffffff">
					<select name="Cabin" style="WIDTH: 395px;">
						<option value="0">Select a Cabin</option>
					<% Do While Not ObjRS4.EOF %>
						<option value="<%= objRS4("CabinNme") %>"<%=iif( objRS3("Cabin") = objRS4("CabinNme")," selected","")%>><%= objRS4("CabinNme") %></option>
					<% objRS4.MoveNext %>
					<% Loop %>
					</select>
					</td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Peak Season:</td>
					<td valign="top" bgcolor="#ffffff">Week:&nbsp;&nbsp;<input type="Text" name="SMRateW" style="WIDTH: 350px;" value="<%= objRS3("SMRateW") %>"><hr width="395" size="1">Night:&nbsp;&nbsp;&nbsp;<input type="Text" name="SMRateN" style="WIDTH: 350px;" value="<%= objRS3("SMRateN") %>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Other:</td>
					<td valign="top" bgcolor="#ffffff">Week:&nbsp;&nbsp;<input type="Text" name="SPRateW" style="WIDTH: 350px;" value="<%= objRS3("SPRateW") %>"><hr width="395" size="1">Night:&nbsp;&nbsp;&nbsp;<input type="Text" name="SPRateN" style="WIDTH: 350px;" value="<%= objRS3("SPRateN") %>"></td>
				</tr>

				<tr>
					<td width="100" bgcolor="#ffffff">Spring & Fall:</td>
					<td valign="top" bgcolor="#ffffff">Week:&nbsp;&nbsp;<input type="Text" name="FRateW" style="WIDTH: 350px;" value="<%= objRS3("FRateW") %>"><hr width="395" size="1">Night:&nbsp;&nbsp;&nbsp;<input type="Text" name="FRateN" style="WIDTH: 350px;" value="<%= objRS3("FRateN") %>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Winter:</td>
					<td valign="top" bgcolor="#ffffff">Week:&nbsp;&nbsp;<input type="Text" name="WRateW" style="WIDTH: 350px;" value="<%= objRS3("WRateW") %>"><hr width="395" size="1">Night:&nbsp;&nbsp;&nbsp;<input type="Text" name="WRateN" style="WIDTH: 350px;" value="<%= objRS3("WRateN") %>"></td>
				</tr>

				<tr>
					<td colspan="2" bgcolor="#ffffff" align="center"><input type="Submit" value="Update Rates" name="update" class="inputSubmit"> <input type="button" value="Cancel" name="Cancel" class="inputSubmit" onClick="parent.location='rates.asp'"> <% If (objRS2("Admin") = "1") Then %><input type="Submit" value="Delete" name="Delete" class="inputSubmit"><% End If %></td>
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

objRs4.Close
Set objRS4 = Nothing

objConn.Close
Set objConn = Nothing
%>

<!--#include file="includes/footer.inc"-->


