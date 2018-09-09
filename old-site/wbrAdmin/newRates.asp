<%@ Language=VBScript %>
<%  Option Explicit%>

<!--#include file="connect.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<!--#include file="security.asp"-->

<%
Dim objRS, SQLString, SQLString2, objRS2, SQLString3, objRS3

SQLString3 = "Select * From tblCabins Order by CabinNme"

Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet




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


<form action="updateRates.asp?new=1" method="POST">
<table cellpadding="0" cellspacing="1" width="500" style="border: 1px solid #841D17;" align="center">
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" width="500">
				<tr>
					<td height="23" class="surHeader" align="left" valign="middle" bgcolor="#FFFFFF"><img src="images/header_logo.gif" title="Woodland Beach Resort"></td>
					<td height="23" width="100%" class="surHeader" align="left" background="images/header_fill.gif">Add New Rates</td>
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
					<% Do While Not ObjRS3.EOF %>
						<option value="<%= objRS3("CabinNme") %>"><%= objRS3("CabinNme") %></option>
					<% objRS3.MoveNext %>
					<% Loop %>
					</select>
					</td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Peak Season:</td>
					<td valign="top" bgcolor="#ffffff">Week:&nbsp;&nbsp;<input type="Text" name="SMRateW" style="WIDTH: 350px;"><hr width="395" size="1">Night:&nbsp;&nbsp;&nbsp;<input type="Text" name="SMRateN" style="WIDTH: 350px;" value="N/A"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Other:</td>
					<td valign="top" bgcolor="#ffffff">Week:&nbsp;&nbsp;<input type="Text" name="SPRateW" style="WIDTH: 350px;"><hr width="395" size="1">Night:&nbsp;&nbsp;&nbsp;<input type="Text" name="SPRateN" style="WIDTH: 350px;"></td>
				</tr>

				<tr>
					<td width="100" bgcolor="#ffffff">Spring & Fall:</td>
					<td valign="top" bgcolor="#ffffff">Week:&nbsp;&nbsp;<input type="Text" name="FRateW" style="WIDTH: 350px;"><hr width="395" size="1">Night:&nbsp;&nbsp;&nbsp;<input type="Text" name="FRateN" style="WIDTH: 350px;"></td>
				</tr>

				<tr>
					<td width="100" bgcolor="#ffffff">Winter:</td>
					<td valign="top" bgcolor="#ffffff">Week:&nbsp;&nbsp;<input type="Text" name="WRateW" style="WIDTH: 350px;"><hr width="395" size="1">Night:&nbsp;&nbsp;&nbsp;<input type="Text" name="WRateN" style="WIDTH: 350px;"></td>
				</tr>


				<tr>
					<td colspan="2" bgcolor="#ffffff" align="center"><input type="Submit" value="Add Rates" class="inputSubmit"> <input type="button" value="Cancel" name="Cancel" class="inputSubmit" onClick="parent.location='rates.asp'"> <input type="reset" value="Reset" class="inputSubmit"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>


<% 
objRS3.Close
Set objRS3 = Nothing

objConn.Close
Set objConn = Nothing
%>

<!--#include file="includes/footer.inc"-->


