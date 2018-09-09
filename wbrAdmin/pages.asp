
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


SQLString = "Select * FROM tblGeneral" + " Order by PageID"



Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet



%>	

<table cellpadding="0" cellspacing="0" width="500" bgcolor="#FFFFFF" align="center">
	<tr>
		<td height="20" width="500" align="right" valign="middle" bgcolor="#FFFFFF" class="surTitles2"><a href="newPage.asp"><strong>Add a New Page &raquo;</strong></a></td>
	</tr>
</table>
<table cellpadding="0" cellspacing="1" width="500" style="border: 1px solid #841D17;" align="center">
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" width="500">
				<tr>
					<td height="23" class="surHeader" align="left" valign="middle" bgcolor="#FFFFFF"><img src="images/header_logo.gif" title="Woodland Beach Resort"></td>
					<td height="23" width="100%" class="surHeader" align="left" background="images/header_fill.gif">Edit/Add Page</td>
				</tr>
			</table>
			<table cellpadding="0" cellspacing="0" width="500" class="surBody">
				<tr>
					<td height="23" width="80" align="left" valign="top" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB; border-top:1px solid #DEDDDD; border-right:1px solid #DEDDDD;"><strong>PageID</strong></td>
					<td height="23" width="190" valign="top" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB; border-top:1px solid #DEDDDD; border-right:1px solid #DEDDDD;"><strong>Page Title</strong></td>
					<td height="23" width="110" valign="top" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB; border-top:1px solid #DEDDDD; border-right:1px solid #DEDDDD;"><strong>Edit By/On</strong></td>
					<td height="23" width="50" valign="top" align="center" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB; border-top:1px solid #DEDDDD;"><strong>Edit</strong></td>
				</tr>
				<% Do While Not ObjRS.EOF %>
				<tr>
					<td height="23" width="80" align="left" valign="top" bgcolor="#FFFFFF" class="surBody" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB; border-top:1px solid #DEDDDD; border-right:1px solid #DEDDDD;"><%= objRS("PageID") %></td>
					<td height="23" width="190" valign="top" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB; border-top:1px solid #DEDDDD; border-right:1px solid #DEDDDD;"><%= objRS("Title") %></td>
					<td height="23" width="110" valign="top" align="left" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB; border-top:1px solid #DEDDDD; border-right:1px solid #DEDDDD;"><%= objRS("Edit_On") %><br><%= objRS("Edit_By") %></td>
					<td height="23" width="50" valign="top" align="center" bgcolor="#FFFFFF" style="padding:3px 10px 3px 10px; border-bottom:1px solid #ABABAB; border-top:1px solid #DEDDDD;"><a href="editPage.asp?ID=<%= objRS("ID") %>">Edit&nbsp;&nbsp;<img src="images/edit.gif" alt="" border="0"></a></td>
				</tr>
				<%
				objRS.MoveNext
				Loop
				
				objRs.Close
				Set objRS = Nothing
				
				objConn.Close
				Set objConn = Nothing
				%>
			</table>
		</td>
	</tr>
</table>



<!--#include file="includes/footer.inc"-->
