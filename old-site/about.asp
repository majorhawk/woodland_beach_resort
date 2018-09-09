<%@ Language=VBScript %>
<% Option Explicit %>


<!--#include file="includes/DBconnect.asp"-->
<!--#include file="includes/adovbs.inc"-->

<%
Dim SQLString3, objRS3, mainPhoto

SQLString3 = "Select * FROM tblGeneral where PageID = '" & request.QueryString("PageID") & "' and Viewable = '1'"

Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet



%>	


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<title><%= objRS3("Title") %> : About Us : Woodland Beach Resort - A Place Where Friends Become Family on Bay Lake, MN</title>

<link rel="STYLESHEET" type="text/css" href="css/Global.css">
<link rel="STYLESHEET" type="text/css" href="css/TopNav.css">
<link rel="stylesheet" media="all" type="text/css" href="css/thumb.css" >

<script type="text/javascript" src="js/nav.js"></script>

</head>
<body bgcolor="#F5F4F3" text="#000000" leftmargin="0" topmargin="0" bottommargin="0" rightmargin="0" style="margin:0;">


<!--#include file="includes/header.asp"-->

								<td valign="top" height="97"><h1 title="About Us: Find out a little more about us, the lake and the resort."><img src="images/title_about.gif" width="540" height="97" border="0" title="About Us: Find out a little more about us, the lake and the resort."></h1></td>
							</tr>
							<tr> 
								<!-- <td></td> -->
								<td valign="top" class="body" height="274"><span class="bodyTitle"><%= objRS3("Title") %></span><br><%= objRS3("Body") %><br><br></td>
							</tr>
						</table>
						
<!--#include file="includes/footer.asp"-->








</body>
</html>
<% 
objRs3.Close
Set objRS3 = Nothing
%>

<%   
objConn.Close
Set objConn = Nothing
%>