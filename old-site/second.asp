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

<title>.:: Woodland Beach Resort: A Place Where Friends Become Family - About Us: <%= objRS3("Title") %> ::.</title>

<link rel="STYLESHEET" type="text/css" href="css/Global.css">
<link rel="STYLESHEET" type="text/css" href="css/TopNav.css">

<script type="text/javascript" src="js/nav.js"></script>

</head>
<body bgcolor="#F5F4F3" text="#000000" leftmargin="0" topmargin="0" bottommargin="0" rightmargin="0" style="margin:0;">



<table cellpadding="0" cellspacing="0" border="0" width="774" height="100" align="center"> 
	<tr> 
		<td width="11" background="images/leftSide.gif"><img src="images/clear.gif" width="11"></td>
		<td>
			<table cellpadding="0" border="0" cellspacing="0" height="100" width="752" align="center" class="homeBox">
				<tr>
					<td valign="top" height="100">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><a href="index.asp"><img src="images/WBR_Logo.gif" alt="Woodland Beach Resort Logo" width="250" height="100" border="0" title="Woodland Beach Resort Home Page"></a></td>
								<td width="500" height="100" background="images/header.gif" align="right" valign="bottom" class="Time" style="padding-bottom:50px; padding-right:5px;"><span id="tP">&nbsp;</span>&nbsp;<script type="text/javascript" src="js/clock.js"></script>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" height="34"><!--#include file="includes/TopMenu.inc"--></td>
				</tr> 
				<tr>
					<td valign="top" height="274" bgcolor="#FFFFFF">
<!-- MIDDLE CONTENT AND SPECIALS SECTION -->
						<table cellpadding="0" cellspacing="0" border="0" width="750">
							<tr> 
								<td rowspan="2" valign="top" width="110" height="274" background="images/map2.jpg" class="mapText" style="border-right:2px solid #CECAC7;"><img src="images/clear.gif" width="110" height="161"><br><span id="mapB">&raquo;</span> <a href="">Lake Map</a><br><span id="mapB">&raquo;</span> <a href="">Road Map</a><br><span id="mapB">&raquo;</span> <a href="">Resort Map</a></td>
								<td><h1><img src="images/title_about.gif" width="540" height="97" border="0"></h1></td>
							</tr>
							<tr> 
								<!-- <td></td> -->
								<td valign="top" class="body">
								<span class="bodyTitle"><%= objRS3("Title") %></span><br>
								<%= objRS3("Body") %><br><br>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="middle" height="28" background="images/bottomLines.gif" class="cc" align="center">Copyright © 2002 - <%= year(now())%> Woodland Beach Resort&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;www.brainerd.com</td>
				</tr>
			</table>
		</td>
		<td width="11" background="images/rightSide.gif"><img src="images/clear.gif" width="11"></td>
	</tr>
	<tr> 
		<td colspan="3" width="774" height="17"><img src="images/bottom.gif" width="774" height="17"></td>
	</tr>
</table>








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