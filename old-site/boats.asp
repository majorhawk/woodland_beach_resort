<%@ Language=VBScript %>
<% Option Explicit %>


<!--#include file="includes/DBconnect.asp"-->
<!--#include file="includes/adovbs.inc"-->

<%
Dim SQLString, objRS, SQLString3, objRS3, mainPhoto

SQLString3 = "Select * FROM tblGeneral where PageID = '3000' and Viewable = '1'"

SQLString = "Select * FROM tblBoat where Display = '1' order by BoatNme"




Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet





%>	


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<title><%= objRS3("Title") %> : Woodland Beach Resort - A Place Where Friends Become Family on Bay Lake, MN</title>

<link rel="STYLESHEET" type="text/css" href="css/Global.css">
<link rel="STYLESHEET" type="text/css" href="css/TopNav.css">

<script type="text/javascript" src="js/nav.js"></script>

</head>
<body bgcolor="#F5F4F3" text="#000000" leftmargin="0" topmargin="0" bottommargin="0" rightmargin="0" style="margin:0;">


<!--#include file="includes/header.asp"-->

								<td valign="top" height="97"><h1 title="Boats: Find out boat fits your need for speed."><img src="images/title_boats.gif" width="540" height="97" border="0" alt="Boats: Find out boat fits your need for speed."></h1></td>
							</tr>
							<tr> 
								<!-- <td></td> -->
								<td valign="top" class="body" height="274">
								<%= objRS3("Body") %>
								<br>
								<% Do While Not ObjRS.EOF %>
									<table width="480">
										<tr>
											<td width="206" height="131" valign="top" align="center" class="cabinDes"><img src="images/<%= objRS("MainPhoto") %>" width="206" height="131" border="0"></td>
											<td valign="top" align="left" style="padding:3px 0 0 5px;"><span class="cabinNme"><%= objRS("BoatNme") %></span><br><span class="cabinDes"><%= objRS("LongDes") %></span><br><br><span id="mapB">&raquo;</span> <strong>Day Rate: $<%= objRS("BoatRateD") %></strong><br><br><span id="mapB">&raquo;</span> <strong>Week Rate: $<%= objRS("BoatRateW") %></strong></span></td>
										</tr>
									</table>
									<br>
									<hr width="100%" size="1" color="#6F0000">
									<br>
								<% objRS.MoveNext %>
								<% Loop %>
								
								</td>
							</tr>
						</table>
						
<!--#include file="includes/footer.asp"-->








</body>
</html>
<% 
objRS.Close
Set objRS = Nothing

objRS3.Close
Set objRS3 = Nothing

%>

<%   
objConn.Close
Set objConn = Nothing
%>