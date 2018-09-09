<%@ Language=VBScript %>
<% Option Explicit %>


<!--#include file="includes/DBconnect.asp"-->
<!--#include file="includes/adovbs.inc"-->

<%
Dim SQLString, objRS, SQLString3, objRS3, SQLString4, objRS4, SQLString5, objRS5, mainPhoto

SQLString3 = "Select * FROM tblGeneral where PageID = '2000' and Viewable = '1'"

SQLString = "Select * FROM tblCabins where ID = '" & request.QueryString("ID") & "'"

SQLString4 = "SELECT * FROM tblCabins INNER JOIN tblRates ON tblCabins.CabinNme = tblRates.Cabin WHERE tblCabins.ID = '" & Request.QueryString("ID") & "'"

SQLString5 = "Select * FROM tblDates"


Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

Set objRS4 = Server.CreateObject("ADODB.Recordset")
objRS4.Open SQLString4, objConn, AdOpenKeySet

Set objRS5 = Server.CreateObject("ADODB.Recordset")
objRS5.Open SQLString5, objConn, AdOpenKeySet

%>	


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<title><%= objRS("CabinNme") %> : Woodland Beach Resort - A Place Where Friends Become Family on Bay Lake, MN</title>

<link rel="STYLESHEET" type="text/css" href="css/Global.css">
<link rel="STYLESHEET" type="text/css" href="css/TopNav.css">

<script type="text/javascript" src="js/global.js"></script>
<script type="text/javascript" src="js/nav.js"></script>

</head>
<body bgcolor="#F5F4F3" text="#000000" leftmargin="0" topmargin="0" bottommargin="0" rightmargin="0" style="margin:0;">


<!--#include file="includes/header.asp"-->

								<td valign="top" height="97"><h1 title="Cabins: Find out what your home away from home looks like."><img src="images/title_cabins.gif" width="540" height="97" border="0" alt="Cabins: Find out what your home away from home looks like."></h1></td>
							</tr>
							<tr> 
								<!-- <td></td> -->
								<td valign="top" class="body" height="274">
									<table width="480">
										<tr>
											<td width="206" height="131" valign="top" align="center" class="cabinDes"><img src="images/<%= objRS("MainPhoto") %>" width="206" height="131" border="0" alt="<%= objRS("CabinNme") %>"><br><img src="images/est_<%= objRS("EST") %>.jpg" border="0" alt="Est.<%= objRS("EST") %> on Bay Lake, MN"></td>
											<td valign="top" align="left" style="padding:3px 0 0 5px;"><span class="cabinNme"><%= objRS("CabinNme") %></span><br><span class="cabinDes"><%= objRS("LongDes") %><br><br><strong>Bedrooms: <%= objRS("BedNum") %></strong>&nbsp;&nbsp;|&nbsp;&nbsp;<strong>Bathrooms: <%= objRS("BathNum") %></strong></span></td>
										</tr>
									</table>
									<br>
									<hr width="100%" size="1" color="#6F0000">
									<br>
									<table cellpadding="0" cellspacing="0" border="0" width="480">
										<tr>
											<td height="18" width="18" align="center" class="OCStyle"><a href="javascript:toggleLayer('Rates');toggleLayer2('open1');toggleLayer3('close1');"><img src="images/close_Btn.gif" border="0"title="Open/Close" id="close1" hspace="3" style="visibility:hidden; display:none;"><img src="images/open_Btn.gif" border="0"title="Open/Close" id="open1" hspace="3" style="visibility:visible; display:inline;"></a></td>
											<td height="18" align="left" class="specTitle"><a href="javascript:toggleLayer('Rates');toggleLayer2('open1');toggleLayer3('close1');">Rates</a></td>
										</tr>
									</table>
									<div id="Rates" style="visibility:hidden; margin: 0px 0px 0px 0px; display: none;">
									<% dim season1, season2, season3, season4 %>
									
									<% Do While Not objRS5.EOF %>
									
										<% If objRS5("ID") = 1 then %>
											<% season1 = objRS5("RateDates") %>
										<% End If %>
										<% If objRS5("ID") = 2 then %>
											<% season2 = objRS5("RateDates") %>
										<% End If %>
										<% If objRS5("ID") = 3 then %>
											<% season3 = objRS5("RateDates") %>
										<% End If %>
										<% If objRS5("ID") = 4 then %>
											<% season4 = objRS5("RateDates") %>
										<% End If %>
									
									<% objRS5.MoveNext %>
									<% Loop %>
									<table cellpadding="0" cellspacing="0" border="0" width="480" class="OCStyle2" style="margin-top:3px;">
										<tr>
											<td width="300" height="20" align="left" style="padding-left:30px;" class="OCStyle3"><strong>Dates</strong></td>
											<td width="90" align="left" class="OCStyle3"><strong>Week</strong></td>
											<td width="90" align="left" class="OCStyle3"><strong>Night</strong></td>
										</tr>
										<tr>
											<td height="20" align="left" class="OCStyle4" style="padding-left:30px;"><%= season1 %></td>
											<td align="left" class="OCStyle4"><% If objRS4("SMRateW") <> "N/A" Then %>$<% End If %><%= objRS4("SMRateW") %></td>
											<td align="left" class="OCStyle4"><% If objRS4("SMRateN") <> "N/A" Then %>$<% End If %><%= objRS4("SMRateN") %></td>
										</tr>
										<tr>
											<td height="20" align="left" class="OCStyle4" style="padding-left:30px;" bgcolor="#EAEAEA"><%= season2 %></td>
											<td align="left" class="OCStyle4" bgcolor="#EAEAEA"><% If objRS4("SPRateW") <> "N/A" Then %>$<% End If %><%= objRS4("SPRateW") %></td>
											<td align="left" class="OCStyle4" bgcolor="#EAEAEA"><% If objRS4("SPRateN") <> "N/A" Then %>$<% End If %><%= objRS4("SPRateN") %></td>
										</tr>
										<tr>
											<td height="20" align="left" class="OCStyle4" style="padding-left:30px;"><%= season3 %></td>
											<td align="left" class="OCStyle4"><% If objRS4("FRateW") <> "N/A" Then %>$<% End If %><%= objRS4("FRateW") %></td>
											<td align="left" class="OCStyle4"><% If objRS4("FRateN") <> "N/A" Then %>$<% End If %><%= objRS4("FRateN") %></td>
										</tr>
										<tr>
											<td height="20" align="left" class="OCStyle4" style="padding-left:30px;"><%= season4 %></td>
											<td align="left" class="OCStyle4"><% If objRS4("WRateW") <> "N/A" Then %>$<% End If %><%= objRS4("WRateW") %></td>
											<td align="left" class="OCStyle4"><% If objRS4("WRateN") <> "N/A" Then %>$<% End If %><%= objRS4("WRateN") %></td>
										</tr>
									</table>
									</div>
									<br>
									<table cellpadding="0" cellspacing="0" border="0" width="480">
										<tr>
											<td height="18" width="18" align="center" class="OCStyle"><a href="javascript:toggleLayer('Spec');toggleLayer2('open');toggleLayer3('close');"><img src="images/close_Btn.gif" border="0"title="Open/Close" id="close" hspace="3" style="visibility:hidden; display:none;"><img src="images/open_Btn.gif" border="0"title="Open/Close" id="open" hspace="3" style="visibility:visible; display:inline;"></a></td>
											<td height="18" align="left" class="specTitle"><a href="javascript:toggleLayer('Spec');toggleLayer2('open');toggleLayer3('close');">Floorplan and Photos</a></td>
										</tr>
									</table>
									<div id="Spec" style="visibility:hidden; margin: 0px 0px 0px 0px; display: none;">
										<%= objRS("ShortDes") %>
									</div>
									
								<br><br><br>
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

objRS4.Close
Set objRS4 = Nothing

objRS5.Close
Set objRS5 = Nothing
%>

<%   
objConn.Close
Set objConn = Nothing
%>