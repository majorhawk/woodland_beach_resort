<%@ Language=VBScript %>
<% Option Explicit %>


<!--#include file="includes/DBconnect.asp"-->
<!--#include file="includes/adovbs.inc"-->

<%
Dim SQLString, objRS, SQLString3, objRS3, SQLString4, objRS4, SQLString5, objRS5, mainPhoto

SQLString3 = "Select * FROM tblGeneral where PageID = '4000' and Viewable = '1'"

SQLString4 = "SELECT tblCabins.ID, tblCabins.CabinNme, tblCabins.EST, tblCabins.CabinType, tblCabins.CabinOrder, tblCabins.ShortDes, tblCabins.LongDes, tblCabins.BathNum, tblCabins.BedNum, tblCabins.MainPhoto, tblCabins.Display, tblRates.Cabin, tblRates.SPRateW, tblRates.SPRateN, tblRates.SMRateW, tblRates.SMRateN, tblRates.FRateW, tblRates.FRateN, tblRates.WRateW, tblRates.WRateN FROM tblCabins INNER JOIN tblRates ON tblCabins.CabinNme = tblRates.Cabin Order By CabinType, CabinNme"

SQLString5 = "Select * FROM tblDates"


Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet


Set objRS4 = Server.CreateObject("ADODB.Recordset")
objRS4.Open SQLString4, objConn, AdOpenKeySet

Set objRS5 = Server.CreateObject("ADODB.Recordset")
objRS5.Open SQLString5, objConn, AdOpenKeySet

%>	


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<title><%= objRS3("Title") %> : Woodland Beach Resort - A Place Where Friends Become Family on Bay Lake, MN</title>

<link rel="STYLESHEET" type="text/css" href="css/Global.css">
<link rel="STYLESHEET" type="text/css" href="css/TopNav.css">

<script type="text/javascript" src="js/global.js"></script>
<script type="text/javascript" src="js/nav.js"></script>

</head>
<body bgcolor="#F5F4F3" text="#000000" leftmargin="0" topmargin="0" bottommargin="0" rightmargin="0" style="margin:0;">


<!--#include file="includes/header.asp"-->

								<td valign="top" height="97"><h1 title="Rates: When it comes to price we have something for everyone."><img src="images/title_rates.gif" width="540" height="97" border="0" alt="Rates: When it comes to price we have something for everyone."></h1></td>
							</tr>
							<tr> 
								<!-- <td></td> -->
								<td valign="top" class="body" height="274">

									<% dim season1, season2, season3, season4, season1a, season2a, season3a, season4a %>
									
									<% Do While Not objRS5.EOF %>
									
										<% If objRS5("ID") = 1 then %>
											<% season1 = ".::&nbsp;" & objRS5("RateDes") & "&nbsp;::.&nbsp;&nbsp;" %>
											<% season1a = objRS5("RateDates") %>
										<% End If %>
										<% If objRS5("ID") = 2 then %>
											<% season2 = ".::&nbsp;" & objRS5("RateDes") & "&nbsp;::.&nbsp;&nbsp;" %>
											<% season2a = objRS5("RateDates") %>
										<% End If %>
										<% If objRS5("ID") = 3 then %>
											<% season3 = ".::&nbsp;" & objRS5("RateDes") & "&nbsp;::.&nbsp;&nbsp;" %>
											<% season3a = objRS5("RateDates") %>
										<% End If %>
										<% If objRS5("ID") = 4 then %>
											<% season4 = ".::&nbsp;" & objRS5("RateDes") & "&nbsp;::.&nbsp;&nbsp;" %>
											<% season4a = objRS5("RateDates") %>
										<% End If %>
									
									<% objRS5.MoveNext %>
									<% Loop %>
									<table cellpadding="0" cellspacing="0" border="0" width="480" class="OCStyle2" style="margin-top:3px;">
										<tr>
											<td height="18" width="18" align="center" class="OCStyle3"><a href="javascript:toggleLayer('Rate1');toggleLayer2('open');toggleLayer3('close');"><img src="images/close_Btn.gif" border="0"title="Open/Close" id="close" hspace="3" style="visibility:visible; display:inline;"><img src="images/open_Btn.gif" border="0"title="Open/Close" id="open" hspace="3" style="visibility:hidden; display:none;"></a></td>
											<td width="132" height="20" align="left" style="padding-left:10px;" class="OCStyle3"><a href="javascript:toggleLayer('Rate1');toggleLayer2('open');toggleLayer3('close');"><strong><%= season1 %></strong></a></td>
											<td width="330" height="20" align="right" style="padding-right:10px;" class="OCStyle3"><%= season1a %></td>
										</tr>
									</table>
									<div id="Rate1" style="visibility:visible; margin: 0px 0px 0px 0px; display: block;">
									<Table cellpadding="0" cellspacing="0" border="0" width="480" class="OCStyle2">
										<tr>
											<td width="260" height="20" align="left" class="OCStyle4" style="padding-left:10px;" bgcolor="#F0EFE2"><strong>Cabin Name</strong></td>
											<td width="110" align="right" class="OCStyle4" bgcolor="#F0EFE2" style="padding-right:10px;"><strong>Week</strong></td>
											<td width="110" align="right" class="OCStyle4" bgcolor="#F0EFE2" style="padding-right:10px;"><strong>Night</strong></td>
										</tr>
									<% dim bgcolor %>
									<% bgcolor = "EAEAEA" %>
									<% Do While Not objRS4.EOF %>
										<% If bgcolor = "FFFFFF" Then %>
											<% bgcolor = "EAEAEA" %>
										<% Else %>
											<% bgcolor = "FFFFFF" %>
										<% End If %>
										<tr>
											<td valign="top" width="260" height="20" align="left" class="OCStyle4" style="padding:3px 0 3px 10px;" bgcolor="#<%= bgcolor %>"><a href="cabin.asp?ID=<%= objRS4("ID") %>"><strong><%= objRS4("CabinNme")%></strong>&nbsp;&nbsp;<span class="cc2">est. <%=objRS4("EST") %></span></a><br><span class="cc3">Bedrooms: <%= objRS4("BedNum") %>&nbsp;&nbsp;|&nbsp;&nbsp;Bathrooms: <%= objRS4("BathNum") %></span></td>
											<td valign="top" width="110" align="right" class="OCStyle4" style="padding:3px 10px 3px 0;" bgcolor="#<%= bgcolor %>"><% If objRS4("SMRateW") <> "N/A" Then %>$<% End If %><%= objRS4("SMRateW") %></td>
											<td valign="top" width="110" align="right" class="OCStyle4" style="padding:3px 10px 3px 0;" bgcolor="#<%= bgcolor %>"><% If objRS4("SMRateN") <> "N/A" Then %>$<% End If %><%= objRS4("SMRateN") %></td>
										</tr>
									<% objRS4.MoveNext %>
									<% Loop %>
									<% objRS4.MoveFirst %>
									</Table>
									<br>
									</div>
									<br>
									<table cellpadding="0" cellspacing="0" border="0" width="480" class="OCStyle2" style="margin-top:3px;">
										<tr>
											<td height="18" width="18" align="center" class="OCStyle3"><a href="javascript:toggleLayer('Rate2');toggleLayer2('open2');toggleLayer3('close2');"><img src="images/close_Btn.gif" border="0"title="Open/Close" id="close2" hspace="3" style="visibility:hidden; display:none;"><img src="images/open_Btn.gif" border="0"title="Open/Close" id="open2" hspace="3" style="visibility:visible; display:inline;"></a></td>
											<td width="132" height="20" align="left" style="padding-left:10px;" class="OCStyle3"><a href="javascript:toggleLayer('Rate2');toggleLayer2('open2');toggleLayer3('close2');"><strong><%= season2 %></strong></a></td>
											<td width="330" height="20" align="right" style="padding-right:10px;" class="OCStyle3"><%= season2a %></td>
										</tr>
									</table>
									<div id="Rate2" style="visibility:hidden; margin: 0px 0px 0px 0px; display: none;">
									<Table cellpadding="0" cellspacing="0" border="0" width="480" class="OCStyle2">
										<tr>
											<td width="260" height="20" align="left" class="OCStyle4" style="padding-left:10px;" bgcolor="#F0EFE2"><strong>Cabin Name</strong></td>
											<td width="110" align="right" class="OCStyle4" bgcolor="#F0EFE2" style="padding-right:10px;"><strong>Week</strong></td>
											<td width="110" align="right" class="OCStyle4" bgcolor="#F0EFE2" style="padding-right:10px;"><strong>Night</strong></td>
										</tr>
									<% bgcolor = "EAEAEA" %>
									<% Do While Not objRS4.EOF %>
										<% If bgcolor = "FFFFFF" Then %>
											<% bgcolor = "EAEAEA" %>
										<% Else %>
											<% bgcolor = "FFFFFF" %>
										<% End If %>
										<tr>
											<td valign="top" width="260" height="20" align="left" class="OCStyle4" style="padding:3px 0 3px 10px;" bgcolor="#<%= bgcolor %>"><a href="cabin.asp?ID=<%= objRS4("ID") %>"><strong><%= objRS4("CabinNme")%></strong>&nbsp;&nbsp;<span class="cc2">est. <%=objRS4("EST") %></span></a><br><span class="cc3">Bedrooms: <%= objRS4("BedNum") %>&nbsp;&nbsp;|&nbsp;&nbsp;Bathrooms: <%= objRS4("BathNum") %></span></td>
											<td valign="top" width="110" align="right" class="OCStyle4" style="padding:3px 10px 3px 0;" bgcolor="#<%= bgcolor %>"><% If objRS4("SPRateW") <> "N/A" Then %>$<% End If %><%= objRS4("SPRateW") %></td>
											<td valign="top" width="110" align="right" class="OCStyle4" style="padding:3px 10px 3px 0;" bgcolor="#<%= bgcolor %>"><% If objRS4("SPRateN") <> "N/A" Then %>$<% End If %><%= objRS4("SPRateN") %></td>
										</tr>
									<% objRS4.MoveNext %>
									<% Loop %>
									<% objRS4.MoveFirst %>
									</Table>
									<br>
									</div>
									<br>
									<table cellpadding="0" cellspacing="0" border="0" width="480" class="OCStyle2" style="margin-top:3px;">
										<tr>
											<td height="18" width="18" align="center" class="OCStyle3"><a href="javascript:toggleLayer('Rate3');toggleLayer2('open3');toggleLayer3('close3');"><img src="images/close_Btn.gif" border="0"title="Open/Close" id="close3" hspace="3" style="visibility:hidden; display:none;"><img src="images/open_Btn.gif" border="0"title="Open/Close" id="open3" hspace="3" style="visibility:visible; display:inline;"></a></td>
											<td width="132" height="20" align="left" style="padding-left:10px;" class="OCStyle3"><a href="javascript:toggleLayer('Rate3');toggleLayer2('open3');toggleLayer3('close3');"><strong><%= season3 %></strong></a></td>
											<td width="330" height="20" align="right" style="padding-right:10px;" class="OCStyle3"><%= season3a %></td>
										</tr>
									</table>
									<div id="Rate3" style="visibility:hidden; margin: 0px 0px 0px 0px; display: none;">
									<Table cellpadding="0" cellspacing="0" border="0" width="480" class="OCStyle2">
										<tr>
											<td width="260" height="20" align="left" class="OCStyle4" style="padding-left:10px;" bgcolor="#F0EFE2"><strong>Cabin Name</strong></td>
											<td width="110" align="right" class="OCStyle4" bgcolor="#F0EFE2" style="padding-right:10px;"><strong>Week</strong></td>
											<td width="110" align="right" class="OCStyle4" bgcolor="#F0EFE2" style="padding-right:10px;"><strong>Night</strong></td>
										</tr>
									<% bgcolor = "EAEAEA" %>
									<% Do While Not objRS4.EOF %>
										<% If bgcolor = "FFFFFF" Then %>
											<% bgcolor = "EAEAEA" %>
										<% Else %>
											<% bgcolor = "FFFFFF" %>
										<% End If %>
										<tr>
											<td valign="top" width="260" height="20" align="left" class="OCStyle4" style="padding:3px 0 3px 10px;" bgcolor="#<%= bgcolor %>"><a href="cabin.asp?ID=<%= objRS4("ID") %>"><strong><%= objRS4("CabinNme")%></strong>&nbsp;&nbsp;<span class="cc2">est. <%=objRS4("EST") %></span></a><br><span class="cc3">Bedrooms: <%= objRS4("BedNum") %>&nbsp;&nbsp;|&nbsp;&nbsp;Bathrooms: <%= objRS4("BathNum") %></span></td>
											<td valign="top" width="110" align="right" class="OCStyle4" style="padding:3px 10px 3px 0;" bgcolor="#<%= bgcolor %>"><% If objRS4("FRateW") <> "N/A" Then %>$<% End If %><%= objRS4("FRateW") %></td>
											<td valign="top" width="110" align="right" class="OCStyle4" style="padding:3px 10px 3px 0;" bgcolor="#<%= bgcolor %>"><% If objRS4("FRateN") <> "N/A" Then %>$<% End If %><%= objRS4("FRateN") %></td>
										</tr>
									<% objRS4.MoveNext %>
									<% Loop %>
									<% objRS4.MoveFirst%>
									</Table>
								<br>
								</div>
									<br>
									<table cellpadding="0" cellspacing="0" border="0" width="480" class="OCStyle2" style="margin-top:3px;">
										<tr>
											<td height="18" width="18" align="center" class="OCStyle3"><a href="javascript:toggleLayer('Rate4');toggleLayer2('open4');toggleLayer3('close4');"><img src="images/close_Btn.gif" border="0"title="Open/Close" id="close4" hspace="3" style="visibility:hidden; display:none;"><img src="images/open_Btn.gif" border="0"title="Open/Close" id="open4" hspace="3" style="visibility:visible; display:inline;"></a></td>
											<td width="132" height="20" align="left" style="padding-left:10px;" class="OCStyle3"><a href="javascript:toggleLayer('Rate4');toggleLayer2('open4');toggleLayer3('close4');"><strong><%= season4 %></strong></a></td>
											<td width="330" height="20" align="right" style="padding-right:10px;" class="OCStyle3"><%= season4a %></td>
										</tr>
									</table>
									<div id="Rate4" style="visibility:hidden; margin: 0px 0px 0px 0px; display: none;">
									<Table cellpadding="0" cellspacing="0" border="0" width="480" class="OCStyle2">
										<tr>
											<td width="260" height="20" align="left" class="OCStyle4" style="padding-left:10px;" bgcolor="#F0EFE2"><strong>Cabin Name</strong></td>
											<td width="110" align="right" class="OCStyle4" bgcolor="#F0EFE2" style="padding-right:10px;"><strong>Week</strong></td>
											<td width="110" align="right" class="OCStyle4" bgcolor="#F0EFE2" style="padding-right:10px;"><strong>Night</strong></td>
										</tr>
									<% bgcolor = "EAEAEA" %>
									<% Do While Not objRS4.EOF %>
										<% If bgcolor = "FFFFFF" Then %>
											<% bgcolor = "EAEAEA" %>
										<% Else %>
											<% bgcolor = "FFFFFF" %>
										<% End If %>
										<tr>
											<td valign="top" width="260" height="20" align="left" class="OCStyle4" style="padding:3px 0 3px 10px;" bgcolor="#<%= bgcolor %>"><a href="cabin.asp?ID=<%= objRS4("ID") %>"><strong><%= objRS4("CabinNme")%></strong>&nbsp;&nbsp;<span class="cc2">est. <%=objRS4("EST") %></span></a><br><span class="cc3">Bedrooms: <%= objRS4("BedNum") %>&nbsp;&nbsp;|&nbsp;&nbsp;Bathrooms: <%= objRS4("BathNum") %></span></td>
											<td valign="top" width="110" align="right" class="OCStyle4" style="padding:3px 10px 3px 0;" bgcolor="#<%= bgcolor %>"><% If objRS4("WRateW") <> "N/A" Then %>$<% End If %><%= objRS4("WRateW") %></td>
											<td valign="top" width="110" align="right" class="OCStyle4" style="padding:3px 10px 3px 0;" bgcolor="#<%= bgcolor %>"><% If objRS4("WRateN") <> "N/A" Then %>$<% End If %><%= objRS4("WRateN") %></td>
										</tr>
									<% objRS4.MoveNext %>
									<% Loop %>
									</Table>
								<br>
								</div>
								
								
								
								<br>
								<%= objRS3("Body") %>
								<br><br><br>
								</td>
							</tr>
						</table>
						
<!--#include file="includes/footer.asp"-->








</body>
</html>
<% 

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