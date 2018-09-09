<%@ Language=VBScript %>
<% Option Explicit %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>

<!--#include file="../../includes/DBconnect.asp"-->
<!--#include file="../../includes/adovbs.inc"-->
<!--#include file="../../includes/iif.asp"-->

<%
Dim SQLString, objRS, SQLString4, objRS4, SQLString5, objRS5, ID, SQLString6, objRS6

ID = request.QueryString("ID")

SQLString = "Select * FROM tblCabins where CabinNum = '" & ID & "'"

SQLString4 = "SELECT tblCabins.*, tblRates.*, tblRates.Display AS RateDisplay FROM tblCabins INNER JOIN tblRates ON tblCabins.CabinNme = tblRates.Cabin WHERE tblCabins.CabinNum = '" & ID & "'"

SQLString5 = "Select * FROM tblDates"


Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

Set objRS4 = Server.CreateObject("ADODB.Recordset")
objRS4.Open SQLString4, objConn, AdOpenKeySet

Set objRS5 = Server.CreateObject("ADODB.Recordset")
objRS5.Open SQLString5, objConn, AdOpenKeySet


SQLString6 = "Select * FROM tblCabins where BedNum = '" & objRS("BedNum") & "' and Display = '1'"
Set objRS6 = Server.CreateObject("ADODB.Recordset")
objRS6.Open SQLString6, objConn, AdOpenKeySet

%>


<!--#include file="../../includes/wbrHeader.asp"-->

	<div id="maps">
		<% IF ID = 50 Then %>
		<img src="/images/the-lmr.png" alt="The Lodge Meeting Room at Woodland Beach Resort. (WBR)" width="295" height="367" border="0" usemap="#Maps" />
		<map name="Maps">
		<area shape="poly" coords="260,326,48,333,43,28,249,21,260,327" href="/cabins/cabin/?ID=51" border="0">
		</map>
	<% Else %>
		<img src="/images/floor-<%=ID%>.png" alt="Click to Floorplan of Cabin <%=ID%>. (WBR)" width="295" height="367" border="0" usemap="#Maps" />
		<map name="Maps">
		<area shape="poly" coords="260,326,48,333,43,28,249,21,260,327" href="javascript:lyteflash('floor.asp?ID=<%= objRS("CabinNum") %>','540','440','<%= objRS("CabinNme") %> Floorplans/Layout')" border="0">
	<% End If %>
	</map>
	</div>
	<div id="albmPhoto">
		<img src="/images/<%=smPhoto %>.png" alt="Photos at Woodland Beach Resort. (WRB)" width="258" height="293" border="0" usemap="#PA" />
	<map name="PA">
	  <area shape="poly" coords="199,7,9,42,53,283,246,246,200,7" href="/about/?ID=photo" border="0">
	</map>
	</div>
<!--#include file="../../includes/wbrSides2.asp"-->
<!--#include file="../../includes/wbrMenu.asp"-->

<div id="bodyText">
<table cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td valign="top" >
		<h1><img src="/images/cabin_<%= objRS("CabinNum") %>_title.jpg" width="267" height="26" border="0" alt="<%= objRS("CabinNme") %>"></h1>
		<p>
		<%= objRS("LongDes") %>
		<br><br>
		</p>
		<p>
		<strong>Bedrooms: <%= objRS("BedNum") %></strong>&nbsp;&nbsp;|&nbsp;&nbsp;<strong>Bathrooms: <%= objRS("BathNum") %></strong>&nbsp;&nbsp;|&nbsp;&nbsp;<strong>Sleeps: <%= objRS("SleepNum") %></strong>
		</p>
		<% IF ID < 50 Then %>
		<p class="cabinsSM">
		Other <%= objRS("BedNum") %> Bedroom Cabins: 
		<% Do While Not objRS6.EOF %>
		<% IF cint(objRS6("CabinNum")) = cint(ID) Then %>
		<% Else %>
		<a href="?ID=<%= objRS6("CabinNum") %>"><%= objRS6("CabinNum") %></a> |
		<% End If %>
		<% objRS6.MoveNext %>
		<% Loop %>
		</p>
		<% End IF %>
		</td>
		<td width="231" valign="top" width="331">
		<img src="/images/<%= objRS("MainPhoto") %>-lg.jpg" width="331" height="267" border="0" alt="<%= objRS("CabinNme") %> : Est.<%= objRS("EST") %> on Bay Lake, MN"><br><a href="photos.asp?FID=<%= objRS("FlickrID") %>&CID=<%= objRS("CabinNum") %>" rel="lyteframe" title="More Photos From <%= objRS("CabinNme") %> : Est.<%= objRS("EST") %>" rev="width: 609px; height: 480px; scrolling: no;"><img src="/images/btn-view-photos.jpg" width="119" height="17" border="0" alt="View More Photos" align="right" style="padding-right:25px; "/></a>
		</td> 
	</tr>
</table>
<br>
<img src="/images/spacer-line.jpg" width="615" height="8" border="0">
<br><br>
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
<% If objRS4("RateDisplay") = "1" Then %>
<table cellpadding="0" cellspacing="0" border="0" width="632" height="206" background="/images/cabin_rate.gif" style="background-repeat:no-repeat;" class="boatRates">
	<tr>
		<td width="632" valign="top" style="padding:0 0 0 15px;" colspan="6" height="30"><h2 style="text-transform:uppercase;"><%= objRS("CabinNme") %> RATE SHEET</h2></td>
	</tr>
	<% IF ID = 51 Then %>
	<tr>
		<td width="21" valign="top">&nbsp;</td>
		<td width="187" valign="top"><h3>DATES</h3></td>
		<td width="38" valign="top" align="center">&nbsp;</td>
		<td width="83" valign="top" align="center"><h3>WEEK</h3></td>
		<td width="86"valign="top" align="center"><h3>4 HRS</h3></td>
		<td width="190" style="padding:0px 0 0 10px;" valign="top"><h3>THINGS TO REMEMBER</h3></td>
	</tr>
	<% else %>
	<tr>
		<td width="21" valign="top">&nbsp;</td>
		<td width="187" valign="top"><h3>DATES</h3></td>
		<td width="38" valign="top" align="center">&nbsp;</td>
		<td width="83" valign="top" align="center"><h3>WEEK</h3></td>
		<td width="86"valign="top" align="center"><h3>NIGHT</h3></td>
		<td width="190" style="padding:0px 0 0 10px;" valign="top"><h3>THINGS TO REMEMBER</h3></td>
	</tr>
	<% end if %>
	<tr>
		<td width="21" valign="top">&nbsp;</td>
		<td width="187" valign="top"><%= season3 %></td>
		<td width="38" valign="top" align="center">&nbsp;</td>
		<td width="83" valign="top" align="center"><% If objRS4("FRateW") <> "N/A" Then %>$<% End If %><%= objRS4("FRateW") %></td>
		<td width="86" valign="top" align="center"><% If objRS4("FRateN") <> "N/A" Then %>$<% End If %><%= objRS4("FRateN") %></td>
		<td width="190" style="padding:0px 27px 0 10px;" valign="top" rowspan="4"><%= objRS("CabinThings") %></td>
	</tr>
	<tr>
		<td width="21" valign="top">&nbsp;</td>
		<td width="187" valign="top"><%= season1 %></td>
		<td width="38" valign="top" align="center">&nbsp;</td>
		<td width="83" valign="top" align="center"><% If objRS4("SMRateW") <> "N/A" Then %>$<% End If %><%= objRS4("SMRateW") %></td>
		<td width="86" valign="top" align="center"><% If objRS4("SMRateN") <> "N/A" Then %>$<% End If %><%= objRS4("SMRateN") %></td>
	</tr>
	<tr>
		<td width="21" valign="top">&nbsp;</td>
		<td width="187" valign="top"><%= season2 %></td>
		<td width="38" valign="top" align="center">&nbsp;</td>
		<td width="83" valign="top" align="center"><% If objRS4("SPRateW") <> "N/A" Then %>$<% End If %><%= objRS4("SPRateW") %></td>
		<td width="86" valign="top" align="center"><% If objRS4("SPRateN") <> "N/A" Then %>$<% End If %><%= objRS4("SPRateN") %></td>
	</tr>
	<tr>
		<td height="100%" width="21" valign="top">&nbsp;</td>
		<td width="187" valign="top"><%= season4 %></td>
		<td width="38" valign="top" align="center">&nbsp;</td>
		<td width="83" valign="top" align="center"><% If objRS4("WRateW") <> "N/A" Then %>$<% End If %><%= objRS4("WRateW") %></td>
		<td width="86" valign="top" align="center"><% If objRS4("WRateN") <> "N/A" Then %>$<% End If %><%= objRS4("WRateN") %></td>
	</tr>
</table>
<% End IF %>
<br><br>
<br><br><br><br><br>
</div>
<!--#include file="../../includes/wbrLinks.asp"-->
<!--#include file="../../includes/wbrFooter.asp"-->
<%   
objConn.Close
Set objConn = Nothing
%>
