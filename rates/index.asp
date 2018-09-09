<%@ Language=VBScript %>
<% Option Explicit %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>

<!--#include file="../includes/DBconnect.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/iif.asp"-->

<%
Dim SQLString, objRS, ID, typeID, SQLString5, objRS5


ID = "4000"
typeID = "1"


SQLString = "Select * FROM tblGeneral where PageID = '" & ID & "' and Viewable = '1'"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

Dim SQLString3, objRS3
SQLString3 = "Select * FROM tblCabins where CabinType = '" & typeID & "' and Display = '1' order by CabinOrder"
Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet

SQLString5 = "Select * FROM tblDates"
Set objRS5 = Server.CreateObject("ADODB.Recordset")
objRS5.Open SQLString5, objConn, AdOpenKeySet
%>


<!--#include file="../includes/wbrHeader.asp"-->
		<div id="maps">
		<img src="/images/the-lmr.png" alt="The Lodge Meeting Room at Woodland Beach Resort. (WBR)" width="295" height="367" border="0" usemap="#Maps" />
	<map name="Maps">
	  <area shape="poly" coords="260,326,48,333,43,28,249,21,260,327" href="/cabins/cabin/?ID=51" border="0">
	</map>
	</div>
	<div id="albmPhoto">
		<img src="/images/<%=smPhoto %>.png" alt="Photos at Woodland Beach Resort. (WRB)" width="258" height="293" border="0" usemap="#PA" />
	<map name="PA">
	  <area shape="poly" coords="199,7,9,42,53,283,246,246,200,7" href="/about/?ID=photo" border="0">
	</map>
	</div>
<!--#include file="../includes/wbrSides2.asp"-->
<!--#include file="../includes/wbrMenu.asp"-->

<div id="bodyText">
<h1><%= objRS("Title") %></h1>
<% Do While Not ObjRS3.EOF %>

<%
Dim SQLString2, objRS2, SQLString4, objRS4, SQLString6, objRS6

ID = ObjRS3("CabinNum")

SQLString4 = "SELECT tblCabins.*, tblRates.*, tblRates.Display AS RateDisplay FROM tblCabins INNER JOIN tblRates ON tblCabins.CabinNme = tblRates.Cabin WHERE tblCabins.CabinNum = '" & ID & "'"

Set objRS4 = Server.CreateObject("ADODB.Recordset")
objRS4.Open SQLString4, objConn, AdOpenKeySet

%>



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
		<td width="632" valign="top" style="padding:0 0 0 15px;" colspan="6" height="30"><h2 style="text-transform:uppercase;"><%= objRS3("CabinNme") %> RATE SHEET - <a href="/cabins/cabin/?ID=<%= objRS3("CabinNum") %>" class="rateLink">VIEW CABIN</a></h2></td>
	</tr>
	<tr>
		<td width="21" valign="top">&nbsp;</td>
		<td width="187" valign="top"><h3>DATES</h3></td>
		<td width="38" valign="top" align="center">&nbsp;</td>
		<td width="83" valign="top" align="center"><h3>WEEK</h3></td>
		<td width="86"valign="top" align="center"><h3>NIGHT</h3></td>
		<td width="190" style="padding:0px 0 0 10px;" valign="top"><h3>OTHER INFORMATION</h3></td>
	</tr>
	<tr>
		<td width="21" valign="top">&nbsp;</td>
		<td width="187" valign="top"><%= season3 %></td>
		<td width="38" valign="top" align="center">&nbsp;</td>
		<td width="83" valign="top" align="center"><% If objRS4("FRateW") <> "N/A" Then %>$<% End If %><%= objRS4("FRateW") %></td>
		<td width="86" valign="top" align="center"><% If objRS4("FRateN") <> "N/A" Then %>$<% End If %><%= objRS4("FRateN") %></td>
		<td width="190" style="padding:0px 27px 0 10px;" valign="top" rowspan="4">Bedrooms: <%= objRS3("BedNum") %><br>Bathrooms: <%= objRS3("BathNum") %><br>Sleeps: <%= objRS3("SleepNum") %></td>
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
		<td height="50" width="21" valign="top">&nbsp;</td>
		<td width="187" valign="top"><%= season4 %></td>
		<td width="38" valign="top" align="center">&nbsp;</td>
		<td width="83" valign="top" align="center"><% If objRS4("WRateW") <> "N/A" Then %>$<% End If %><%= objRS4("WRateW") %></td>
		<td width="86" valign="top" align="center"><% If objRS4("WRateN") <> "N/A" Then %>$<% End If %><%= objRS4("WRateN") %></td>
	</tr>
</table>
<br>
<img src="/images/spacer-line.jpg" width="615" height="8" border="0">
<% End IF %>
<br>
<% objRS3.MoveNext %>
<% Loop %>
<p>
<%= objRS("Body") %>
<br><br>
</p>
<br><br><br><br><br>
</div>
<!--#include file="../includes/wbrLinks.asp"-->
<!--#include file="../includes/wbrFooter.asp"-->

<%   
objConn.Close
Set objConn = Nothing
%>

