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
Dim SQLString, objRS
SQLString = "Select * FROM tblGeneral where PageID = '3000' and Viewable = '1'"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

Dim SQLString3, objRS3
SQLString3 = "Select * FROM tblBoat where Display = '1' order by BoatNme"
Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet


%>


<!--#include file="../includes/wbrHeader.asp"-->
	<div id="maps">
		<img src="/images/maps.png" alt="Click to View Maps of Woodland Beach Resort. (WBR)" width="295" height="367" border="0" usemap="#Maps" />
	<map name="Maps">
	  <area shape="poly" coords="260,326,48,333,43,28,249,21,260,327" href="/maps/" border="0">
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
<p>
<%= objRS("Body") %>
</p>
<% Do While Not ObjRS3.EOF %>
<table cellspacing="0" cellpadding="0" border="0" class="boatRates">
	<tr>
		<td valign="top"><img src="/images/<%= objRS3("MainPhoto") %>" width="204" height="205" border="0"></td>
		<td width="408" valign="top">
		<table cellpadding="0" cellspacing="0" border="0" width="408" height="205" background="/images/boat_rate.gif" style="background-repeat:no-repeat; ">
			<tr>
				<td width="408" valign="top" style="padding:0 0 0 15px;" colspan="4" height="30"><h2><%= objRS3("BoatNme") %></h2></td>
			</tr>
			<tr>
				<td width="21" valign="top">&nbsp;</td>
				<td width="83" valign="top" align="center"><h3>DAY</h3></td>
				<td width="77" valign="top" align="center"><h3>WEEK</h3></td>
				<td width="190" style="padding:0px 0 0 10px;" valign="top"><h3>DESCRIPTION</h3></td>
			</tr>
			<tr>
				<td height="130" width="21" valign="top">&nbsp;</td>
				<td height="130" width="83" valign="top" align="center">$<%= objRS3("BoatRateD") %></td>
				<td height="130" width="77" valign="top" align="center">$<%= objRS3("BoatRateW") %></td>
				<td height="130" width="190" style="padding:0px 27px 0 10px;" valign="top"><%= objRS3("LongDes") %></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<% objRS3.MoveNext %>
<% Loop %>
<br><br><br><br><br>
</div>
<!--#include file="../includes/wbrLinks.asp"-->
<!--#include file="../includes/wbrFooter.asp"-->

<%   
objConn.Close
Set objConn = Nothing
%>
