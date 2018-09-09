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
Dim SQLString, objRS, ID

ID = "1130"




SQLString = "Select * FROM tblGeneral where PageID = '" & ID & "' and Viewable = '1'"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

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
<p>
<%= objRS("Body") %>
</p>
<br>
<table cellpadding="0" cellspacing="0" border="0" class="maps" style="margin-bottom:130px; ">
	<tr>
		<td><p><a href="/pdfs/lakeMap.pdf" target="_blank"><img src="/images/map_lake.jpg" height="279" width="196" border="0" alt="Baky Lake, Brainerd MN Lake Map"/>
<br>Download Lake Map of Bay Lake</a></p></td>
		<td><p><a href="/pdfs/roadMap.pdf" target="_blank"><img src="/images/map_road.jpg" height="279" width="196" border="0" alt="Road Map to Woodland Beach Resort on Bay Lake, MN"/>
<br>Download WBR Road Map</a></p></td>
		<td><p><a href="/pdfs/resortMap.pdf" target="_blank"><img src="/images/map_resort.jpg" height="279" width="196" border="0" alt="Resort Map of Woodland Beach Resort on Bay Lake, MN"/>
<br>Download WBR Resort Map</a></p></td>
	</tr>
</table>
</div>
<!--#include file="../includes/wbrLinks.asp"-->
<!--#include file="../includes/wbrFooter.asp"-->

<%   
objConn.Close
Set objConn = Nothing
%>

