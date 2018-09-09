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
Dim SQLString, objRS, ID, typeID

If request.QueryString("ID") = "map" Then
ID = "1005"
typeID = "1"
ElseIf  request.QueryString("ID") = "meeting" Then
ID = "2005"
typeID = "1"
ElseIf  request.QueryString("ID") = "private" Then
ID = "2010"
typeID = "2"
Else
ID = "2000"
typeID = "1"
End If

SQLString = "Select * FROM tblGeneral where PageID = '" & ID & "' and Viewable = '1'"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

Dim SQLString3, objRS3
SQLString3 = "Select * FROM tblCabins where CabinType = '" & typeID & "' and Display = '1' order by CabinOrder"
Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet


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
<br><br>
</p>
<% If ID = "2000" or ID = "2010" Then %>
<% Do While Not ObjRS3.EOF %>
<table cellspacing="0" cellpadding="0" border="0" class="cabin">
	<tr>
		<td valign="top" align="center"><a href="cabin/?ID=<%= objRS3("CabinNum") %>" style="text-decoration:none;"><img src="/images/<%= objRS3("MainPhoto") %>.jpg" width="204" height="145" border="0"><br>Click to View Cabin</a>&nbsp;&nbsp;</td>
		<td width="398" valign="top" style="padding-left:10px;">
		<h2><%= objRS3("CabinNme") %></h2>
		<p style="font-size:11px; ">
		<strong>Bedrooms: <%= objRS3("BedNum") %></strong>&nbsp;&nbsp;|&nbsp;&nbsp;<strong>Bathrooms: <%= objRS3("BathNum") %></strong>&nbsp;&nbsp;|&nbsp;&nbsp;<strong>Sleeps: <%= objRS3("SleepNum") %></strong>
		</p>
		<p>
		<%= objRS3("LongDes") %>
		</p>
		</td>
	</tr>
</table>
<br>
<img src="/images/spacer-line.jpg" width="615" height="8" border="0">
<br><br>
<% objRS3.MoveNext %>
<% Loop %>
<% End IF %>
<br><br><br><br><br>
</div>
<!--#include file="../includes/wbrLinks.asp"-->
<!--#include file="../includes/wbrFooter.asp"-->

<%   
objConn.Close
Set objConn = Nothing
%>
