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

If request.QueryString("ID") = "fishing" Then
ID = "1015"
ElseIf  request.QueryString("ID") = "photo" Then
ID = "1020"
ElseIf  request.QueryString("ID") = "ac" Then
ID = "9002"
Else
ID = "1000"
End If



SQLString = "Select * FROM tblGeneral where PageID = '" & ID & "' and Viewable = '1'"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

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
<!--#include file="../includes/wbrSides.asp"-->
<!--#include file="../includes/wbrMenu.asp"-->

<div id="bodyText">
<h1><%= objRS("Title") %></h1>
<p>
<%= objRS("Body") %>
</p>

<% If  request.QueryString("ID") = "photo" Then %>


<iframe src="http://www.flickr.com/photos/46656001@N07/albums/72157623230132270/player/" width="640" height="480" frameborder="0" allowfullscreen webkitallowfullscreen mozallowfullscreen oallowfullscreen msallowfullscreen></iframe>


<% End If %>

</div>

<!--#include file="../includes/wbrFooter.asp"-->

<%   
objConn.Close
Set objConn = Nothing
%>
