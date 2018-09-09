<%@ Language=VBScript %>
<% Option Explicit %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>

<!--#include file="includes/DBconnect.asp"-->
<!--#include file="includes/adovbs.inc"-->
<!--#include file="includes/iif.asp"-->

<%
Dim SQLString3, objRS3
SQLString3 = "Select * FROM tblGeneral where PageID = '8000'"
Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet

Dim SQLString4, objRS4
SQLString4 = "Select * FROM tblGeneral where PageID = '1150'"
Set objRS4 = Server.CreateObject("ADODB.Recordset")
objRS4.Open SQLString4, objConn, AdOpenKeySet

Dim SQLString, objRS
SQLString = "Select * FROM tblGeneral where PageID = '9000'"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet


If objRS3("Viewable") = 1 then 
	session("SP") = 1
else
	session("SP") = 0
End If

If objRS4("Viewable") = 1 then 
	session("WT") = 1
else
	session("WT") = 0
End If
%>


<!--#include file="includes/wbrHeader-home.asp"-->
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
	</div><!--http://www.youtube.com/watch?v=FyBT-_WwC-Ehttps://youtu.be/vFn1wvzVeQw-->
    <div id="youtube"><a href="http://woodlandbeachresort.com/videos/"><img src="images/video.png" alt="WBR Video" border="0" /></a></div>
	<% If session("WT") = 1 then %>
		<div id="winter"><a href="/winter/"><img src="images/winter-stay.png" alt="WBR Winter Specials" width="177" height="203" border="0" /></a></div>
	<% End IF %>
	<% If session("SP") = 1 then %>
		<div id="specials"><a href="/specials/"><img src="images/specials.png" alt="WBR Specials" width="180" height="214" border="0" /></a></div>
	<% End IF %>
<!--#include file="includes/wbrSides.asp"-->
<!--#include file="includes/wbrMenu.asp"-->
			<script type="text/javascript">
			  $(function() {
				 $('#photo').crossSlide({
					sleep: 3,
					fade: 1
				 }, [
					
					{ src: 'images/wbr-03.jpg' },
					{ src: 'images/wbr-05.jpg' },
					{ src: 'images/wbr-02.jpg' },
					{ src: 'images/wbr-06.jpg' },
					{ src: 'images/wbr-07.jpg' },
					{ src: 'images/wbr-04.jpg' },
					{ src: 'images/wbr-01.jpg' }
				 ]);
			  });
			</script>
			
			<div id="photo"><!-- Loading? --></div>
			
			<div id="homeText">
			<h1><img src="images/a_place.gif" width="440" height="27" border="0" alt="<%= objRS("Title") %>"></h1>
			<%= objRS("Body") %>
			</div>
<!--#include file="includes/wbrFooter.asp"-->
