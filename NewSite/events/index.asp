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

ID = "1140"




SQLString = "Select * FROM tblGeneral where PageID = '" & ID & "' and Viewable = '1'"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

%>

<!--#include file="../includes/wbrHeader.asp"-->

<style>

h2{
	padding-top:10px;
	}
</style>

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
<br><br>
<div id="PictoBrowser100117152120" style="border:8px solid #ffffff;">Get the flash player here: http://www.adobe.com/flashplayer</div><script type="text/javascript" src="http://www.db798.com/pictobrowser/swfobject.js"></script><script type="text/javascript"> var so = new SWFObject("http://www.db798.com/pictobrowser.swf", "PictoBrowser", "599", "470", "8", "#EEEEEE"); so.addVariable("source", "sets"); so.addVariable("names", "Events"); so.addVariable("userName", "Woodland Beach Resort"); so.addVariable("userId", "46656001@N07"); so.addVariable("ids", "72157623180419425"); so.addVariable("titles", "on"); so.addVariable("displayNotes", "on"); so.addVariable("thumbAutoHide", "off"); so.addVariable("imageSize", "medium"); so.addVariable("vAlign", "mid"); so.addVariable("vertOffset", "0"); so.addVariable("colorHexVar", "EEEEEE"); so.addVariable("initialScale", "off"); so.addVariable("bgAlpha", "90"); so.write("PictoBrowser100117152120");	</script>
<br><br><br><br><br>
</div>
<!--#include file="../includes/wbrLinks.asp"-->
<!--#include file="../includes/wbrFooter.asp"-->

<%   
objConn.Close
Set objConn = Nothing
%>

