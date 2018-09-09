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

ID = "7000"




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
<!--#include file="../includes/wbrSides2.asp"-->
<!--#include file="../includes/wbrMenu.asp"-->
<div id="bodyText">
<h1><%= objRS("Title") %></h1>
<p>
<%= objRS("Body") %>
</p>
<a name="colors"></a>
<table cellpadding="0" cellspacing="0" border="0" style="margin:20px 0 0 0; ">
	<tr>
		<td valign="top" align="center"><img src="/images/mens-clothing.jpg" width="125" height="24" border="0"></td>
		<td valign="middle"><img src="/images/spacer-shirts3.jpg" width="50" height="24" border="0" alt=""></td>
		<td valign="top" align="center"><img src="/images/womens-clothing.jpg" width="149" height="24" border="0"></td>
	</tr>
	<tr>
		<td colspan="3" style="padding:10px 0 10px 0;"><img src="/images/spacer-line.jpg" width="615" height="8" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="/images/mens-1.jpg" width="230" height="152" border="0" alt="WBR Mens Sweatshirt" name="m" align="right">
		<div class="colors">
		<ul>
		<li><a href="#colors" onClick="ChangeImage ('m','mens-1.jpg','m-1-lg.jpg','IDm'); return true"><img src="/images/sm-1.gif" width="17" height="17" border="1" alt="Brown"></a></li>
		<li><a href="/images/m-1-lg.jpg" rel="lytebox" title="WBR Mens Sweatshirt" id="IDm"><img src="/images/zoom.jpg" width="21" height="23" border="0" alt="zoom"></a>
		</ul>
		</div>
		<div style="width:250px; display:block; clear:both;">Available Sizes: S,M,L,LX&nbsp;&nbsp;| Price: <strong>$44.95</strong></div>
		</td>
		<td valign="middle"><img src="/images/spacer-shirts.jpg" width="50" height="152" border="0" alt=""></td>
		<td valign="top"><img src="/images/womens-1.jpg" width="230" height="152" border="0" alt="WBR Womens Sweatshirt" name="w" align="right">
		<div class="colors">
		<ul>
		<li><a href="#colors" onClick="ChangeImage ('w','womens-1.jpg','w-1-lg.jpg','IDw'); return true"><img src="/images/sw-1.gif" width="17" height="17" border="1" alt="Red"></a></li>
		<li><a href="/images/w-1-lg.jpg" rel="lytebox" title="WBR Womens Sweatshirt" id="IDw"><img src="/images/zoom.jpg" width="21" height="23" border="0" alt="zoom"></a>
		</ul>
		</div>
		<div style="width:250px; display:block; clear:both;">Available Sizes: S,M,L,LX&nbsp;&nbsp;| Price: <strong>$44.95</strong></div>
		</td>
	</tr>
	<tr>
		<td colspan="3" style="padding:10px 0 10px 0;"><img src="/images/spacer-line.jpg" width="615" height="8" border="0"></td>
	</tr>
	<tr>
		<td valign="top" style="padding:0 0 20px 0;"><img src="/images/mens-h-1.jpg" width="230" height="194" border="0" alt="WBR Mens Hooded Sweatshirt" name="mh" align="right">
		<div class="colors">
		<ul>
		<li><a href="#colors" onClick="ChangeImage ('mh','mens-h-1.jpg', 'mh-1-lg.jpg','IDmh'); return true"><img src="/images/sm-h-1.gif" width="17" height="17" border="1" alt="Gray"></a></li>
		<li><a href="/images/mh-1-lg.jpg" rel="lytebox" title="WBR Mens Hooded Sweatshirt" id="IDmh"><img src="/images/zoom.jpg" width="21" height="23" border="0" alt="zoom"></a>
		</ul>
		</div>
		<div style="width:250px; display:block; clear:both;">Available Sizes: S,M,L,LX&nbsp;&nbsp;| Price: <strong>$49.95</strong></div>
		</td>
		<td valign="middle"><img src="/images/spacer-shirts2.jpg" width="50" height="194" border="0" alt=""></td>
		<td valign="top"><img src="/images/womens-h-1.jpg" width="230" height="194" border="0" alt="WBR Womens Hooded Sweatshirt" name="wh" align="right">
		<div class="colors">
		<ul>
		<li><a href="#colors" onClick="ChangeImage ('wh','womens-h-2.jpg','wh-2-lg.jpg','IDwh'); return true"><img src="/images/sw-h-2.gif" width="17" height="17" border="1" alt="Blue"></a></li>
		<li><a href="#colors" onClick="ChangeImage ('wh','womens-h-3.jpg','wh-3-lg.jpg','IDwh'); return true"><img src="/images/sw-h-3.gif" width="17" height="17" border="1" alt="Plum"></a></li>
		<li><a href="#colors" onClick="ChangeImage ('wh','womens-h-4.jpg','wh-4-lg.jpg','IDwh'); return true"><img src="/images/sw-h-4.gif" width="17" height="17" border="1" alt="Orange"></a></li>
		<li><a href="#colors" onClick="ChangeImage ('wh','womens-h-1.jpg','wh-1-lg.jpg','IDwh'); return true"><img src="/images/sw-h-1.gif" width="17" height="17" border="1" alt="Pink"></a></li>
		<li><a href="/images/wh-1-lg.jpg" rel="lytebox" title="WBR Womens Hooded Sweatshirt" id="IDwh"><img src="/images/zoom.jpg" width="21" height="23" border="0" alt="zoom"></a>
		</ul>
		</div>
		<div style="width:250px; display:block; clear:both;">Available Sizes: S,M,L,LX&nbsp;&nbsp;| Price: <strong>$49.95</strong></div>
		</td>
	</tr>
	<tr>
		<td colspan="3" style="padding:10px 0 10px 0;"><img src="/images/spacer-line.jpg" width="615" height="8" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="/images/mens-t-1.jpg" width="230" height="163" border="0" alt="WBR Men's T-Shirt" name="mt" align="right">
		<div class="colors">
		<ul>
		<li><a href="#colors" onClick="ChangeImage ('mt','mens-t-2.jpg','mt-2-lg.jpg','IDmt'); return true"><img src="/images/sm-t-2.gif" width="17" height="17" border="1" alt="Green"></a></li>
		<li><a href="#colors" onClick="ChangeImage ('mt','mens-t-1.jpg','mt-1-lg.jpg','IDwt'); return true"><img src="/images/sm-t-1.gif" width="17" height="17" border="1" alt="Blue"></a></li>
		<li><a href="/images/mt-1-lg.jpg" rel="lytebox" title="WBR Men's T-Shirt" id="IDmt"><img src="/images/zoom.jpg" width="21" height="23" border="0" alt="zoom"></a></li>
		</ul>
		</div>
		<div style="width:250px; display:block; clear:both;">Available Sizes: S,M,L,LX&nbsp;&nbsp;| Price: <strong>$19.95</strong></div>
		</td>
		<td valign="middle"><img src="/images/spacer-shirts.jpg" width="50" height="152" border="0" alt=""></td>
		<td valign="top"><img src="/images/blank.gif" width="230" height="163" border="0" alt=""></td>
	</tr>
	<tr>
		<td colspan="3" style="padding:10px 0 10px 0;"><img src="/images/spacer-line.jpg" width="615" height="8" border="0"></td>
	</tr>
</table>
<br><br>
</div>
<!--#include file="../includes/wbrLinks.asp"-->
<!--#include file="../includes/wbrFooter.asp"-->

<%   
objConn.Close
Set objConn = Nothing
%>

