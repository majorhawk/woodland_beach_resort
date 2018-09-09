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

ID = "6000"




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
<div style="width:619px; height:286px; background-image:url(/images/map-frame.jpg); background-repeat:no-repeat; padding:23px 0 0 24px;">
<iframe width="567" height="233" frameborder="0" scrolling="no" marginheight="0" marginwidth="0" src="http://maps.google.com/maps?f=q&amp;source=s_q&amp;hl=en&amp;geocode=&amp;q=15596+Woodland+Beach+Lane+Deerwood,+MN+56444&amp;sll=37.0625,-95.677068&amp;sspn=61.452931,116.279297&amp;ie=UTF8&amp;hq=&amp;hnear=15596+Woodland+Beach+Ln,+Deerwood,+Crow+Wing,+Minnesota+56444&amp;ll=46.371925,-93.856888&amp;spn=0.027597,0.097504&amp;z=13&amp;iwloc=A&amp;output=embed"></iframe><br /><br /><small><a href="http://maps.google.com/maps?f=q&amp;source=embed&amp;hl=en&amp;geocode=&amp;q=15596+Woodland+Beach+Lane+Deerwood,+MN+56444&amp;sll=37.0625,-95.677068&amp;sspn=61.452931,116.279297&amp;ie=UTF8&amp;hq=&amp;hnear=15596+Woodland+Beach+Ln,+Deerwood,+Crow+Wing,+Minnesota+56444&amp;ll=46.371925,-93.856888&amp;spn=0.027597,0.097504&amp;z=13&amp;iwloc=A" style="color:#0000FF;text-align:left">View Larger Map</a></small></div>


<!-- CONTACT US EMIAL FORM STARTS HERE -->
				<% 'If Request.QueryString("email") = "1" Then %>
				<div style="float:right; display:block; padding:15px 30px 0 0; width:320px;"><img src="/images/spacer-line2.jpg" width="9" height="325" border="0" alt="" align="left" style="margin:0px 20px 0px 0px; "><%= objRS("Body") %></div>
				<div>
				<table cellspacing="0" width="210" height="232" border="0" style="background-image:url(/images/thanks.jpg); background-repeat:no-repeat;">
					<tr>
						<td valign="top" style="padding:100px 40px 20px 40px; font-size:16px; ">If you need further assistance please call us at: <br><br><strong>1.888.436.7770</strong></td>
					</tr>
				</table>
				</div>
				<br><br><br><br><br><br>
				<% 'Else %>
				<!-- <div style="float:right; display:block; padding:15px 30px 0 0; width:280px;"><img src="/images/spacer-line2.jpg" width="9" height="325" border="0" alt="" align="left" style="margin:0px 20px 0px 0px; "><%' objRS("Body") %></div>
				<div><h3><img src="/images/comment-form-title.jpg" width="200" height="17" border="0" alt="Questions and Eomments Form"></h3><!-- include here --></div>
				<!--<br><br><br><br>  -->
				<% 'End If %>			
<!-- CONTACT US EMIAL FORM STARTS HERE -->



</p>
</div>
<!--#include file="../includes/wbrLinks.asp"-->
<!--#include file="../includes/wbrFooter.asp"-->
<script type="text/javascript" src="/js/contact.js"></script>
<%   
objConn.Close
Set objConn = Nothing
%>

