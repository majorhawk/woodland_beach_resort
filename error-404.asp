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

Dim SQLString, objRS
SQLString = "Select * FROM tblGeneral where PageID = '9500'"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet


If objRS3("Viewable") = 1 then 
	session("SP") = 1
else
	session("SP") = 0
End If
%>


<!--#include file="includes/wbrHeader.asp"-->
<!--#include file="includes/wbrSides.asp"-->
<!--#include file="includes/wbrMenu.asp"-->
			<div id="homeText" style="padding-left:30px; ">
			<h1>Error 404 - File Not Found</h1>
			<%= objRS("Body") %>
			</div>
<!--#include file="includes/wbrFooter.asp"-->
