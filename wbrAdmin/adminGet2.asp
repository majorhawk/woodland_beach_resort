

<%@ Language=VBScript %>
<%  Option Explicit%>


<%
if instr(Request.Form("Login"),"'" ) >0 or instr(Request.Form("Login"),"+" ) >0 or instr(Request.Form("Login"),";" ) >0 or instr(Request.Form("Login"),"-" ) >0 or instr(Request.Form("Login"),"=" ) >0 or instr(Request.Form("Login"),"," ) >0 or instr(Request.Form("Login"),"""" ) >0 or instr(Request.Form("Login"),"(" ) >0 or instr(Request.Form("Login"),")" ) >0 or instr(Request.Form("Login"),"{" ) >0 or instr(Request.Form("Login"),"}" ) >0 or instr(Request.Form("Login"),"[" ) >0 or instr(Request.Form("Login"),"]" ) >0 or instr(Request.Form("Login"),"." ) >0 or instr(Request.Form("password"),"'" ) >0 or instr(Request.Form("password"),"+" ) >0 or instr(Request.Form("password"),";" ) >0 or instr(Request.Form("password"),"-" ) >0 or instr(Request.Form("password"),"=" ) >0 or instr(Request.Form("password"),"," ) >0 or instr(Request.Form("password"),"""" ) >0 or instr(Request.Form("password"),"(" ) >0 or instr(Request.Form("password"),")" ) >0 or instr(Request.Form("password"),"{" ) >0 or instr(Request.Form("password"),"}" ) >0 or instr(Request.Form("password"),"[" ) >0 or instr(Request.Form("password"),"]" ) >0 or instr(Request.Form("password"),"." ) >0 then
   Response.Redirect "index.asp?access=no"
end if
%>




<!--#include file="connect.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<%
Dim objRS, SQLString

Request.Form("Login") 
Request.Form("password") 

SQLString = "Select UserNme, Password, FNme, LNme, ID FROM tblUsers where UserNme = '" & Request.Form("Login") & "' and Password = '"  & Request.Form("password") & "'"

Set objRS = Server.CreateObject ("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet
%>



<html>
<head>
	<title>.:: <%= Session("ComNme") %> ::. </title>
</head>
<meta name="robots" content="noindex, nofollow" />
<LINK href="main.css" type="text/css" rel="stylesheet">

<body>

<!--- home page text starts here --->



<%
If (not objRS.EOF) then

Session("ID") = objRS("ID")

Session("FirstName") = objRS("FNme") & " " & objRS("LNme")

Response.Redirect "general.asp"



else
   Response.Redirect "index.asp?access=no"
end if



objConn.Close
Set objConn = Nothing

%>


</body>
</html>
