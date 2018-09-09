<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cabin Floor Plans</title>

<!--#include file="../../includes/DBconnect.asp"-->
<!--#include file="../../includes/adovbs.inc"-->
<!--#include file="../../includes/iif.asp"-->

<%
Dim SQLString, objRS, ID

ID = request.QueryString("ID")

SQLString = "Select * FROM tblCabins where CabinNum = '" & ID & "'"

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

%>
<style>
body {
	margin:0;
	background-color:#ffffff;
	text-align:center;
	}
</style>
</head>

<body>
<center><img src="<%=objrs("ShortDes")%>" border="0" alt="Cabin <%=ID%> Floorplan" style="margin-top:30px;"></center>
</body>
</html>



