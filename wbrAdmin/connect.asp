<%
Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = "DSN=wbr2010;UID=bigfish_hochstaetter;PWD=paygejaden"
objConn.Open
%>