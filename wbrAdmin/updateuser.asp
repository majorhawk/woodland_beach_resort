

<%@ Language=VBScript %>
<%  Option Explicit%>

<!--#include file="connect.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<!--#include file="security.asp"-->


<html>
<head>
	<title>.:: <%= Session("ComNme") %> ::.</title>
</head>
<LINK href="css/main.css" type="text/css" rel="stylesheet">
<meta name="robots" content="noindex, nofollow" />
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0"  bgcolor="#F7F7F8" background="images/bg.gif">


<% 
Dim canc, objRS, SQLString2, objRS2

if request.QueryString("new") = "1" then

canc = Request.Form("Cancel")

If (canc="")then

else
Response.Redirect "users.asp"
End If


Set objRS = Server.CreateObject ("ADODB.Recordset")
objRS.Open "tblUsers", objConn, , adLockOptimistic, adCmdTable

objRS.AddNew

objRS("FNme") = Request.Form("nme")
objRS("LNme") = Request.Form("lnme")
objRS("UserNme") = Request.Form("login")
objRS("Password") = Request.Form("password")
objRS("Admin") = Request.Form("Rights")
objRS("Edit_On") = NOW()
objRS("Added_On") = NOW()
objRS("Edit_By") = Session("FirstName")


objRS.Update

objRs.Close
Set objRS = Nothing

objConn.Close
Set objConn = Nothing

 Response.Redirect "users.asp"

else

end if 


'The information below talks to the database and adds or deletes the information

Dim bolFound, strProdName, del2, del
strProdName = Request.Form("ID")
del = Request.Form("Delete")
del2 = Request.Form("Delete2")

canc = Request.Form("Cancel")


If (del="")then
	If (canc="")then

	else
		Response.Redirect "users.asp"
	End If




Set objRS = Server.CreateObject ("ADODB.Recordset")
objRS.Open "tblUsers", objConn, , adLockOptimistic, adCmdTable

bolFound = False

Do Until objRS.EOF OR bolFound
If (StrComp(objRS("ID"), strProdName, _
vbTextCompare) = 0) Then

BOlFound = True
Else
objRS.MoveNext
End IF
Loop

objRS("FNme") = Request.Form("nme")
objRS("LNme") = Request.Form("lnme")
objRS("UserNme") = Request.Form("login")
objRS("Password") = Request.Form("password")
objRS("Admin") = Request.Form("Rights")
objRS("Edit_On") = NOW()
objRS("Edit_By") = Session("FirstName")


	If (del2="")then
		objRS.Update
	else
		objRS.Delete
	End If




objRs.Close
Set objRS = Nothing

objConn.Close
Set objConn = Nothing

Response.Redirect "users.asp"

else

'The statement below pulls the information for the delete information
Dim objRS3,SQLStrng3,ID

SQLStrng3 = "Select * From tblUsers Where ID = " + request.Form("ID")

Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLStrng3, objConn, AdOpenKeySet


%>	
<!--#include file="includes/header.inc"-->
<form action="updateuser.asp" method="POST">
<table border="0" width="500" cellpadding="3" cellspacing="2" align="center" class="surBody">
	<tr>
		<td width="100%" colspan="2"><strong>Are you sure that you want to delete the user: <font color="#ff0000"><%=objRS3("FNme")%>&nbsp;<%=objRS3("LNme")%></font></strong></td>
	</tr><tr>
		<td width="10%" class="head2">Name:</td>
		<td valign="top" width="90%"><%=objRS3("FNme")%><input type="hidden" name="nme" value="<%=objRS3("FNme")%>" size="30"></td>
	</tr><tr>
		<td width="10%" class="head2">Login:</td> 
		<td valign="top" width="90%"><%=objRS3("UserNme")%><input type="hidden" name="login" value="<%=objRS3("UserNme")%>" size="30"></td>
	</tr><tr>
		<td width="10%" class="head2" valign="top">&nbsp;</td> 
		<td valign="top" class="body2"><input type="hidden" name="password" value="<%=objRS3("Password")%>" size="30"></td>
	</tr><tr>
		<td colspan="2"><input type="Submit" value="Yes" name="Delete2" class="inputSubmit"> <input type="submit" value="&nbsp;No&nbsp;" name="Cancel" class="inputSubmit"></td>
		</tr>
</table>
		<input type="Hidden" name="lnme" value="<%=objRS3("LNme")%>" size="30">
		<input type="Hidden" name="ID" value="<%=objRS3("ID")%>">
		<input type="Hidden" name="rights" value="<%=objRS3("Admin")%>">
</form>

<!--#include file="includes/footer.inc"-->

<%
objRs3.Close
Set objRS3 = Nothing


objConn.Close
Set objConn = Nothing

End If
%>








</body>
</html>
