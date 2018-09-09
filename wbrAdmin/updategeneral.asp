

<%@ Language=VBScript %>
<%  Option Explicit%>


<!--#include file="connect.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<!--#include file="security.asp"-->

<%
Dim objRS, SQLString, SQLString2, objRS2, canc
%>

<html>
<head>
	<title>.:: <%= Session("ComNme") %> ::.</title>
</head>
<LINK href="css/main.css" type="text/css" rel="stylesheet">
<meta name="robots" content="noindex, nofollow" />
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0"  bgcolor="#010066" background="../images/bg.gif">


<!--- home page text starts here --->

<!--#include file="includes/header.inc"-->


<!--- home page text starts here --->

<%

if request.QueryString("new") = "1" then

canc = Request.Form("Cancel")

If (canc="")then

else
Response.Redirect "general.asp"
End If


Set objRS = Server.CreateObject ("ADODB.Recordset")
objRS.Open "tblPage", objConn, , adLockOptimistic, adCmdTable

objRS.AddNew

objRS("PageID") = Request.Form("PageID")
objRS("HTML") = Request.Form("HTML")
objRS("Title") = Request.Form("Title")
objRS("BodySection") = Request.Form("BodySection")
objRS("Searchable") = Request.Form("Searchable")
objRS("Edit_On") = Request.Form("Edit_On")
objRS("Edit_By") = Request.Form("Edit_By")

objRS.Update

objRs.Close
Set objRS = Nothing

objConn.Close
Set objConn = Nothing


 Response.Redirect "general.asp"

else

end if 

'The information below talks to the database and adds or deletes the information

Dim objRS4, bolFound, strProdName, del, del2
strProdName = Request.Form("ID")
del = Request.Form("Delete")
del2 = Request.Form("Delete2")

canc = Request.Form("Cancel")

If (del="")then

If (canc="")then

else
Response.Redirect "general.asp"
End If

Set objRS4 = Server.CreateObject ("ADODB.Recordset")
objRS4.Open "tblPage", objConn, , adLockOptimistic, adCmdTable

bolFound = False

Do Until objRS4.EOF OR bolFound
If (StrComp(objRS4("ID"), strProdName, _
vbTextCompare) = 0) Then

BOlFound = True
Else
objRS4.MoveNext
End IF
Loop

objRS4("PageID") = Request.Form("PageID")
objRS4("HTML") = Request.Form("HTML")
objRS4("Title") = Request.Form("Title")
objRS4("BodySection") = Request.Form("BodySection")
objRS4("Searchable") = Request.Form("Searchable")
objRS4("Edit_On") = Request.Form("Edit_On")
objRS4("Edit_By") = Request.Form("Edit_By")

If (del2="")then
objRS4.Update
else
objRS4.Delete
End If

objRs4.Close
Set objRS = Nothing

objConn.Close
Set objConn = Nothing


Response.Redirect "general.asp"


else

'The statement below pulls the information for the delete information
Dim objRS3,SQLStrng3,ID, up

SQLStrng3 = "Select * From tblPage Where ID = " + request.Form("ID")

Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLStrng3, objConn, AdOpenKeySet



%>	

<table border="0" width="500" cellpadding="3" cellspacing="2" align="center" class="body">
	<tr>
		<td>
			<form action="updategeneral.asp" method="POST">
			<span class="head3">Are you sure that you want to delete this Page?
			<br><img src="../images/clear.gif" width="1" height="6" alt="" border="0"><br>
			<font color="#ff0000" size="-1"><%=objRS3("Title")%>.</font></span>
		</td>
	</tr>
</table>
<table border="0" width="500" cellpadding="5" cellspacing="1" align="center" class="body">
	<tr>
		<td>
		<input type="Hidden" name="ID" value="<%=objRS3("ID")%>">
		</td>
	</tr><tr>
		<td colspan="2"><input type="Submit" value="Yes" name="Delete2"> <input type="submit" value="&nbsp;No&nbsp;" name="Cancel"></td>
		</tr>
</table>
</form>



<%
objRs3.Close
Set objRS3 = Nothing


objConn.Close
Set objConn = Nothing

End If


%>	


<!--#include file="includes/footer.inc"-->

