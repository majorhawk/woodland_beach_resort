

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
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0"  bgcolor="#FFFFFF" background="../images/bg.gif">


<!--- home page text starts here --->

<!--#include file="includes/header.inc"-->


<!--- home page text starts here --->

<%

if request.QueryString("new") = "1" then

canc = Request.Form("Cancel")

If (canc="")then

else
Response.Redirect "rates.asp"
End If


Set objRS = Server.CreateObject ("ADODB.Recordset")
objRS.Open "tblRates", objConn, , adLockOptimistic, adCmdTable

objRS.AddNew


objRS("Cabin") = Request.Form("Cabin")
objRS("SPRateW") = Request.Form("SPRateW")
objRS("SPRateN") = Request.Form("SPRateN")
objRS("SMRateW") = Request.Form("SMRateW")
objRS("SMRateN") = Request.Form("SMRateN")
objRS("FRateW") = Request.Form("FRateW")
objRS("FRateN") = Request.Form("FRateN")
objRS("WRateW") = Request.Form("WRateW")
objRS("WRateN") = Request.Form("WRateN")
objRS("Edit_By") = Session("FirstName")
objRS("Added_By") = Session("FirstName")
objRS("Edit_On") = NOW()
objRS("Added_On") = NOW()
objRS("Display") = Request.Form("display")

objRS.Update

objRs.Close
Set objRS = Nothing

objConn.Close
Set objConn = Nothing


 Response.Redirect "rates.asp"

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
Response.Redirect "rates.asp"
End If

Set objRS4 = Server.CreateObject ("ADODB.Recordset")
objRS4.Open "tblRates", objConn, , adLockOptimistic, adCmdTable

bolFound = False

Do Until objRS4.EOF OR bolFound
If (StrComp(objRS4("ID"), strProdName, _
vbTextCompare) = 0) Then

BOlFound = True
Else
objRS4.MoveNext
End IF
Loop

If (del2="")then
objRS4("Cabin") = Request.Form("Cabin")
objRS4("SPRateW") = Request.Form("SPRateW")
objRS4("SPRateN") = Request.Form("SPRateN")
objRS4("SMRateW") = Request.Form("SMRateW")
objRS4("SMRateN") = Request.Form("SMRateN")
objRS4("FRateW") = Request.Form("FRateW")
objRS4("FRateN") = Request.Form("FRateN")
objRS4("WRateW") = Request.Form("WRateW")
objRS4("WRateN") = Request.Form("WRateN")
objRS4("Edit_By") = Session("FirstName")
objRS4("Edit_On") = NOW()
objRS4("Display") = Request.Form("display")

objRS4.Update
else
objRS4.Delete
End If

objRs4.Close
Set objRS = Nothing

objConn.Close
Set objConn = Nothing


Response.Redirect "rates.asp"


else

'The statement below pulls the information for the delete information
Dim objRS3,SQLStrng3,ID, up

SQLStrng3 = "Select * From tblRates Where ID = " + request.Form("ID")

Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLStrng3, objConn, AdOpenKeySet



%>	
<table cellpadding="0" cellspacing="1" width="500" style="border: 1px solid #6193D9;" align="center">
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" width="500">
				<tr>
					<td height="23" class="surHeader" align="left" valign="middle" bgcolor="#FFFFFF"><img src="images/header_logo.gif" title="Woodland Beach Resort"></td>
					<td height="23" width="100%" class="surHeader" align="left" background="images/header_fill.gif">Are you sure that you want to delete this Rates of this Cabin?</td>
				</tr>
			</table>
			<table border="0" width="500" cellpadding="5" cellspacing="1" align="center" bgcolor="C0C0C0" class="surBody">
				<tr>
					<td align="center" class="surHeader" bgcolor="#FFFFFF">
						<form action="updateRates.asp" method="POST">
						<br>
						<font color="#ff0000"><strong><%=objRS3("Cabin")%></strong></font>
						<input type="Hidden" name="ID" value="<%=objRS3("ID")%>">
					</td>
				</tr>
				<tr>
					<td colspan="2" align="center" bgcolor="#FFFFFF"><input type="Submit" value="Yes" name="Delete2" class="inputSubmit"> <input type="submit" value="&nbsp;No&nbsp;" name="Cancel" class="inputSubmit"></td>
				</tr>
			</table>
		</td>
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

