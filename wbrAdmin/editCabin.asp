

<%@ Language=VBScript %>
<%  Option Explicit%>


<!--#include file="connect.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<!--#include file="security.asp"-->

<% Dim objRS, SQLString, SQLString2, objRS2 %>

<html>
<head>
	<title>.:: <%= Session("ComNme") %> ::.</title>
</head>
<% ' You must set "Enable Parent Paths" on your web site in order this relative include to work. %>
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->

<LINK href="css/main.css" type="text/css" rel="stylesheet">
<meta name="robots" content="noindex, nofollow" />
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0"  bgcolor="#F7F7F8" background="images/bg.gif">


<!--- home page text starts here --->

<!--#include file="includes/header.inc"-->

<!--- home page text starts here --->

<%
Dim objRS3,SQLStrng3,ID

SQLStrng3 = "Select * From tblCabins Where ID = " + request.QueryString("ID")

Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLStrng3, objConn, AdOpenKeySet



'************************************
'* BELOW IS THE INLINE IF STATEMENT *
'************************************
Function iif(trueFalse, a, b)
	If cbool(trueFalse) Then
		iif = a
	Else
		iif = b
	End If
End Function
%>	

<form action="updateCabins.asp" method="POST">
<table cellpadding="0" cellspacing="1" width="500" style="border: 1px solid #841D17;" align="center">
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" width="500">
				<tr>
					<td height="23" class="surHeader" align="left" valign="middle" bgcolor="#FFFFFF"><img src="images/header_logo.gif" title="Woodland Beach Resort"></td>
					<td height="23" width="100%" class="surHeader" align="left" background="images/header_fill.gif">Edit Cabin -  <%=objRS3("CabinNme")%></td>
				</tr>
			</table>
			<table border="0" width="500" cellpadding="5" cellspacing="1" align="center" bgcolor="C0C0C0" class="surBody">
				<tr>
					<td width="100" bgcolor="#ffffff">Edited On:</td>
					<td valign="top" class="body2a" bgcolor="#ffffff"><input type="hidden" name="Edit_On" value="<%=NOW%>"><%=NOW%></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Edited By:</td>
					<td valign="top" class="body2a" bgcolor="#ffffff"><input type="hidden" name="Edit_By" value="<%=Session("FirstName")%>" style="WIDTH: 395px;"><%=Session("FirstName")%></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Viewable:</td> 
					<td width="400" valign="top" bgcolor="#ffffff"><input type="radio" name="viewable" value="1"<%=iif( objRS3("Display") = 1," checked","")%>>&nbsp;Yes&nbsp;&nbsp;&nbsp;<input type="radio" name="viewable" value="0"<%=iif( objRS3("Display") = 0," checked","")%>>&nbsp;No</td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Cabin Type:</td> 
					<td width="400" valign="top" bgcolor="#ffffff"><input type="radio" name="CabinType" value="1"<%=iif( objRS3("CabinType") = 1," checked","")%>>&nbsp;Resort&nbsp;&nbsp;&nbsp;<input type="radio" name="CabinType" value="2"<%=iif( objRS3("CabinType") = 2," checked","")%>>&nbsp;Private&nbsp;&nbsp;&nbsp;<input type="radio" name="CabinType" value="3"<%=iif( objRS3("CabinType") = 3," checked","")%>>&nbsp;Other</td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Cabin Name:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="CabinNme" style="WIDTH: 395px;" value="<%=objRS3("CabinNme")%>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Cabin Number:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="CabinNum" style="WIDTH: 395px;" value="<%=objRS3("CabinNum")%>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Cabin Order:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="Order" style="WIDTH: 395px;" value="<%=objRS3("CabinOrder")%>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">EST. Date:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="EST" style="WIDTH: 395px;" value="<%=objRS3("EST")%>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Bedrooms:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="BedNum" style="WIDTH: 395px;" value="<%=objRS3("BedNum")%>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Bathrooms:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="BathNum" style="WIDTH: 395px;" value="<%=objRS3("BathNum")%>"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Sleeps:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="SleepNum" style="WIDTH: 395px;" value="<%=objRS3("SleepNum")%>"></td>
				</tr>
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">Long Description:</td>
					<td width="400" valign="top" bgcolor="#ffffff">
					<%
					' Automatically calculates the editor base path based on the _samples directory.
					' This is usefull only for these samples. A real application should use something like this:
					' oFCKeditor.BasePath = /fckeditor/ ;	// /fckeditor/ is the default value.
					Dim sBasePath
					sBasePath = Request.ServerVariables("PATH_INFO")
					sBasePath = "fckeditor/" 
					
					Dim oFCKeditor
					Set oFCKeditor = New FCKeditor
					oFCKeditor.BasePath = sBasePath
					
					If Request.QueryString("Toolbar") <> "" Then
						oFCKeditor.ToolbarSet = Server.HTMLEncode( Request.QueryString("Toolbar") )
					End If
					
					oFCKeditor.Value = objRS3("LongDes")
					oFCKeditor.Create "LongDes"
					%>
				</td>
				</tr>
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">FloorPlan:</td>
					<td width="400" valign="top" bgcolor="#ffffff"><% If (objRS2("Admin") = "1") Then %><textarea style="WIDTH: 390px; HEIGHT: 150px;" name="ShortDes" wrap="off" class="body2a"><%=objRS3("ShortDes")%></textarea><% Else %><input type="hidden" name="ShortDes" value="<%=objRS3("ShortDes")%>"><%=objRS3("ShortDes")%><% End IF %></td>
				</tr>
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">Main Photo:<br></td>
					<td width="400" valign="top" bgcolor="#ffffff"><% If (objRS2("Admin") = "1") Then %><input type="Text" name="MainPhoto" style="WIDTH: 395px;" value="<%=objRS3("MainPhoto")%>"><% Else %><input type="hidden" name="MainPhoto" value="<%=objRS3("MainPhoto")%>"><%=objRS3("MainPhoto")%><% End IF %></td>
				</tr>
				
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">Flickr ID:<br></td>
					<td width="400" valign="top" bgcolor="#ffffff"><% If (objRS2("Admin") = "1") Then %><input type="Text" name="FlickrID" style="WIDTH: 395px;" value="<%=objRS3("FlickrID")%>"><% Else %><input type="hidden" name="FlickrID" value="<%=objRS3("FlickrID")%>"><%=objRS3("FlickrID")%><% End IF %></td>
				</tr>
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">Things to Remember:</td>
					<td width="400" valign="top" bgcolor="#ffffff">
					<%
					' Automatically calculates the editor base path based on the _samples directory.
					' This is usefull only for these samples. A real application should use something like this:
					' oFCKeditor.BasePath = /fckeditor/ ;	// /fckeditor/ is the default value.
					
					If Request.QueryString("Toolbar") <> "" Then
						oFCKeditor.ToolbarSet = Server.HTMLEncode( Request.QueryString("Toolbar") )
					End If
					
					oFCKeditor.Value = objRS3("CabinThings")
					oFCKeditor.Create "CabinThings"
					%>
					</td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Meta Title:</td>
					<td width="400" valign="top" bgcolor="#ffffff"><input type="Text" name="metaTitle" style="WIDTH: 395px;" value="<%=objRS3("metaTitle")%>"></td>
				</tr>
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">Meta Description:</td>
					<td width="400" valign="top" bgcolor="#ffffff"><textarea style="WIDTH: 395px; HEIGHT: 150px;" name="metaDes" wrap="off" class="body2a"><%=objRS3("metaDes")%></textarea></td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="#ffffff" align="center"><input type="Submit" value="Update Cabin" name="update" class="inputSubmit"> <input type="button" value="Cancel" name="Cancel" class="inputSubmit" onClick="parent.location='cabins.asp'"> <% If (objRS2("Admin") = "1") Then %><input type="Submit" value="Delete" name="Delete" class="inputSubmit"><% End If %></td>
				</tr>
			</table>
			
			
			
			
		</td>
	</tr>
</table>
<input type="Hidden" name="ID" value="<%=objRS3("ID")%>">
</form>


<% 
objRs3.Close
Set objRS3 = Nothing

objConn.Close
Set objConn = Nothing
%>

<!--#include file="includes/footer.inc"-->


