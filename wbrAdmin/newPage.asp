<%@ Language=VBScript %>
<%  Option Explicit%>

<!--#include file="connect.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<!--#include file="security.asp"-->

<%
Dim objRS, SQLString, SQLString2, objRS2
%>

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


<form action="updatePage.asp?new=1" method="POST">
<table cellpadding="0" cellspacing="1" width="500" style="border: 1px solid #841D17;" align="center">
	<tr>
		<td>
			<table cellpadding="0" cellspacing="0" width="500">
				<tr>
					<td height="23" class="surHeader" align="left" valign="middle" bgcolor="#FFFFFF"><img src="images/header_logo.gif" title="Woodland Beach Resort"></td>
					<td height="23" width="100%" class="surHeader" align="left" background="images/header_fill.gif">Add A New Page</td>
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
					<td width="100" bgcolor="#ffffff">Page ID:</td> 
					<td width="400" valign="top" bgcolor="#ffffff"><input type="Text" name="PageID" style="WIDTH: 395px;"></td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Viewable:</td> 
					<td width="400" valign="top" bgcolor="#ffffff"><input type="radio" name="viewable" value="1" checked>&nbsp;Yes&nbsp;&nbsp;&nbsp;<input type="radio" name="viewable" value="0">&nbsp;No</td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Title:</td>
					<td valign="top" bgcolor="#ffffff"><input type="Text" name="Title" style="WIDTH: 395px;"></td>
				</tr>
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">Body:</td>
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
					
					'oFCKeditor.Value = objRS3("Body")
					oFCKeditor.Create "Body"
					%>	</td>
				</tr>
				<tr>
					<td width="100" bgcolor="#ffffff">Meta Title:</td>
					<td width="400" valign="top" bgcolor="#ffffff"><input type="Text" name="metaTitle" style="WIDTH: 395px;"></td>
				</tr>
				<tr>
					<td width="100" valign="top" bgcolor="#ffffff">Meta Description:</td>
					<td width="400" valign="top" bgcolor="#ffffff"><textarea style="WIDTH: 395px; HEIGHT: 150px;" name="HTML" wrap="off" class="metaDes"></textarea></td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="#ffffff" align="center"><input type="Submit" value="Add Page" class="inputSubmit"> <input type="button" value="Cancel" name="Cancel" class="inputSubmit" onClick="parent.location='pages.asp'"> <input type="reset" value="Reset" class="inputSubmit"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>


<% 
objConn.Close
Set objConn = Nothing
%>

<!--#include file="includes/footer.inc"-->


