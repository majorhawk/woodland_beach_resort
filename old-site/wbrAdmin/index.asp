

<%@ Language=VBScript %>
<% Option Explicit %>

<!--#include file="connect.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<%
'Session.Abandon
'Session.Timeout=5
%>

<html>
<head>
<%dim SQLString, objRS

SQLString = "Select * FROM tblCompany"

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open SQLString, objConn, AdOpenKeySet

Session("ComNme") = objRS("ComNme") 
 
objRs.Close
Set objRS = Nothing
 
 %>
 
<title>.:: <%= Session("ComNme") %> ::.</title>
<script type="text/javascript">
	function login(){
	document.lg.Login.focus();
	}
</script>	
</head>
<LINK href="css/main.css" type="text/css" rel="stylesheet">
<meta name="robots" content="noindex, nofollow" />
<body onLoad=login() topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0"  bgcolor="#F7F7F8" background="images/bg.gif">


<!--#include file="includes/header.inc"-->

	<FORM ACTION="adminGet2.asp" METHOD="post" name="lg">
<TABLE BORDER=0 align="center" class="surBody">
<TR>
		<TD colspan="2" align="center" class="surTitles3"><%= Session("ComNme") %> Site Editor Login<br><br></TD>
		<!--- <TD></TD> --->
	</TR>
	<TR>
		<TD ALIGN="center" colspan="2">
			<Table cellpadding="0" cellspacing="0" border="0" class="surBody">
				<TR>
					<TD ALIGN="right">User Name:&nbsp;</TD>
					<TD><INPUT TYPE="text" NAME="Login" value="" class="input1"></INPUT></TD>
				</TR>
				<TR>
					<TD ALIGN="right">Password:&nbsp;</TD>
					<TD><INPUT TYPE="password" NAME="Password" class="input1"></INPUT>&nbsp;&nbsp;<INPUT TYPE="submit" VALUE="Login" class="inputSubmit"></INPUT></TD>
				<TR>
			</Table>
		</TD>
	</TR>

		<TD colspan="2" align="center"><br><% If Request("access") = "no" Then %> 
<font color="#ff0000">The user name and password you supplied was invalid. <br>Please enter a valid user name and password</font>
<% End If %></TD>
		<!--- <TD></TD> --->
	</TR>
</TABLE>
</FORM>
	
<!--#include file="includes/footer.inc"-->

<% objConn.Close
Set objConn = Nothing %>
