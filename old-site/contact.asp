<%@ Language=VBScript %>
<% Option Explicit %>


<!--#include file="includes/DBconnect.asp"-->
<!--#include file="includes/adovbs.inc"-->

<%
Dim SQLString3, objRS3, mainPhoto

SQLString3 = "Select * FROM tblGeneral where PageID = '6000' and Viewable = '1'"






Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet




%>	


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<title><%= objRS3("Title") %> : Woodland Beach Resort - A Place Where Friends Become Family on Bay Lake, MN</title>

<link rel="STYLESHEET" type="text/css" href="css/Global.css">
<link rel="STYLESHEET" type="text/css" href="css/TopNav.css">

<script type="text/javascript" src="js/nav.js"></script>
<script type="text/javascript" src="js/contact.js"></script>

</head>
<body bgcolor="#F5F4F3" text="#000000" leftmargin="0" topmargin="0" bottommargin="0" rightmargin="0" style="margin:0;">


<!--#include file="includes/header.asp"-->

								<td valign="top" height="97"><h1 title="Contact Us: Let us know your questions and comments."><img src="images/title_contact.gif" width="540" height="97" border="0" alt="Contact Us: Let us know your questions and comments."></h1></td>
							</tr>
							<tr> 
								<!-- <td></td> -->
								<td valign="top" class="body" height="274">
								<%= objRS3("Body") %>
					<!-- CONTACT US EMIAL FORM STARTS HERE -->
								<% If Request.QueryString("PageID") = "6000" Then %>
									<% If Request.QueryString("email") = "1" Then %>
									<table cellpadding="0" cellspacing="0" border="0" width="450" height="24" style="padding-bottom:1px;">
										<tr>
											<td width="450" height="24" class="formHeader">Thank you</td>
										</tr>
									</table>
									<table cellpadding="4" cellspacing="0" border="0" width="450" bgcolor="#FBF7EF" style="border: 1px solid #BFB18E;">
										<tr>
											<td class="inputBorder1">Your information has been sent. <br>If you need further assistance please call us at 1.888.436.7770.<br></td>
										</tr>
									</table>
									<% Else %>
										<!--#include file="includes/contactForm.asp"-->
									<% End If %>
								<% End If %>					
					<!-- CONTACT US EMIAL FORM STARTS HERE -->
								<br><br><br>
								</td>
							</tr>
						</table>
						
<!--#include file="includes/footer.asp"-->








</body>
</html>
<% 
objRs3.Close
Set objRS3 = Nothing
%>

<%   
objConn.Close
Set objConn = Nothing
%>