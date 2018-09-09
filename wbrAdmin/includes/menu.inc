<table cellpadding="0" cellspacing="0" border="0" width="185">
				
<%


SQLString2 = "Select * FROM tblUsers where ID = '" & Session("ID") & "'"



Set objRS2 = Server.CreateObject("ADODB.Recordset")
objRS2.Open SQLString2, objConn, AdOpenKeySet





If (objRS2("Admin") = "1") Then%>

				<tr>
					<td class="Nav2" width="185" height="21" background="images/head.gif" valign="middle" colspan="2">Administrator Information</td>
				</tr>
				<tr>
					<td class="Nav" width="27" height="25" valign="middle"><img src="images/user1.gif" width="27" height="25" alt="" border="0"></td>
					<td class="Nav3" width="158" height="25" valign="middle" align="left"><a href="users.asp">Edit/Add Users</a></td>
				</tr>
<% End If  %>	
				<tr>
					<td class="Nav2" width="185" height="21" background="images/head.gif" valign="middle" colspan="2">Website Information</td>
				</tr>
				<tr>
					<td class="Nav" width="27" height="25" valign="middle"><img src="images/icon_pages.gif" width="27" height="25" alt="" border="0" vspace="2"></td>
					<td class="Nav3" width="158" height="25" valign="middle" align="left"><a href="pages.asp">Edit/Add Pages</a></td>
				</tr>
				<tr>
					<td class="Nav" width="27" height="25" valign="middle"><img src="images/icon_prod.gif" width="27" height="25" alt="" border="0" vspace="2"></td>
					<td class="Nav3" width="158" height="25" valign="middle" align="left"><a href="cabins.asp">Edit/Add Cabins</a></td>
				</tr>
				<tr>
					<td class="Nav" width="27" height="25" valign="middle"><img src="images/icon_prod.gif" width="27" height="25" alt="" border="0" vspace="2"></td>
					<td class="Nav3" width="158" height="25" valign="middle" align="left"><a href="boats.asp">Edit/Add Boats</a></td>
				</tr>
				<tr>
					<td class="Nav" width="27" height="25" valign="middle"><img src="images/icon_acc.gif" width="27" height="25" alt="" border="0" vspace="2"></td>
					<td class="Nav3" width="158" height="25" valign="middle" align="left"><a href="rates.asp">Edit/Add Rates</a></td>
				</tr>
				<tr>
					<td class="Nav" width="27" height="25" valign="middle"><img src="images/icon_news.gif" width="27" height="25" alt="" border="0" vspace="2"></td>
					<td class="Nav3" width="158" height="25" valign="middle" align="left"><a href="dates.asp">Edit/Add Rate Dates</a></td>
				</tr>
				<tr>
					<td class="Nav2" width="185" height="21" background="images/head.gif" valign="middle" colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td class="Nav" width="27" height="25" valign="middle"><img src="images/log.gif" width="27" height="25" alt="" border="0"></td>
					<td class="Nav3" width="158" height="25" valign="middle" align="left"><a href="index.asp">Log Out</a></td>
				</tr>
</table>

	
			
			
