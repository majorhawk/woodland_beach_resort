<%
Dim SQLString6, objRS6

SQLString6 = "Select * FROM tblGeneral where PageID = '8000'"

Set objRS6 = Server.CreateObject("ADODB.Recordset")
objRS6.Open SQLString6, objConn, AdOpenKeySet


If objRS6("Viewable") = 1 then 
	session("SP") = 1
else
	session("SP") = 0
End If


objRS6.Close
Set objRS6 = Nothing

%>	

<table cellpadding="0" cellspacing="0" border="0" width="774" height="100" align="center"> 
	<tr> 
		<td width="11" background="images/leftSide.gif"><img src="images/clear.gif" width="11"></td>
		<td>
			<table cellpadding="0" border="0" cellspacing="0" height="100" width="752" align="center" class="homeBox">
				<tr>
					<td valign="top" height="100">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><a href="index.asp"><img src="images/WBR_Logo.gif" alt="Woodland Beach Resort Logo" width="250" height="100" border="0" title="Woodland Beach Resort Home Page"></a></td>
								<td width="500" height="100" background="images/header.gif" align="right" valign="bottom" class="Time" style="padding-bottom:50px; padding-right:5px;"><span id="tP">&nbsp;</span>&nbsp;<script type="text/javascript" src="js/clock.js"></script>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" height="34"><!--#include file="TopMenu.inc"--></td>
				</tr> 
				<tr>
					<td valign="top" height="274" bgcolor="#FFFFFF">
<!-- MIDDLE CONTENT AND SPECIALS SECTION -->
						<table cellpadding="0" cellspacing="0" border="0" width="750" height="100">
							<tr> 
								<td rowspan="2" valign="top" width="210" height="274" background="images/map2.jpg" class="mapText2" style="border-right:2px solid #CECAC7;">
								<img src="images/clear.gif" width="210" height="161"><br>
								<img src="images/clear.gif" width="98" height="50" align="left"><span id="mapB">&raquo;</span> <a href="pdfs/LakeMap.pdf" target="_blank">Lake Map</a><br>
								<span id="mapB">&raquo;</span> <a href="pdfs/roadMap.pdf" target="_blank">Road Map</a><br>
								<span id="mapB">&raquo;</span> <a href="pdfs/resortMap.pdf" target="_blank">Resort Map</a>
								<br><br>
								<a href="meeting.asp"><img src="images/meeting.gif" width="210" vspace="55" border="0"></a>
								<br>
								</td>