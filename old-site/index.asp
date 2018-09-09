<%@ Language=VBScript %>
<% Option Explicit %>


<!--#include file="includes/DBconnect.asp"-->
<!--#include file="includes/adovbs.inc"-->

<%
dim Random, smPhoto
Function RandomNumber(intHighestNumber)
	Randomize
	RandomNumber = Int(Rnd * intHighestNumber) + 1
End Function

Random = RandomNumber(8)


Dim SQLString3, objRS3, mainPhoto

SQLString3 = "Select * FROM tblGeneral where PageID = '8000'"

Set objRS3 = Server.CreateObject("ADODB.Recordset")
objRS3.Open SQLString3, objConn, AdOpenKeySet


mainPhoto = "mainPhoto01"

smPhoto = "smPhoto0" & Random


If objRS3("Viewable") = 1 then 
	session("SP") = 1
else
	session("SP") = 0
End If

'***************************************************************************
'* BELOW IS THE FUNCTION THE SET THE UNIQUE KEY CODE FOR EACH IMAGE *
'***************************************************************************

dim key, i
  
	function createKeyString(length)
	  dim dstr
	  dim num
		randomize timer
		dstr = ""
	  for i = 0 to length
		  num = int((rnd * 43) + 48)
		  if num >= 58 and num <= 64 then num = num + 7
		  dstr = dstr & chr(num)
	  next
		createkeystring = dstr
	end function 

key = createkeystring(1)


%>	
<%'= RandomNumber(3) %>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<title>.:: Woodland Beach Resort - Located on Bay Lake in Beautiful Minnesota ::.</title>

<link rel="STYLESHEET" type="text/css" href="css/Global.css">
<link rel="STYLESHEET" type="text/css" href="css/TopNav.css">
<meta name="verify-v1" content="sxvMI5Qw8p3VndoV//YiWtAUI3VRcR2PgkjeRn/3QzY=" />

<script type="text/javascript" src="js/nav.js"></script>


<script type="text/javascript">
<!--
    if (document.images)
    {
		
        image2 = new Image(96, 34);
        image2.src = "images/menu02_roll.gif";
		
        image3 = new Image(736, 34);
        image3.src = "images/menu03_roll.gif";
		
        image4 = new Image(75, 34);
        image4.src = "images/menu04_roll.gif";
		
        image5 = new Image(74, 34);
        image5.src = "images/menu05_roll.gif";
		
        image6 = new Image(81, 34);
        image6.src = "images/menu06_roll.gif";
		
        image6 = new Image(102, 34);
        image6.src = "images/menu07_roll.gif";
		
        image6 = new Image(98, 34);
        image6.src = "images/menu08_roll.gif";
		
        image6 = new Image(104, 34);
        image6.src = "images/menu09_roll.gif";
    }
	
//-->
</script>
</head>
<body bgcolor="#F5F4F3" text="#000000" leftmargin="0" topmargin="0" bottommargin="0" rightmargin="0" style="margin:0;">

<table cellpadding="0" cellspacing="0" border="0" width="774" height="568" align="center"> 
	<tr> 
		<td width="11" background="images/leftSide.gif"><img src="images/clear.gif" width="11"></td>
		<td>
			<table cellpadding="0" border="0" cellspacing="0" height="549" width="752" align="center" class="homeBox">
				<tr>
					<td valign="top" height="100">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/WBR_Logo.gif" alt="Woodland Beach Resort Logo" width="250" height="100"></td>
								<td width="500" height="100" background="images/header.gif" align="right" valign="bottom" class="Time" style="padding-bottom:50px; padding-right:5px;"><span id="tP">&nbsp;</span>&nbsp;<script type="text/javascript" src="js/clock.js"></script>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" height="34"><!--#include file="includes/TopMenu.inc"--></td>
				</tr>
				<tr>
					<td valign="top" height="274" bgcolor="#FFFFFF">
<!-- MIDDLE CONTENT AND SPECIALS SECTION -->
						<table cellpadding="0" cellspacing="0" border="0" width="750">
							<tr>
								<td valign="top" width="110" height="274" background="images/map.jpg" rowspan="3" class="mapText"><img src="images/clear.gif" width="110" height="161"><br><span id="mapB">&raquo;</span> <a href="pdfs/lakeMap.pdf" target="_blank">Lake Map</a><br><span id="mapB">&raquo;</span> <a href="pdfs/roadMap.pdf" target="_blank">Road Map</a><br><span id="mapB">&raquo;</span> <a href="pdfs/resortMap.pdf" target="_blank">Resort Map</a></td>
								<td rowspan="3" height="274"><img src="images/shutterLeft.jpg" width="78" height="274" border="0"></td>
								<td align="right" height="226" rowspan="2" background="images/<%= mainPhoto%>.gif">
								<% If session("SP") = 1 then %>
								<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width="214" height="226" id="specials" align="middle">
								<param name="allowScriptAccess" value="sameDomain" />
								<param name="movie" value="flash/specials.swf" />
								<param name="quality" value="high" />
								<param name="wmode" value="transparent" />
								<param name="bgcolor" value="#ffffff" />
								<embed src="flash/specials.swf" quality="high" wmode="transparent" bgcolor="#ffffff" width="214" height="226" name="specials" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
								</object>
								<% End If %>
								</td>
								<td rowspan="3" height="274"><img src="images/shutterRight.jpg" width="48" height="274" border="0"></td>
							</tr>
							<tr>
								<!-- <td></td> -->
								<!-- <td></td> -->
								<!-- <td></td> -->
								<!-- <td></td> -->
							</tr>
							<tr>
								<!-- <td></td> -->
								<!-- <td></td> -->
								<td width="414" height="48" background="images/middleFill.jpg" class="homeText2" valign="top">Located on beautiful Bay Lake, a 2,435 acre lake of bays, points<br>& islands, that is one of the top bass fishing lakes in Minnesota.</td>
								<!-- <td></td> -->
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" height="105" background="images/backGrad.gif">
<!-- BOTTOM CALLOUT BOXES --><!-- BOTTOM CALLOUT BOXES --><!-- BOTTOM CALLOUT BOXES -->
						<table cellpadding="0" cellspacing="0" border="0" width="750">
							<tr>
								<td height="105" width="249" style="border-right: 1px solid #6F0000;" class="homeText"><a href="about.asp?PageID=1010"><img src="images/home01.jpg" alt="Updates and Additions" width="142" height="103" border="0" align="left"><img src="images/updates_title.gif" alt="Updates and Additions" border="0" style="margin:19px 0 5px 0;" title="Updates and Additions"><br>See What We’ve Added & Where We're Going</a></td>
								<td height="105" width="249" style="border-right: 1px solid #6F0000;" class="homeText"><a href="cabins.asp?PageID=2000"><img src="images/home02.jpg" alt="Experience a Family Tradition" width="142" height="103" border="0" align="left"><img src="images/family_title.gif" alt="Experience a Family Tradition" border="0" style="margin:19px 0 5px 0;" title="Experience a Family Tradition"><br>Earlier Days Have Been Blended With Today</a></td>
								<td width="250" class="homeText"><a href="about.asp?PageID=1020"><img src="images/<%=smPhoto %>.jpg" alt="Memories of Yeasterday" width="142" height="103" border="0" align="left"><img src="images/yesterday_title.gif" alt="Memories of Yeasterday" border="0" style="margin:19px 0 5px 0;" title="Memories of Yeasterday"><br>Images of our Friends & Family</a></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" height="17" bgcolor="#FFFFFF"><img src="images/clear.gif" width="568" height="17"></td>
				</tr>
				<tr>
					<td valign="middle" height="28" background="images/bottomLines.gif" class="cc" align="center">Copyright © 2002 - <%= year(now())%> Woodland Beach Resort - Bay Lake, MN&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://www.brainerd.com" target="_blank">www.brainerd.com</a></td>
				</tr>
			</table>
		</td>
		<td width="11" background="images/rightSide.gif"><img src="images/clear.gif" width="11"></td>
	</tr>
	<tr> 
		<td colspan="3" width="774" height="17"><img src="images/bottom.gif" width="774" height="17"></td>
	</tr>
</table>



<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-1004322-1";
urchinTracker();
</script>


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