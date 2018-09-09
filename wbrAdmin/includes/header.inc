
<table cellpadding="0" cellspacing="0" border="0" width="780" height="100%" align="center" class="border">
<tr>
	<td valign="top" align="left" colspan="2" width="780" height="74"><img src="images/top.jpg" width="780" alt="" border="0"></td>
	<!--- <td valign="top"></td> --->
</tr>
<tr>
	<td colspan="2" height="32" width="780" background="images/fill1.gif" align="right" class="surTitles" style="padding-right: 15px;"><%= Session("ComNme") %> Site Editor 
<%  if instr(request.serverVariables("PATH_INFO"), "index.asp") > 0 then %>	

<% else %>

: <font color="#ff0000"><font style="text-transform: capitalize"><%=Session("FirstName")%></font> is Logged-in</font>

<% end if %>

</td>
	<!--- <td></td> --->
</tr>
<tr>
	<td height="100%" width="185" valign="top" bgcolor="#F8F8F8" class="border2">		

<%  if instr(request.serverVariables("PATH_INFO"), "index.asp") > 0 then %>	
&nbsp;
<% else %>

<!--#include file="menu.inc"-->

<% end if %>
	</td>
	<td height="374" width="595" valign="top"><br>