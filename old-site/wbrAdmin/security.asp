<% If Session("FirstName") = "" and Session("ComNme") = "" THEN
 Response.Redirect "index.asp?access=no"  
End If
%>