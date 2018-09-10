<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cabin Photos</title>
<%
Dim FID, CID

CID = Request.QueryString("CID")
FID = Request.QueryString("FID")

%>
<style>
body {
	margin:0;
	background-color:#ededed;
	text-align:center;
	}
</style>
</head>

<body>
<iframe src="https://www.flickr.com/photos/46656001@N07/albums/<%=FID%>/player/" width="599" height="470" frameborder="0" allowfullscreen webkitallowfullscreen mozallowfullscreen oallowfullscreen msallowfullscreen></iframe>
</body>
</html>



