<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("item")="ads" Then
  SQL="Select Title_imgURL As Upload_FileURL From ADS Where ser_no='"&request("ser_no")&"'"
Else
  SQL="Select Upload_FileURL From UPLOAD Where ser_no='"&request("ser_no")&"'"
End If
Call QuerySQL(SQL,RS)
If Not RS.EOF Then
  Upload_FileURL=RS("Upload_FileURL")
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>影音連結</title>
</head>
<body topmargin="0" leftmargin="0" bgcolor="#dcdcba" rightmargin="0" bottommargin="0" background="../images/loading.gif">
<%Response.Write Upload_FileURL%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->