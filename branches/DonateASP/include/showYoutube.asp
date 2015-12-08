<!--#include file="../include/dbinclude.asp"-->
<%
sql="select * from ads where ser_no='"&request("ser_no")&"'"
Call QuerySQL(SQL,RS)
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>圖片顯示</title>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#dcdcba" rightmargin="0" bottommargin="0" background="../images/loading.gif">
<%
   response.write rs("ad_URL")
%>  
 </p>
 </p>
</body>

</html>