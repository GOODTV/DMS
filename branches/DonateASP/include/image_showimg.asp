<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>圖片顯示</title>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#dcdcba">
<p align="center">
<%if right(request("imgfile"),3)="gif" or right(request("imgfile"),3)="jpg" then%>
<img border="0" src="../Upload/<%=request("imgfile")%>">
<%else
   response.redirect "../Upload/"&request("imgfile")
 end if%>  
 </p>
</body>

</html>