<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <%If request("SQL")<>"" Then%>	
  <META HTTP-EQUIV="refresh" CONTENT="1;URL=<%=request("URL")%>?SQL=<%=request("SQL")%>">
  <%Else%>
  <META HTTP-EQUIV="refresh" CONTENT="1;URL=<%=request("URL")%>">
  <%End If%>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>資料處理中</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
  <img border="0" src="../images/movebar.gif">
  <p>&nbsp;&nbsp;&nbsp;&nbsp; 資料處理中 ........</p> 
</body>
</html>