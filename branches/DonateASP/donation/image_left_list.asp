<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>圖文繞圖檔上傳</title>
</head>
<body class=left style="background-color:#EEEEE3">
<a href style="text-decoration: none" onclick="MM_openBrWindow('image_left_upload.asp?dept_id=<%=request("dept_id")%>&item=<%=request("item")%>&ser_no=<%=request("ser_no")%>&img_width=&code_id=<%=request("code_id")%>','','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=90')"><button style="cursor:hand;" class="addbutton">圖文繞圖檔上傳</button></a>
<%
  SQL="select ser_no,圖片=Upload_FileURL from upload where object_id='"&request("ser_no")&"' and Ap_name='"&request("item")&"' and attach_type='img'"
  HLink="image_left_delete.asp?code_id="&request("code_id")&"&item="&request("item")&"&ser_no="
  LinkParam="ser_no"
  LinkTarget="_self"
  call ImageList (SQL,HLink,LinkParam,LinkTarget)
%>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->