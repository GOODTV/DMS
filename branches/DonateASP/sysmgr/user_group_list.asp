<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>帳號權限設定</title>
</head>
<body class="tool">
<%
SQL="select usergroup=user_group,user_group,帳號類別=user_group,安全控管=security from groupfile"
HLink="user_group_edit.asp?usergroup="
LinkParam="usergroup"
LinkTarget="main"
DataLink="user_group_delete.asp?usergroup="
DataParam="usergroup"
DataTarget="main"
call DataLinkGrid (SQL,HLink,LinkParam,LinkTarget,DataLink,DataParam,DataTarget)
%>
</body>

</html>
<!--#include file="../include/dbclose.asp"-->