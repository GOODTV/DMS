<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>使用者管理</title>
</head>
<body class=tool>
<%
if request("SQL")="" then
   SQL="select uid=userfile.dept_id+user_id,帳號=user_id,帳號名稱=user_name,部門=dept_desc,帳號群組=user_group,建檔日期=create_date from userfile"_
+"      join dept on userfile.dept_id=dept.dept_id where"_
+"     ('"&request("dept_id")&"'='' or userfile.dept_id='"&request("dept_id")&"') and"
If Session("user_id")<>"npois" Then SQL=SQL&" user_id<>'npois' And "
SQL=SQL&" ('"&request("user_group")&"'='' or user_group='"&request("user_group")&"') order by user_id"
else
   SQL=request("SQL")   
end if
PageSize=18
if request("nowPage")="" then
   nowPage=1
else
   nowPage=request("nowPage")
end if
ProgID="userfile_list"
HLink="userfile_edit.asp?uid="
LinkParam="uid"
LinkTarget="main"
AddLink="userfile_Add.asp"
call PageList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
%>
</body>

</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->