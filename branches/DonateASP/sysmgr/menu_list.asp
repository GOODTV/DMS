<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>程式群組管理</title>
<link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>

<body class=tool>

<%
if request("SQL")="" then
   SQL="select menu_id,程式群組=menu_id,menu_seq as 排序,URL=menu_url,程式群組說明=menu_desc from menu order by menu_seq"
else
   SQL=request("SQL")   
end if
PageSize=18
if request("nowPage")="" then
   nowPage=1
else
   nowPage=request("nowPage")
end if
ProgID="menu_list"
HLink="menu_edit.asp?menu_id="
LinkParam="menu_id"
LinkTarget="main"
AddLink="menu_Add.asp"
call PageList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
%>

</body>

</html>
<!--#include file="../include/dbclose.asp"-->