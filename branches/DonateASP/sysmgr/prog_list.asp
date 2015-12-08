<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>程式群組管理</title>
</head>
<body class=tool>

<%
if request("SQL")="" then
   SQL="select prog_id,程式代號=prog_id,程式群組=menu_id,URL=prog_url,程式名稱=prog_desc,prog_seq as 排序,程式類別=prog_type from prog left join prog_menu"_
+"     on prog.prog_id=prog_menu.progid where ('"&request("menu_id")&"'='' or menu_id='"&request("menu_id")&"') and"_
+"     ('"&request("prog_desc")&"'='' or prog_desc like '%"&request("prog_desc")&"%')"_
+"     order by (select menu_seq from menu where prog_menu.menu_id=menu.menu_id),prog_seq"
else
   SQL=request("SQL")   
end if
PageSize=18
if request("nowPage")="" then
   nowPage=1
else
   nowPage=request("nowPage")
end if
ProgID="prog_list"
HLink="prog_edit.asp?prog_id="
LinkParam="prog_id"
LinkTarget="main"
AddLink="prog_Add.asp"
call PageList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
%>

</body>

</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->