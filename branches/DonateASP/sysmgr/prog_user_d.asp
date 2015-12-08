<!--#include file="../_function/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>使用者程式設定</title>
</head>
<body class=tool>
<% 
SQL="select ser_no,程式群組,user_prog.User_ID,姓名=user_name,使用者群組 from user_prog join userfile on user_prog.user_id=userfile.user_id"_
+"     where prog_id='"&request("prog_id")&"' order by 程式群組,user_prog.User_ID" 
HLink="prog_user_delete.asp?ser_no=" 
LinkParam="ser_no"
LinkTarget="right"
call DataGrid(SQL,HLink,LinkParam,LinkTarget) 
%> 
</body> 
</html> 
<!--#include file="../_function/dbclose.asp"--> 
<!--#include file="../_function/vbfunction.inc"-->