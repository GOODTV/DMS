<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>郵件管理</title>
</head>
<body class=tool>
<%
If request("SQL")="" then
  SQL="Select Ser_No,郵件標題=EmailMgr_Subject,建立日期=EmailMgr_RegDate From EMAILMGR Where EmailMgr_Default<>'Y' "
  If request("KeyWord")<>"" Then SQL=SQL & "And (EmailMgr_Subject Like '%"&request("KeyWord")&"%') "   
  SQL=SQL&" Order By EmailMgr_RegDate Desc,Ser_No Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="emailmgr_list"
HLink="emailmgr_edit.asp?ser_no="
LinkParam="ser_no"
LinkTarget="main"
AddLink="emailmgr_add.asp"
Hit_Nemu=""
call GridListHit (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,Hit_Nemu)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->