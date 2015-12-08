<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>電子印鑑管理</title>
<link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>

<body class=tool>
<%
If request("SQL")="" Then
   SQL="Select ser_no,印鑑標題=Seal_Subject,檔名=Seal_TitleImg From SEAL Where Ser_No>1 "
   If request("Seal_Subject")<>"" Then SQL=SQL&"And (Seal_Subject Like '%"&request("Seal_Subject")&"%') "
   SQL=SQL&"Order By ser_no"
Else
   SQL=request("SQL")
End If

PageSize=18
If request("nowPage")="" then
  nowPage=1
Else
  nowPage=request("nowPage")
End if
Session("SQL")=SQL
ProgID="seal_list"
HLink="seal_edit.asp?ser_no="
LinkParam="ser_no"
LinkTarget="main"
AddLink="seal_add.asp"
call PageList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)   
%>
</body>

</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->