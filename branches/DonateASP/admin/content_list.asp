<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>網站內容管理</title>
</head>
<body class=tool>
<%
If request("SQL")="" Then 
  SQL="Select Ser_No,主功能=Content_Type,頁面=Content_SubType,網頁標題=Content_Subject,發佈日期=Content_RegDate " & _
      "From CONTENT Join CodeFile On CONTENT.Content_Type=CodeFile.CodeName Where CodeFile.Category='單一網頁' "
  If Session("user_group")<>"系統管理員" And Session("BranchType")<>"" Then SQL=SQL & "And BranchType = '"&Session("BranchType")&"' "
  If request("Content_Type")<>"" Then SQL=SQL & "And Content_Type = '"&request("Content_Type")&"' "
  If request("KeyWord")<>"" Then SQL=SQL & "And (Content_Subject Like '%"&request("KeyWord")&"%' Or Content_Brief Like '%"&request("KeyWord")&"%' Or Content_Desc Like '%"&request("KeyWord")&"%') "   
  SQL=SQL&" Order By CodeFile.Menu_Seq,CONTENT.Content_Type,CONTENT.Ser_No"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="content_list"
HLink="content_edit.asp?ser_no="
LinkParam="ser_no"
LinkTarget="main"
AddLink="content_add.asp"
Hit_Nemu="content"
call GridListHit (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,Hit_Nemu)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->