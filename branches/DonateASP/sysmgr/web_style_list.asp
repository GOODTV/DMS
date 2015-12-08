<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>系統參數設定</title>
</head>
<body class=tool>
<%
If request("SQL")="" Then 
  SQL="Select Ser_No,類別=Style_Type,控制項=Style_Name,控制項說明=Style_Desc,設定值=Style_Value,排序=Style_Seq " & _
      "From WEB_STYLE Where 1=1"
  If request("Style_Type")<>"" Then SQL=SQL & "And Style_Type = '"&request("Style_Type")&"' "
  If request("KeyWord")<>"" Then SQL=SQL & "And (Style_Name Like '%"&request("KeyWord")&"%' Or Style_Desc Like '%"&request("KeyWord")&"%' Or Style_Value Like '%"&request("KeyWord")&"%') "   
  SQL=SQL&" Order By Style_Type,Style_Seq,Ser_No Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="web_style_list"
HLink="web_style_edit.asp?ser_no="
LinkParam="ser_no"
LinkTarget="main"
AddLink="web_style_add.asp"
Hit_Nemu=""
call GridListHit (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,Hit_Nemu)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->