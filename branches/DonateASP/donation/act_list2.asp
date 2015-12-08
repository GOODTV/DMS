<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>勸募活動維護</title>
  <%If request("action")<>"export" Then%><link REL="stylesheet" type="text/css" HREF="../include/dms.css"><%End If%>
</head>
<body class=tool>
<%
If request("SQL")="" Then
  SQL="Select Act_Id,核心價值=Core_Value,主責單位=Comp_ShortName,主責人=Act_MainPower,活動簡稱=Act_ShortName,活動主題=Act_Subject,活動起日=Act_BeginDate,活動迄日=Act_EndDate,方案狀態=OutCome From ACT Join Dept On ACT.Dept_Id=Dept.Dept_Id "
  If request("KeyWords")<>"" Then SQL=SQL & "Where (Act_Name Like '%"&request("KeyWords")&"%' Or Act_ShortName Like '%"&request("KeyWords")&"%' Or Act_OrgName Like '%"&request("KeyWords")&"%' Or Act_Subject Like '%"&request("KeyWords")&"%') "
  SQL=SQL&"Order By Act.Dept_Id,Act_Id Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="act_list"
HLink="act_edit2.asp?act_id="
LinkParam="act_id"
LinkTarget="main"
AddLink="act_add2.asp"
call GridList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->