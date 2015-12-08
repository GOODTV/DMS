<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="act"%>
<!--#include file="../include/head.asp"-->
<body class=tool>
<%
If request("SQL")="" Then 
  SQL="Select Act_Id,機構名稱=Comp_ShortName,活動簡稱=Act_ShortName,主辦單位=Act_OrgName,協辦單位=Act_OrgName2,活動主題=Act_Subject,活動起日=Act_BeginDate,活動迄日=Act_EndDate From ACT Join Dept On ACT.Dept_Id=Dept.Dept_Id Where Dept.Dept_Id In ("&Session("all_dept_type")&")"
  If request("KeyWords")<>"" Then SQL=SQL & "And (Act_Name Like '%"&request("KeyWords")&"%' Or Act_ShortName Like '%"&request("KeyWords")&"%' Or Act_OrgName Like '%"&request("KeyWords")&"%' Or Act_Subject Like '%"&request("KeyWords")&"%') "
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
HLink="act_edit.asp?act_id="
LinkParam="act_id"
LinkTarget="main"
AddLink="act_add.asp"
call GridList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->