<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="licence"%>
<!--#include file="../include/head.asp"-->
<body class=tool>
<%
If request("SQL")="" Then 
  SQL="Select Ser_No,機構名稱=Comp_ShortName,勸募文號=Licence_No,使用起日=Licence_BeginDate,使用迄日=Licence_EndDate,預設=(Case When Licence_Default='Y' Then 'V' Else '' End) From DONATE_LICENCE Join Dept On DONATE_LICENCE.Dept_Id=Dept.Dept_Id Where Dept.Dept_Id In ("&Session("all_dept_type")&") "
  If request("KeyWords")<>"" Then SQL=SQL & "And (Licence_No Like '%"&request("KeyWords")&"%') "
  SQL=SQL&"Order By DONATE_LICENCE.Dept_Id,Ser_No Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="licence_list"
HLink="licence_edit.asp?ser_no="
LinkParam="ser_no"
LinkTarget="main"
AddLink="licence_add.asp"
call GridList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->