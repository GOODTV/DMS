<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="donate_sale"%>
<!--#include file="../include/head.asp"-->
<body class=tool>
<%
If request("SQL")="" Then 
  If request("action")="query" Then
    SQL="Select Ser_No,標題=Sale_Subject,列印起日=Sale_BeginDate,列印迄日=Sale_EndDate From DONATE_SALE Where Sale_EndDate>='"&Date()&"' "
  Else
    SQL="Select Ser_No,標題=Sale_Subject,列印起日=Sale_BeginDate,列印迄日=Sale_EndDate From DONATE_SALE Where Sale_EndDate<'"&Date()&"' "
  End If
  If request("Dept_Id")<>"" Then
    SQL=SQL&"And Dept_Id = '"&request("Dept_Id")&"' "
  Else
    SQL=SQL&"And Dept_Id In ("&Session("all_dept_type")&") "
  End If
  If request("KeyWord")<>"" Then SQL=SQL & "And (Sale_Subject Like '%"&request("KeyWord")&"%' Or Sale_Content Like '%"&request("KeyWord")&"%') "
  SQL=SQL&"Order By Sale_BeginDate Desc,Sale_EndDate Desc,Ser_No Desc"  
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="donate_sale_list"
HLink="donate_sale_edit.asp?ser_no="
LinkParam="ser_no"
LinkTarget="main"
AddLink="donate_sale_add.asp"
Hit_Nemu=""
call GridListHit (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,Hit_Nemu)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->