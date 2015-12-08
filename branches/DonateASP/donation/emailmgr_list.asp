<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="emailmgr"%>
<!--#include file="../include/head.asp"-->
<body class=tool>
<%
If request("SQL")="" then
  SQL="Select EMAILMGR.Ser_No,郵件類別=EmailMgr_Type,郵件標題=EmailMgr_Subject " & _
      "From EMAILMGR Join CASECODE On EMAILMGR.EmailMgr_Type=CASECODE.CodeDesc Where CASECODE.CodeType='EmailMgrType' "
  If request("EmailMgr_Type")<>"" Then SQL=SQL & "And EmailMgr_Type = '"&request("EmailMgr_Type")&"' "
  If request("KeyWord")<>"" Then SQL=SQL & "And (EmailMgr_Subject Like '%"&request("KeyWord")&"%'Or EmailMgr_Desc Like '%"&request("KeyWord")&"%') "   
  SQL=SQL&" Order By CASECODE.Seq,EmailMgr_RegDate Desc,EMAILMGR.Ser_No Desc"
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
AddLink="emailmgr_add.asp?code_id="&request("code_id")
Hit_Nemu=""
call GridListHit (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,Hit_Nemu)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->