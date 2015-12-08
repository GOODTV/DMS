<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function GridListHit_Signup (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,Hit_Nemu)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  If Not RS1.EOF Then 
    FieldsCount = RS1.Fields.Count-1
    totRec=RS1.Recordcount
    If totRec>0 Then 
      RS1.PageSize=PageSize
      If nowPage="" or nowPage=0 Then 
        nowPage=1
      ElseIf cint(nowPage) > RS1.PageCount Then 
        nowPage=RS1.PageCount 
      End If
      session("nowPage")=nowPage
      RS1.AbsolutePage=nowPage
      totPage=RS1.PageCount
      SQL=server.URLEncode(SQL)
    End If
    Response.Write "<Form action='' id=form1 name=form1>" 
    Response.Write "<table border=0 cellspacing='0' cellpadding='1' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='30%'></td>"
    Response.Write "<td width='55%'><span style='font-size: 9pt; font-family: 新細明體'>第" & nowPage & "/" & totPage & "頁&nbsp;&nbsp;</span>"
    Response.Write "<span style='font-size: 9pt; font-family: 新細明體'>共" & FormatNumber(totRec,0) & "筆&nbsp;&nbsp;</span>"
    If cint(nowPage) <>1 Then             
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage-1)&"&SQL="&SQL&"'>上一頁</a>" 
    End If      
    If cint(nowPage)<>RS1.PageCount and cint(nowPage)<RS1.PageCount Then 
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage+1)&"&SQL="&SQL&"'>下一頁</a>" 
    End If
    Response.Write " |&nbsp;<span style='font-size: 9pt; font-family: 新細明體'> 跳至第<select name=GoPage size='1' style='font-size: 9pt; font-family: 新細明體' onchange='GoPage_OnChange(this.value)'>"
    For iPage=1 to totPage
      If iPage=cint(nowPage) Then
        strSelected = "selected"
      Else
	      strSelected = "" 
      End If   
      Response.Write "<option value='"&iPage&"'" & strSelected & ">" & iPage & "</option>"          
    Next   
    Response.Write "</select>頁</span></td>" 
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增報名資料</a></span></td>"
    End If   
    Response.Write "</tr></table>"
    Dim I
    Dim J
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For J = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J).Name & "</span></font></td>"
    Next
    Response.Write "</tr>"
    I = 1
    While Not RS1.EOF And I <= RS1.PageSize
      If Hlink<>"" Then
        Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
        showhand="style='cursor:hand'"
      End If
      Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
      For J = 1 To FieldsCount
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
      Next
      I = I + 1
      RS1.MoveNext
      Response.Write "</tr>"
      If Hit_Nemu<>"activity" Then Response.Write "</a>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  Else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>"
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增報名資料</a></span></td>"
    End If  
    Response.Write "</table>"
  End If 
  RS1.Close
  Set RS1=Nothing
End Function

SQL1="Select Member_Act_IsFood,Member_Act_IsPrice From MEMBER_ACT Where Member_Act_Id='"&request("member_act_id")&"'"
call QuerySQL(SQL1,RS1)
Member_Act_IsFood=RS1("Member_Act_IsFood")
Member_Act_IsPrice=RS1("Member_Act_IsPrice")
RS1.Close
Set RS1=Nothing
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>活動報名清單</title>
  <%If request("action")<>"export" Then%><link REL="stylesheet" type="text/css" HREF="../include/dms.css"><%End If%>
</head>
<body class=tool>
<%
If request("SQL")="" Then
  SQL="Select Ser_No,報名日期=Member_Signup_DateTime,編號=(Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),DONOR.Donor_Id) End),姓名=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),聯絡電話=(Case When Cellular_Phone<>'' Then Cellular_Phone Else Case When Tel_Office<>'' Then Tel_Office Else Tel_Home End End),會員否=(Case When IsMember='Y' Then 'V' Else '' End),報名狀態=Member_Signup_Status "
  If Member_Act_IsFood="Y" Then SQL=SQL & ",飲食習慣=Member_Signup_Food "
  If Member_Act_IsPrice="Y" Then SQL=SQL & ",報名費用=Member_Signup_Price,繳費日期=Member_Signup_Pay "
  SQL=SQL & "From MEMBER_SIGNUP Join DONOR On MEMBER_SIGNUP.Donor_Id=DONOR.Donor_Id Where Member_Act_Id='"&request("member_Act_id")&"' "
  If request("Member_Signup_Date_B")<>"" Then SQL=SQL & "And Member_Signup_Date>='"&request("Member_Signup_Date_B")&"' "
  If request("Member_Signup_Date_E")<>"" Then SQL=SQL & "And Member_Signup_Date<='"&request("Member_Signup_Date_E")&"' "
  If request("Member_Signup_Status")<>"" Then SQL=SQL & "And Member_Signup_Status='"&request("Member_Signup_Status")&"' "
  If request("KeyWords")<>"" Then SQL=SQL & "And (Donor_Name Like '%"&request("KeyWords")&"%') "
  SQL=SQL&"Order By Member_Signup_DateTime Desc,Ser_No Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="member_signup_list"
If Member_Act_IsPrice="Y" Then
  HLink="member_signup_donate_edit.asp?member_act_id="&request("member_act_id")&"&ser_no="
Else
  HLink="member_signup_edit.asp?member_act_id="&request("member_act_id")&"&ser_no="
End If
LinkParam="ser_no"
LinkTarget="main"
AddLink="member_signup_input.asp?member_act_id="&request("member_act_id")
If request("action")="report" Or request("action")="export" Then Server.Transfer "member_signup_rpt.asp"
call GridListHit_Signup (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,"")
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->