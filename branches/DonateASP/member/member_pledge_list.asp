<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function PledgeList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Donate_Amt=0
  SQL1="Select Donate_Amt=Sum(Donate_Amt) From "&Split(Split(SQL," From ")(1)," Order By ")(0)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If Not RS1.EOF Then Donate_Amt=RS1("Donate_Amt")
  RS1.Close
  Set RS1=Nothing
    
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
    Response.Write "<tr><td width='30%'>扣款金額合計："&FormatNumber(Donate_Amt,0)&" 元</td>"
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
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
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
        If J=5 Then
          Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(J),0) & "</span></td>"
        Else
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
        End If
      Next
      I = I + 1
      RS1.MoveNext
      Response.Write "</tr>"
      Response.Write "</a>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  Else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>"
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If
    Response.Write "</table>"
  End If
  RS1.Close
  Set RS1=Nothing
End Function
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>固定轉帳授權書維護</title>
  <%If request("action")<>"export" Then%><link REL="stylesheet" type="text/css" HREF="../include/dms.css"><%End If%>
</head>
<body class=tool>
<%
If request("SQL")="" Then 
  SQL="Select Pledge_Id,授權編號=Pledge_Id,會員姓名=DONOR.Donor_Name,授權方式=Donate_Payment,指定用途=Donate_Purpose,扣款金額=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),授權起日=CONVERT(VarChar,donate_fromdate,111),授權迄日=CONVERT(VarChar,donate_todate,111),轉帳週期=donate_period,下次扣款日期=CONVERT(VarChar,Next_DonateDate,111),信用卡有效月年=(Case When valid_date<>'' Then Left(valid_date,2)+'/'+Right(Valid_date,2) Else '' End),授權狀態=status,經手人=PLEDGE.Create_User " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id "
  If request("Dept_Id")<>"" Then
    SQL=SQL & "Where PLEDGE.Dept_Id = '"&request("Dept_Id")&"' "
  Else
    SQL=SQL & "Where PLEDGE.Dept_Id In ("&Session("all_dept_type")&") "
  End If
  If request("Donor_Name")<>"" Then SQL=SQL & "And (DONOR.Donor_Name Like '%"&request("Donor_Name")&"%' Or DONOR.NickName Like '%"&request("Donor_Name")&"%' Or DONOR.Contactor Like '%"&request("Donor_Name")&"%' Or DONOR.Invoice_Title Like '%"&request("Donor_Name")&"%') "
  If request("From_Year")<>"" Then SQL=SQL & "And Year(Donate_ToDate) = '"&request("From_Year")&"' "
  If request("From_Month")<>"" Then SQL=SQL & "And Month(Donate_ToDate) = '"&request("From_Month")&"' "
  If request("Donate_Payment")<>"" Then SQL=SQL & "And Donate_Payment = '"&request("Donate_Payment")&"' "
  If request("Donate_Purpose")<>"" Then SQL=SQL & "And Donate_Purpose = '"&request("Donate_Purpose")&"' "
  If request("Valid_Month")<>"" Then SQL=SQL & "And Substring((Valid_Date),1,2) = '"&request("Valid_Month")&"' "
  If request("Valid_Year")<>"" Then SQL=SQL & "And Substring((Valid_Date),3,2) = '"&Right(request("Valid_Year"),2)&"' "
  If request("Donate_Period")<>"" Then SQL=SQL & "And Donate_Period = '"&request("Donate_Period")&"' "
  If request("Donate_Date")<>"" Then SQL=SQL & "And Donate_FromDate <= '"&request("Donate_Date")&"' And Donate_ToDate>= '"&request("Donate_Date")&"' "
  If request("Status")<>"" Then SQL=SQL & "And Status = '"&request("Status")&"' "
  If request("Create_User")<>"" Then SQL=SQL & "And PLEDGE.Create_User = '"&request("Create_User")&"' "
  SQL=SQL&"Order By Pledge_Id Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="pledge_list"
HLink="pledge_edit.asp?pledge_id="
LinkParam="pledge_id"
LinkTarget="main"
AddLink="pledge_input.asp"
Session("SQL")=SQL
If request("action")="stop" Then
  call GridList_S (AddLink)
Else
  If request("action")="report" Then
    Response.Redirect "../include/movebar.asp?URL=../donation/pledge_rpt.asp"
  ElseIf request("action")="export" Then 
    Response.Redirect "pledge_rpt.asp?action=export"
  End if
  call PledgeList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
End If  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->