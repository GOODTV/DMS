<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function ContributeList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Contribute_Amt=0
  SQL1="Select Contribute_Amt=Isnull(Sum(Contribute_Amt),0) From "&Split(Split(SQL,"From")(1),"Order By")(0)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If Not RS1.EOF Then Contribute_Amt=RS1("Contribute_Amt")
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
    Response.Write "<tr><td width='30%'>折合現金合計："&FormatNumber(Contribute_Amt,0)&" 元</td>"
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
        If J=4 Then
          Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(J),0) & "</span></td>"
        Else
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & Data_Minus(RS1(J)) & "</span></td>"
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
<%Prog_Id="contribute"%>
<!--#include file="../include/head.asp"-->
<body class=tool>
<%
If request("SQL")="" Then 
  SQL="Select Contribute_id,編號=(Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),CONTRIBUTE.Donor_Id) End),捐贈人=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),捐贈日期=CONVERT(nvarchar,Contribute_Date,111),折合現金=Isnull(Contribute_Amt,0),捐款方式=Contribute_Payment,捐款用途=Contribute_Purpose,收據開立=CONTRIBUTE.Invoice_Type,收據編號=Invoice_Pre+Invoice_No,列印=(Case When Invoice_Print='1' Then 'V' Else '' End),狀態=(Case When Issue_Type='M' Then '手開' Else Case When Issue_Type='D' Then '作廢' Else '' End End),經手人=CONTRIBUTE.Create_User " & _
      "From CONTRIBUTE Join DONOR On CONTRIBUTE.Donor_Id=DONOR.Donor_Id Left Join ACT On CONTRIBUTE.Act_Id=ACT.Act_Id "
  If request("Dept_Id")<>"" Then
    SQL=SQL&"Where CONTRIBUTE.Dept_Id = '"&request("Dept_Id")&"' "
  Else
    SQL=SQL&"Where CONTRIBUTE.Dept_Id In ("&Session("all_dept_type")&") "
  End If
  If request("Donor_Name")<>"" Then SQL=SQL&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or CONTRIBUTE.Invoice_Title Like '%"&request("Donor_Name")&"%') "
  If request("Member_No")<>"" Then SQL=SQL & "And DONOR.Donor_Id Like '%"&request("Member_No")&"%' "
  If request("Contribute_Payment")<>"" Then SQL=SQL&"And Contribute_Payment = '"&request("Contribute_Payment")&"' "
  If request("Contribute_Purpose")<>"" Then SQL=SQL&"And Contribute_Purpose = '"&request("Contribute_Purpose")&"' "
  If request("Contribute_Date_B")<>"" Then SQL=SQL&"And Contribute_Date >= '"&request("Contribute_Date_B")&"' "
  If request("Contribute_Date_E")<>"" Then SQL=SQL&"And Contribute_Date <= '"&request("Contribute_Date_E")&"' "
  If request("Invoice_Type")<>"" Then SQL=SQL&"And CONTRIBUTE.Invoice_Type = '"&request("Invoice_Type")&"' "
  If request("Invoice_No_B")<>"" Then SQL=SQL&"And Invoice_No >= '"&request("Invoice_No_B")&"' "
  If request("Invoice_No_E")<>"" Then SQL=SQL&"And Invoice_No <= '"&request("Invoice_No_E")&"' "
  If request("Accoun_Date_B")<>"" Then SQL=SQL&"And Accoun_Date >= '"&request("Accoun_Date_B")&"' "
  If request("Accoun_Date_E")<>"" Then SQL=SQL&"And Accoun_Date <= '"&request("Accoun_Date_E")&"' "
  If request("Issue_TypeM")<>"" Then SQL=SQL&"And Issue_Type = '"&request("Issue_TypeM")&"' "
  If request("Issue_TypeD")<>"" Then SQL=SQL&"And Issue_Type = '"&request("Issue_TypeD")&"' "
  If request("Accounting_Title")<>"" Then SQL=SQL&"And Accounting_Title = '"&request("Accounting_Title")&"' "
  If request("Act_Id")<>"" Then SQL=SQL&"And Act_Id = '"&request("Act_Id")&"' "
  If request("Create_User")<>"" Then SQL=SQL&"And CONTRIBUTE.Create_User = '"&request("Create_User")&"' "
  SQL=SQL&"Order By CONVERT(VarChar,Contribute_Date,111) Desc,Contribute_id Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="contribute_list"
HLink="contribute_detail.asp?contribute_id="
LinkParam="contribute_id"
LinkTarget="main"
AddLink="contribute_input.asp"
If request("action")="stop" Then
  call GridList_S (AddLink)
Else
  If request("action")="report" Or request("action")="export" Then Server.Transfer "contribute_rpt.asp"
  call ContributeList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
End If  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->