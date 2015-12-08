<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function DonateList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Donate_Amt=0
  SQL1="Select Donate_Amt=Sum(Donate_Amt) From "&Split(Split(SQL,"From")(1),"Order By")(0)
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
    Response.Write "<tr><td width='30%'>捐款金額合計："&FormatNumber(Donate_Amt,0)&" 元</td>"
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
        If J=2 Then
          Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(J),0) & "</span></td>"
        Else
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(J))&"</span></td>"
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
<%Prog_Desc="捐款記錄"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link rel="stylesheet" type="text/css" href="../include/dms.css">
  <title><%=Prog_Desc%></title>
</head>
<body class=tool>
<%
If request("SQL")="" Then 
	'20130923 Modify by GoodTV Tanya:調整舊收據編號無Invoice_Pre問題
  SQL="Select Donate_Id,捐款日期=CONVERT(VarChar,Donate_Date,111),捐款金額=Donate_Amt,捐款方式=Donate_Payment,捐款用途=Donate_Purpose,收據開立=DONATE.Invoice_Type,收據編號=IsNull(Invoice_Pre,'')+Invoice_No,收據抬頭=DONATE.Invoice_Title,列印=(Case When Invoice_Print='1' Or Invoice_Print_Yearly='1' Then 'V' Else '' End),狀態=(Case When Issue_Type='M' Then '手開' Else Case When Issue_Type='D' Then '作廢' Else '' End End),經手人=DONATE.Create_User " & _
      "From DONATE Where Donor_Id='"&request("Donor_Id")&"' " & _
      "Order By CONVERT(VarChar,Donate_Date,111) Desc,Donate_Id Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="donate_data_list"
HLink="donate_detail.asp?ctype=donate_data&donate_id="
LinkParam="donate_id"
LinkTarget="main"
AddLink="donate_add.asp?ctype=donate_data&donor_id="&request("Donor_Id")
call DonateList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink) 
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->