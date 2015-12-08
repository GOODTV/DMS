<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function IepayList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Donate_Amt=0
  SQL1="Select Donate_Amount=Sum(Donate_Amount) From "&Split(Split(SQL,"From")(1),"Order By")(0)
  'response.write SQL1
  'response.end
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  
  If Not RS1.EOF Then Donate_Amt=RS1("Donate_Amount")
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
        If J=3 Then
          Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(RS1(J),0)&"</span></td>"
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
<%Prog_Id="iepay"%>
<!--#include file="../include/head.asp"-->
<body class=tool>
<%
'page_load不帶全部資料
If request("action") = "" and request("SQL")="" Then
	Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
  Response.Write "  <tr>"
  Response.Write "    <td width='100%' align='center' style='color:#ff0000'>** 請先輸入查詢條件 **</td>"	  
  Response.Write "  </tr>"
  Response.Write "</table>"  
ElseIf request("SQL")="" Then 
  SQL="select od_sob,訂單編號=od_sob,捐款日期=CONVERT(VarChar,Donate_CreateDate,111),捐款金額=Donate_Amount, " & _
  	  "付款方式=(CASE paytype when '1' then '信用卡' when '2' then 'iePay儲值帳戶付款' when '4' then 'PayPal' when '5' then '其他超商 電子帳單' when '8' then 'Web ATM' when '9' then '7-11 電子帳單' when '12' then '玉山銀行eCoin' when '16' then '郵局電子帳單付款' when '17' then '郵局ATM付款' when '25' then '24hr超商取貨付款' when '30' then '7-11 ibon' when '35' then '條碼' when '39' then '全家FamilyPort' end), " & _
	  "編號=W.Donor_Id, 姓名=Donate_DonorName, " & _
	  "地址=(Case When D.IsAbroad = 'Y' or W.Donate_CityCode='' then Donate_Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Donate_Address Else B.mValue+Donate_Address End End), " & _
	  "電話=Donate_CellPhone,捐款用途=Donate_Purpose, " & _
	  "狀態=CASE when Status='0' then '授權成功' when Status IS NULL then '未完成刷卡流程' else '授權失敗' end  " & _
      "From DONATE_WEB W left join DONATE_IEPAY P on P.orderid=W.od_sob Left Join CODECITY As B On W.Donate_CityCode=B.mCode Left Join CODECITY As C On W.Donate_AreaCode=C.mCode left join DONOR D on W.Donor_Id=D.Donor_Id "
  If request("Donate_Date_B")<>"" Then SQL=SQL&"Where Donate_CreateDate >= '"&request("Donate_Date_B")&"' "
  If request("Donate_Date_E")<>"" Then SQL=SQL&"And Donate_CreateDate <= '"&request("Donate_Date_E")&"' "
  If request("Status")<>"" Then 
	Select Case request("Status") 
		Case "0"
		SQL=SQL&"And Status = '0' " '授權成功
		Case "1"
		SQL=SQL&"And Status <> '0' " '授權失敗
		Case Else 
		SQL=SQL&"And Status is null "  '未完成刷卡流程
	End Select
  End If  
  SQL=SQL&"Order By CONVERT(VarChar,Donate_CreateDate,111) Desc,od_sob Desc"
Else
  SQL=request("SQL")
End If
'response.write SQL
'response.end()
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="iepay_list"
HLink=""
LinkParam="donate_id"
LinkTarget="main"
AddLink=""
If request("action")="stop" Then
  call GridList_S (AddLink)
Else
  If request("action")="report" Or request("action")="export" Then Server.Transfer "iepay_rpt.asp"
  If request("action")<>"" Or SQL<>"" Then call IepayList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
End If  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->