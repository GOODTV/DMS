<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=iepay.xls"
End If

Function IepayList (SQL,ReportName)
  Donate_Amt=0
  SQL1="Select Donate_Amount=Sum(Donate_Amount) From DONATE_WEB W left join DONATE_IEPAY P on P.orderid=W.od_sob Where"&Split(Split(SQL,"Where")(1),"Order By")(0)
  'response.write SQL1
  'response.end()
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If Not RS1.EOF Then Donate_Amt=RS1("Donate_Amount")
  RS1.Close
  Set RS1=Nothing
  Response.Write "<center><span style='font-size: 12pt; font-family: 標楷體'>"&ReportName&"</span></center>"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='98%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For I = 1 To FieldsCount
      If I=3 Then
        Response.Write "<td bgcolor='#FFFFFF' align=""right""><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(RS1(I),0)&"</span></td>"
      Else
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(I))&"</span></td>"
	  End If
	  Next
	  Response.Flush
    Response.Clear
    RS1.MoveNext
	  Response.Write "</tr>"
  Wend
  Response.Write "<tr>"
  Response.Write "  <td colspan=""5"" align=""right"">捐款合計："&FormatNumber(Donate_Amt,0)&" 元</td>"
  Response.Write "  <td colspan="""&FieldsCount-5&"""></td>"
  Response.Write "</tr>"	
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function 
%>
<%Prog_Id="iepay"%>
<!--#include file="../include/head_rpt.asp"-->
<style>
td{mso-number-format:\@;}
</style>
<body class=tool>
<%
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
  'response.write SQL
  'response.end()
  call IepayList (SQL,Prog_Desc)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->