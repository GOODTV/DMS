<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=ecbank.xls"
End If

Function DonorList (SQL,ReportName)
  PrintDate=Year(Now())&"/"&Left("00",2-Len(Month(Now())))&Month(Now())&"/"&Left("00",2-Len(Day(Now())))&Day(Now())&" "&Left("00",2-Len(Hour(Now())))&Hour(Now())&":"&Left("00",2-Len(Minute(Now())))&Minute(Now())&":"&Left("00",2-Len(Second(Now())))&Second(Now())
  Response.Write "<table id='headgrid' border='0' cellspacing='2' cellpadding='3' style='border-collapse: collapse;' bordercolor='#111111' width='98%' align='center'>"
  If request("action")="export" Then
    Response.Write "  <tr>"
    Response.Write "    <td width='100%' align='center' style='font-size: 13pt;font-family: 標楷體' colspan='34'>"&ReportName&"</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td align='left' style='font-size: 10pt;font-family: 標楷體' colspan='34'>"
    Response.Write "匯出日期："&PrintDate
    Response.Write "    </td>"
    Response.Write "  </tr>"
  Else
    Response.Write "  <tr>"
    Response.Write "    <td width='100%' align='center' style='font-size: 13pt;font-family: 標楷體' colspan='2'>"&ReportName&"</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td width='70%' align='left' style='font-size: 10pt;font-family: 標楷體'>"
    Response.Write "    </td>"
    Response.Write "    <td width='30%' align='right' style='font-size: 10pt;font-family: 標楷體'>列印日期："&PrintDate&"</td>"
    Response.Write "  </tr>"
  End If
  Response.Write "</table>"
  
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,3
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id='datagrid' border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='98%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount-1
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For J = 1 To FieldsCount-1
      If J=1 Then
        Donate_Type=RS1(2)
        od_sob=RS1(4)
      End If
      If J=2 Then
        If Donate_Type="webatm" Then
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>Web-ATM</span></td>" 
        ElseIf Donate_Type="vacc" Then
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>銀行虛擬帳號</span></td>" 
        ElseIf Donate_Type="barcode" Then
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>超商條碼</span></td>" 
        ElseIf Donate_Type="ibon" Then
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>7-11 ibon</span></td>" 
        ElseIf Donate_Type="famiport" Then
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>全家 Famiport</span></td>" 
        ElseIf Donate_Type="lifeet" Then
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>萊爾富 Life-Et</span></td>"
        ElseIf Donate_Type="okgo" Then
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>OK超商 OK-GO</span></td>"           
        Else
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
        End If
      ElseIf J=5 Then
        If RS1(J)<>"" Then
          Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(RS1(J),0)&"</span></td>"
        Else
          Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(J))&"</span></td>"
        End If
      ElseIf J=6 Then
        expire_datetime=Cstr(RS1(6))
        expire_datetime=Left(Replace(Replace(Replace(expire_datetime,"-","")," ",""),":",""),12)
        expire_date_time=""
        If expire_datetime<>"" Then expire_date_time=Left(expire_datetime,4)&"/"&Mid(expire_datetime,5,2)&"/"&Mid(expire_datetime,7,2)&" "&Mid(expire_datetime,9,2)&":"&Mid(expire_datetime,11,2)
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&expire_date_time&"</span></td>"
      ElseIf J=8 Then
        Response.Write "<td x:str><span style='font-size: 9pt; font-family: 新細明體'>"&RS1(J)&"</span></td>"      
      Else
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(J))&"</span></td>"
      End If
	  Next
	  RS1("Donate_Export")="Y"
	  RS1.Update
	  Response.Flush
    Response.Clear
    RS1.MoveNext
	  Response.Write "</tr>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function 
%>
<%Prog_Id="ecbank_qry"%>
<!--#include file="../include/head_rpt.asp"-->
<body class=tool>
<%
  SQL="Select A.Ser_No,捐款人=A.Donate_DonorName,交易方式=A.Donate_Type,交易日期=A.Donate_CreateDateTime,交易編號=A.od_sob,交易金額=(Case When CONVERT(numeric,D.amt)>0 Then CONVERT(numeric,D.amt) Else CONVERT(numeric,A.Donate_Amount) End)," & _
      "繳費截止日期=(Case When A.Donate_Type='webatm' Then CONVERT(nvarchar,Donate_CreateDateTime,120) Else Case When D.get_expire_date+D.get_expire_time<>'' Then D.get_expire_date+D.get_expire_time Else '' End End)," & _
      "繳費狀態=(Case When D.close_type='ok' Then '付款完成' Else '' End), " & _
      "手機=Donate_CellPhone,聯絡電話=Donate_TelOffice,電子郵件=Donate_Email,捐款用途=Donate_Purpose,收據開立=Donate_Invoice_Type,收據抬頭=Donate_Invoice_Title, " & _
      "收據地址=(Case When A.Donate_Invoice_CityCode='' Then Donate_Invoice_Address Else Case When B.mValue<>C.mValue Then B.mValue+A.Donate_Invoice_ZipCode+C.mValue+A.Donate_Invoice_Address Else B.mValue+A.Donate_Invoice_ZipCode+Donate_Invoice_Address End End),  " & _
      "通訊地址=(Case When A.Donate_CityCode='' Then Donate_Address Else Case When F.mValue<>G.mValue Then F.mValue+A.Donate_ZipCode+G.mValue+A.Donate_Address Else F.mValue+A.Donate_ZipCode+Donate_Address End End), " & _
      "性別=Donate_Sex,身分證=Donate_IDNO,出生日期=Donate_Birthday,教育程度=Donate_Education,職業=Donate_Occupation,Donate_Export " & _
      "From DONATE_WEB A " & _
      "Left Join CODECITY As B On A.Donate_Invoice_CityCode=B.mCode " & _
      "Left Join CODECITY As C On A.Donate_Invoice_AreaCode=C.mCode " & _
      "Left Join CODECITY As F On A.Donate_CityCode=F.mCode " & _
      "Left Join CODECITY As G On A.Donate_AreaCode=G.mCode " & _
      "Left Join DONATE_ECBANK As D On A.od_sob=D.od_sob And A.Donate_Type<>'creditcard' "
  If Request("Donate_DonorName")<>"" Then SQL=SQL&"And Donate_DonorName Like '%"&Request("Donate_DonorName")&"%' "
  If Request("Donate_Purpose")<>"" Then SQL=SQL&"And Donate_Purpose = '"&Request("Donate_Purpose")&"' "
  If Request("Donate_Invoice_Type")<>"" Then SQL=SQL&"And Donate_Invoice_Type = '"&Request("Donate_Invoice_Type")&"' "
  If Request("Donate_CreateDate_B")<>"" Then SQL=SQL&"And Donate_CreateDate >= '"&Request("Donate_CreateDate_B")&"' "
  If Request("Donate_CreateDate_E")<>"" Then SQL=SQL&"And Donate_CreateDate <= '"&Request("Donate_CreateDate_E")&"' "
  If Request("Donate_Type")<>"" Then SQL=SQL&"And Donate_Type = '"&Request("Donate_Type")&"' "
  If Request("Donate_Export")<>"" Then SQL=SQL&"And Donate_Export = '"&Request("Donate_Export")&"' "
  If Request("Close_Type")="L" Then SQL=SQL&"And (Case When A.Donate_Type='webatm' Then CONVERT(nvarchar,Donate_CreateDateTime,120) Else Case When D.get_expire_date+D.get_expire_time<>'' Then D.get_expire_date+D.get_expire_time Else '' End End)='' "
  If Request("Close_Type")="N" Then SQL=SQL&"And (Close_Type Is null Or Close_Type='') "
  If Request("Close_Type")="Y" Then SQL=SQL&"And (Close_Type = 'ok') "
  SQL=SQL&"Order By Donate_CreateDateTime Desc,A.Ser_No Desc"
  ReportName="ECBANK交易紀錄"
  call DonorList (SQL,ReportName)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->