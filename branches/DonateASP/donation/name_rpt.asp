<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=name.xls"
End If

Function DonorList (SQL,ReportName)
  Act_ShortName=""
  If Session("Act_Id")<>"" Then
    SQL1="Select Act_ShortName From ACT Where Act_Id='"&Session("Act_Id")&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    If Not RS1.EOF Then Act_ShortName=RS1("Act_ShortName")
    RS1.Close
    Set RS1=Nothing
  End If
  
  PrintDate=Year(Now())&"/"&Left("00",2-Len(Month(Now())))&Month(Now())&"/"&Left("00",2-Len(Day(Now())))&Day(Now())&" "&Left("00",2-Len(Hour(Now())))&Hour(Now())&":"&Left("00",2-Len(Minute(Now())))&Minute(Now())&":"&Left("00",2-Len(Second(Now())))&Second(Now())
  Response.Write "<table id='headgrid' border='0' cellspacing='2' cellpadding='3' style='border-collapse: collapse;' bordercolor='#111111' width='98%' align='center'>"
  If request("action")="export" Then
    Response.Write "  <tr>"
    Response.Write "    <td width='100%' align='center' style='font-size: 13pt;font-family: 標楷體' colspan='34'>"&ReportName&"</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td align='left' style='font-size: 10pt;font-family: 標楷體' colspan='34'>"
    If Session("Donate_Date_Begin")<>"" Or Session("Donate_Date_End")<>"" Then Response.Write "捐款日期："&Session("Donate_Date_Begin")&" ~ "&Session("Donate_Date_End")&"&nbsp;&nbsp;&nbsp;"
    If Act_ShortName<>"" Then Response.Write "募款活動："&Act_ShortName&"nbsp;&nbsp;&nbsp;"
    Response.Write "匯出日期："&PrintDate
    Response.Write "    </td>"
    Response.Write "  </tr>"
  Else
    Response.Write "  <tr>"
    Response.Write "    <td width='100%' align='center' style='font-size: 13pt;font-family: 標楷體' colspan='2'>"&ReportName&"</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td width='70%' align='left' style='font-size: 10pt;font-family: 標楷體'>"
    If Session("Donate_Date_Begin")<>"" Or Session("Donate_Date_End")<>"" Then Response.Write "捐款日期："&Session("Donate_Date_Begin")&" ~ "&Session("Donate_Date_End")&"&nbsp;&nbsp;&nbsp;"
    If Act_ShortName<>"" Then Response.Write "募款活動："&Act_ShortName
    Response.Write "    </td>"
    Response.Write "    <td width='30%' align='right' style='font-size: 10pt;font-family: 標楷體'>列印日期："&PrintDate&"</td>"
    Response.Write "  </tr>"
  End If
  Response.Write "</table>"
  
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id='datagrid' border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='98%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For I = 1 To FieldsCount
      If I=13 Then
        Response.Write "<td x:str bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>"&RS1(I)&"</span></td>"
      ElseIf I=FieldsCount-1 Or I=FieldsCount Then
        If RS1(I)<>"" Then
          Response.Write "<td bgcolor='#FFFFFF' align='right'><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(I),0) & "</span></td>"
        Else
          Response.Write "<td bgcolor='#FFFFFF' align='right'><span style='font-size: 9pt; font-family: 新細明體'>0</span></td>"
        End If
      Else
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(I))&"</span></td>"
	    End If
	  Next
	  Response.Flush
    Response.Clear
    RS1.MoveNext
	  Response.Write "</tr>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
  Session.Contents.Remove("SQL")
  Session.Contents.Remove("Donate_Date_Begin")
  Session.Contents.Remove("Donate_Date_End")
  Session.Contents.Remove("Act_Id")
End Function 
%>
<%Prog_Id="name_list"%>
<!--#include file="../include/head_rpt.asp"-->
<body class=tool>
<%
  If request("action")="export" Then
  	'20130729 Modify by GoodTV Tanya:修改文宣品項目
    SQL="Select Distinct Donor.Donor_Id,捐款人編號=Donor.Donor_Id,捐款人=Donor_Name,性別=Sex,稱謂=Title,身份類別=Donor_Type,身分證統編=IDNo,出生日期=Birthday,教育程度=Education,職業別=Occupation,婚姻狀況=Marriage,宗教信仰=Religion,道場教會名稱=ReligionName,手機=Cellular_Phone,電話日=Tel_Office,電話夜=Tel_Home,電子信箱=Email,聯絡人=Contactor,服務單位=OrgName,職稱=JobTitle,通訊地址=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+DONOR.ZipCode+C.mValue+Address Else B.mValue+DONOR.ZipCode+Address End End),海外地址=IsAbroad,紙本月刊=IsSendNews,DVD=IsDVD,電子文宣=IsSendEpaper,收據開立=Invoice_Type,收據抬頭=Invoice_Title,收據身分證統編=Invoice_IDNo,收據地址=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+DONOR.Invoice_ZipCode+E.mValue+Invoice_Address Else D.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End),徵信錄匿名=IsAnonymous,匿名=NickName,首次捐款日=Begin_DonateDate,最近捐款日=Last_DonateDate,累計次數=A.Donate_No,累計金額=CONVERT(numeric,A.Donate_Total) "
    Ary_SQL=Split(Session("SQL"),"From")
    For I = 1 To UBound(Ary_SQL)
      SQL=SQL&" From "&Ary_SQL(I)
    Next
  Else
    SQL=Session("SQL")
  End If
  ReportName="捐款人名冊"
  call DonorList (SQL,ReportName)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->