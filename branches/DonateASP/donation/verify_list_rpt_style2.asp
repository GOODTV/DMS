<!--#include file="../include/dbfunctionJ.asp"-->
<%
If Session("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=verify.xls"
End If

Function DonateList (SQL,ReportName)
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
  If Session("action")="export" Then
    Response.Write "  <tr>"
    Response.Write "    <td width='100%' align='center' style='font-size: 13pt;font-family: 標楷體' colspan='3'>"&ReportName&"</td>"
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
  For I = 1 To 2
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  Row=1
  Donate_Total=0
  
  If Not RS1.EOF Then
    Donate_Pre=-1
    Donate_Name=""
    Donate_Title_Old=""
    Donate_Title_No=1
    While Not RS1.EOF
	    If Row=1 Then
	      Donate_Pre=RS1(1)
	      If (Instr(RTrim(LTrim(RS1(2))),"、")>0 Or Instr(LTrim(RTrim(LTrim(RS1(2))))," ")>0) Then
	        Donate_Name="("&RTrim(LTrim(RS1(2)))&")"
	      Else
	        Donate_Name=RTrim(LTrim(RS1(2)))
	      End If
	      If Cstr(RS1(4))="1" Then Donate_Name=Donate_Name&"("&RS1(3)&")"
	    Else
	      If Cstr(RS1(1))<>Cstr(Donate_Pre) Then
	        Response.Write "<tr>"
	        Response.Write "  <td bgcolor='#FFFFFF' align='right'><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(Donate_Pre,0) & "</span></td>"
	        Response.Write "  <td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & Data_Minus(Donate_Name) & "</span></td>"
	        Response.Write "</tr>"
	        Donate_Pre=RS1(1)
	        If (Instr(RTrim(LTrim(RS1(2))),"、")>0 Or Instr(RTrim(LTrim(RS1(2)))," ")>0) Then
	          Donate_Name="("&RTrim(LTrim(RS1(2)))&")"
	        Else
	          Donate_Name=RTrim(LTrim(RS1(2)))
	        End If
	        If Cstr(RS1(4))="1" Then Donate_Name=Donate_Name&"("&RS1(3)&")"
	        Donate_Title_No=1
	      Else
	        If RTrim(LTrim(RS1(5)))<>Donate_Title_Old Then
	          If Donate_Title_No>1 Then Donate_Name=Donate_Name&"X"&Donate_Title_No
	          Donate_Title_No=1
	          If (Instr(RTrim(LTrim(RS1(2))),"、")>0 Or Instr(RTrim(LTrim(RS1(2)))," ")>0) Then
	            Donate_Name=Donate_Name&"、("&RTrim(LTrim(RS1(2)))&")"
	          Else
	            Donate_Name=Donate_Name&"、"&RTrim(LTrim(RS1(2)))
	          End If
	          If Cstr(RS1(4))="1" Then Donate_Name=Donate_Name&"("&RS1(3)&")"
	        Else
	          Donate_Title_No=Donate_Title_No+1
	        End If
	      End If
	    End If
	    Donate_Title_Old=RTrim(LTrim(RS1(5)))
      If RS1(1)<>"" Then Donate_Total=Cdbl(Donate_Total)+Cdbl(RS1(1))
      Row=Row+1
	    Response.Flush
      Response.Clear
      RS1.MoveNext
    Wend
    If Donate_Name<>"" Then
      If Donate_Title_No>0 Then Donate_Name=Donate_Name&"X"&Donate_Title_No
	    Response.Write "<tr>"
	    Response.Write "  <td bgcolor='#FFFFFF' align='right'><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(Donate_Pre,0) & "</span></td>"
	    Response.Write "  <td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & Data_Minus(Donate_Name) & "</span></td>"
	    Response.Write "</tr>"
    End If
  End If
  Response.Write "</table>"
  If Session("Donate_Payment")="1" Then Response.Write "<p align='left' style='font-size: 10pt;font-family: 標楷體'>&nbsp;捐款總金額："&FormatNumber(Donate_Total,0)&"</p>"

  RS1.Close
  Set RS1=Nothing
  Session.Contents.Remove("SQL")
  Session.Contents.Remove("action")
  Session.Contents.Remove("Donate_Date_Begin")
  Session.Contents.Remove("Donate_Date_End")
  Session.Contents.Remove("Act_Id")
  Session.Contents.Remove("Donate_Payment")
End Function 
%>
<%Prog_Id="verify_list"%>
<!--#include file="../include/head_rpt.asp"-->
<body class=tool>
<%
  SQL="Select Donate_Id,捐款金額=Isnull(Donate_Amt,0),捐款人=(Case When IsAnonymous='Y' Then Case When NickName<>'' Then NickName Else Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End End Else Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End End),捐款日期=CONVERT(nvarchar(2),Month(Donate_Date))+'/'+CONVERT(nvarchar(2),Day(Donate_Date)),Donor.Donor_Id,編號抬頭=CONVERT(nvarchar(10),Donate.Donor_Id)+Donate.Invoice_Title "
  Ary_SQL=Split(Session("SQL1"),"From")
  For I = 1 To UBound(Ary_SQL)
    SQL=SQL&" From "&Ary_SQL(I)
  Next
  ReportName="徵信芳名錄"
  call DonateList (SQL,ReportName)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->