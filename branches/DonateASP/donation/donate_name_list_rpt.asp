<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=donate_name.xls"
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
    Response.Write "    <td width='100%' align='center' style='font-size: 13pt;font-family: 標楷體' colspan='13'>"&ReportName&"</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td align='left' style='font-size: 10pt;font-family: 標楷體' colspan='13'>"
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
      If I=3 Then
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
<%Prog_Id="donate_name_list"%>
<!--#include file="../include/head_rpt.asp"-->
<body class=tool>
<%
  SQL=Session("SQL")
  call DonorList (SQL,Prog_Desc)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->