<!--#include file="../include/dbfunctionJ.asp"-->
<%
'邊界:左:8 右:8 上:8 下:8

SQL1="Select Donor_Total=Count(*) From DONOR Where IsMember='Y'"
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
Donor_Total=Cdbl(RS1("Donor_Total"))
RS1.Close
Set RS1=Nothing

CodeDesc=""
'20131003 Modify by GoodTV Tanya
'If Session("Stat_Kind")="Category" Then
'  StatKind="類別"
If Session("Stat_Kind")="Sex" Then
  StatKind="性別"
ElseIf Session("Stat_Kind")="Age" Then
  StatKind="年齡"
  CodeDesc="20,25,30,35,40,45,50,55,60,65,70,71"
ElseIf Session("Stat_Kind")="Education" Then
  StatKind="教育程度"
ElseIf Session("Stat_Kind")="Occupation" Then
  StatKind="職業別"
ElseIf Session("Stat_Kind")="Marriage" Then
  StatKind="婚姻狀況"
ElseIf Session("Stat_Kind")="Religion" Then
  StatKind="宗教信仰"
'ElseIf Session("Stat_Kind")="City" Or Session("Stat_Kind")="Invoice_City" Then
ElseIf Session("Stat_Kind")="City" Then
'  If Session("Stat_Kind")="City" Then StatKind="通訊縣市"
'  If Session("Stat_Kind")="Invoice_City" Then StatKind="收據縣市"
	StatKind="通訊縣市"
  City=""
  SQL1="Select mCode,mValue From CODECITY Where { fn LENGTH(mCode) }=1 Order By Seq"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  While Not RS1.EOF
    If CodeDesc="" Then
      CodeDesc=RS1("mCode")
      City=RS1("mValue")
    Else
      CodeDesc=CodeDesc&","&RS1("mCode")
      City=City&","&RS1("mValue")
    End If
    RS1.MoveNext
  Wend
  RS1.Close
  Set RS1=Nothing
  Ary_City=Split(City,",")
ElseIf Session("Stat_Kind")="Member_Status" Then
  StatKind="狀態"
'ElseIf Session("Stat_Kind")="Member_Type" Then
'  StatKind="會員別"
End If
If CodeDesc="" Then
  If Session("Stat_Kind")="Member_Status" Then
    SQL1="Select CodeDesc From CASECODE Where CodeType='MemberStatus' Order By Seq"
'  ElseIf Session("Stat_Kind")="Member_Type" Then
'    SQL1="Select CodeDesc From CASECODE Where CodeType='MemberType' Order By Seq"
  Else
    SQL1="Select CodeDesc From CASECODE Where CodeType='"&Session("Stat_Kind")&"' Order By Seq"
  End If
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  While Not RS1.EOF
    If CodeDesc="" Then
      CodeDesc=RS1("CodeDesc")
    Else
      CodeDesc=CodeDesc&","&RS1("CodeDesc")
    End If
    RS1.MoveNext
  Wend
  RS1.Close
  Set RS1=Nothing
End If

If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=report.xls"
End If       
%>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>讀者<%=StatKind%>統計</title>
  <%If request("action")<>"export" Then%><link REL="stylesheet" type="text/css" HREF="../include/dms.css"><%End If%>
</head>
<body class=tool>
<p><div align="center"><center>
  <table border="0" width="720" cellspacing="0" cellpadding="4">
    <tr>
			<td align="center" colspan='2'><span style="font-size: 14pt;font-family: 標楷體">讀者<%=StatKind%>統計</span></td>
		</tr>
		<tr>
			<td width="65" align="right"><span style='font-size: 9pt; font-family: 新細明體'>統計項目：</span></td>
			<td width="655"><span style='font-size: 9pt; font-family: 新細明體'><%=StatKind%></span></td>
		</tr>		
  </table>
  <%
    Response.Write "<table width='720' border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0'>"&vbcrlf
    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td width='32%' align='center' bgcolor='#D7E9F1' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKind&"統計</span></td>"&vbcrlf
    Response.Write "    <td width='34%' align='center' bgcolor='#D7E9F1' height='30'><span style='font-size: 9pt; font-family: 新細明體'>人數</span></td>"&vbcrlf
    Response.Write "    <td width='34%' align='center' bgcolor='#D7E9F1' height='30'><span style='font-size: 9pt; font-family: 新細明體'>百分比</span></td>"&vbcrlf
    Response.Write "  </tr>"&vbcrlf
    Ary_CodeDesc=Split(CodeDesc,",")
    For I=0 To UBound(Ary_CodeDesc)
      If Session("Stat_Kind")="Age" Then
        If I=0 Then
          CodeDesc=Cstr(Ary_CodeDesc(I))&"歲以下"
          SQL1="Select Donor_Total_P=Count(*) From DONOR Where IsMember='Y' And Year(Birthday)>='"&Year(Date())-Cdbl(Ary_CodeDesc(I))&"'"
        ElseIf I=UBound(Ary_CodeDesc) Then
          CodeDesc=Cstr(Ary_CodeDesc(I))&"歲以上"
          SQL1="Select Donor_Total_P=Count(*) From DONOR Where IsMember='Y' And Year(Birthday)<='"&Year(Date())-Cdbl(Ary_CodeDesc(I))&"'"
        Else
          CodeDesc=Cstr(Cdbl(Ary_CodeDesc(I-1))+1)&" ～ "&Cstr(Ary_CodeDesc(I))&"歲"
          SQL1="Select Donor_Total_P=Count(*) From DONOR Where IsMember='Y' And Year(Birthday)>='"&Year(Date())-Cdbl(Ary_CodeDesc(I))&"' And Year(Birthday)<='"&Year(Date())-Cdbl(Ary_CodeDesc(I-1))-1&"'"
        End If
      '20131003 Modify by GoodTV Tanya
      'ElseIf Session("Stat_Kind")="City" Or Session("Stat_Kind")="Invoice_City" Then
      ElseIf Session("Stat_Kind")="City" Then
        CodeDesc=Ary_City(I)
        SQL1="Select Donor_Total_P=Count(*) From DONOR Where IsMember='Y' And "&Session("Stat_Kind")&"='"&Ary_CodeDesc(I)&"'"
      Else
        CodeDesc=Ary_CodeDesc(I)
        SQL1="Select Donor_Total_P=Count(*) From DONOR Where IsMember='Y' And "&Session("Stat_Kind")&"='"&Ary_CodeDesc(I)&"'"
      End If
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_No=Cdbl(RS1("Donor_Total_P"))
      Total_NoP=Cdbl((Total_No/Donor_Total))
      RS1.Close
      Set RS1=Nothing
      Response.Write "<tr>"&vbcrlf
      Response.Write "  <td align='center' width='32%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&CodeDesc&"</span></td>"&vbcrlf 
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_No,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf   
      Response.Write "</tr>"&vbcrlf
      Response.Flush
      Response.Clear    
    Next
    
    '資料不詳
    If Session("Stat_Kind")="Age" Then
      SQL1="Select Donor_Total_P=Count(*) From DONOR Where IsMember='Y' And (Birthday Is null)"
    Else
      SQL1="Select Donor_Total_P=Count(*) From DONOR Where IsMember='Y' And ("&Session("Stat_Kind")&"='' Or "&Session("Stat_Kind")&" Is null)"
    End If

    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    If Cdbl(RS1("Donor_Total_P"))>0 Then
      Total_No=RS1("Donor_Total_P")
      Total_NoP=Cdbl((Total_No/Donor_Total))
      Response.Write "<tr>"&vbcrlf
      Response.Write "  <td align='center' width='32%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKind&"不詳</span></td>"&vbcrlf 
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_No,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf   
      Response.Write "</tr>"&vbcrlf
    End If
    RS1.Close
    Set RS1=Nothing

    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td align='center' width='32%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>合計</span></td>"&vbcrlf
    Response.Write "    <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Donor_Total,0)&"</span></td>"&vbcrlf
    Response.Write "    <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(1,0)&"</span></td>"&vbcrlf
    Response.Write "  </tr>"&vbcrlf
    Response.Write "</table>"&vbcrlf
    Session.Contents.Remove("Stat_Kind")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->