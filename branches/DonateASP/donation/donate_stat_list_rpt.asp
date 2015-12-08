<!--#include file="../include/dbfunctionJ.asp"-->
<%
'邊界:左:8 右:8 上:8 下:8
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=donate_sata.xls"
End If

Comp_ShortName=""
If Session("DeptId")<>"" Then
  SQL1="Select Comp_ShortName From DEPT Where Dept_Id='"&Session("DeptId")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Comp_ShortName=RS1("Comp_ShortName")
  RS1.Close
  Set RS1=Nothing
End If

Act_ShortName=""
If Session("Act_Id")<>"" Then
  SQL1="Select Act_ShortName From ACT Where Act_Id='"&Session("Act_Id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Act_ShortName=RS1("Act_ShortName")
  RS1.Close
  Set RS1=Nothing
End If

Donate_Where=""
If Session("DeptId")<>"" Then
  Donate_Where=Donate_Where&"Where 1=1 And Donate.Dept_Id = '"&Session("DeptId")&"' "
Else
  Donate_Where=Donate_Where&"Where 1=1 And Donate.Dept_Id In ("&Session("all_dept_type")&") "
End If
If Session("Donate_Date_Begin")<>"" Then Donate_Where=Donate_Where&"And Donate_Date>='"&Session("Donate_Date_Begin")&"' "
If Session("Donate_Date_End")<>"" Then Donate_Where=Donate_Where&"And Donate_Date<='"&Session("Donate_Date_End")&"' "
If Session("Act_Id")<>"" Then Donate_Where=Donate_Where&"And Act_Id='"&Session("Act_Id")&"' "
If Session("Donate_Purpose_Type")<>"" Then
  Donate_Purpose_Type=""
  Ary_Purpose_Type=Split(Session("Donate_Purpose_Type"),",")
  For I = 0 To UBound(Ary_Purpose_Type)
    If Donate_Purpose_Type="" Then
      Donate_Purpose_Type="'"&Trim(Ary_Purpose_Type(I))&"'"
    Else
      Donate_Purpose_Type=Donate_Purpose_Type&",'"&Trim(Ary_Purpose_Type(I))&"'"
    End If
  Next
  If Donate_Purpose_Type<>"" Then Donate_Where=Donate_Where&"And Donate_Purpose_Type In ("&Donate_Purpose_Type&") "
End If

Total_Amt=0
SQL1="Select Total_Amt=Isnull(Sum(Donate_Amt),0) From Donate "&Donate_Where&" "
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
Total_Amt=Csng(RS1("Total_Amt"))
RS1.Close
Set RS1=Nothing

Total_No=0
SQL1="Select Total_No=Count(*) From Donate "&Donate_Where&" "
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
Total_No=Csng(RS1("Total_No"))
RS1.Close
Set RS1=Nothing
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <%If request("action")<>"export" Then%><link REL="stylesheet" type="text/css" HREF="../include/dms.css"><%End If%>
  <title><%=ProgDesc("stat")%></title>
</head>
<body class=tool>
  <p><div align="center"><center>
  <table border="0" width="720" cellspacing="0" cellpadding="4">
    <tr>
			<td align="center" colspan='2'><span style="font-size: 14pt;font-family: 標楷體">捐款統計分析</span></td>
		</tr>
		<tr>
			<td width="65" align="right"><span style='font-size: 9pt; font-family: 新細明體'>機構：</span></td>
			<td width="655"><span style='font-size: 9pt; font-family: 新細明體'><%=Comp_ShortName%></span></td>
		</tr>
		<tr>
			<td align="right"><span style='font-size: 9pt; font-family: 新細明體'>捐款期間：</span></td>
			<td><span style='font-size: 9pt; font-family: 新細明體'><%=Session("Donate_Date_Begin")%> ～ <%=Session("Donate_Date_End")%></span></td>
		</tr>
		<%If Act_ShortName<>"" Then%>
		<tr>
			<td align="right"><span style='font-size: 9pt; font-family: 新細明體'>勸募活動：</span></td>
			<td><span style='font-size: 9pt; font-family: 新細明體'><%=Act_ShortName%></span></td>
		</tr>
		<%End If%>
  </table>
  <%
    CodeDesc=""
    If Session("Stat_Kind")="Category" Then
      StatKind="類別"
    ElseIf Session("Stat_Kind")="Sex" Then
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
    ElseIf Session("Stat_Kind")="City" Or Session("Stat_Kind")="Invoice_City" Then
      If Session("Stat_Kind")="City" Then StatKind="通訊縣市"
      If Session("Stat_Kind")="Invoice_City" Then StatKind="收據縣市"
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
    ElseIf Session("Stat_Kind")="Donate_Payment" Then
      StatKind="捐款方式"
    ElseIf Session("Stat_Kind")="Donate_Purpose" Then
      StatKind="捐款用途"
    ElseIf Session("Stat_Kind")="Donate_Amt" Then
      StatKind="捐款金額"
      CodeDesc=Session("Donate_Amt_Kind")
      If Instr(CodeDesc,",")>0 Then CodeDesc=CodeDesc&","&Cstr(Csng(Split(CodeDesc,",")(UBound(Split(CodeDesc,","))))+1)
    End If
    If CodeDesc="" Then
      SQL1="Select CodeDesc From CASECODE Where CodeType='"&replace(Session("Stat_Kind"),"Donate_","")&"' Order By Seq"
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

    Response.Write "<table width='720' border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0'>"&vbcrlf
    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td width='32%' align='center' bgcolor='#D7E9F1' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKind&"</span></td>"&vbcrlf
    Response.Write "    <td width='34%' align='center' bgcolor='#D7E9F1' height='30' colspan='2'><span style='font-size: 9pt; font-family: 新細明體'>捐款金額</span></td>"&vbcrlf
    Response.Write "    <td width='34%' align='center' bgcolor='#D7E9F1' height='30' colspan='2'><span style='font-size: 9pt; font-family: 新細明體'>捐款次數</span></td>"&vbcrlf
    Response.Write "  </tr>"&vbcrlf
    Ary_CodeDesc=Split(CodeDesc,",")
    For I=0 To UBound(Ary_CodeDesc)
      If Session("Stat_Kind")="Age" Then
        If I=0 Then
          CodeDesc=Cstr(Ary_CodeDesc(I))&"歲以下"
          DonateWhere=replace(Donate_Where,"Where 1=1","Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDesc(I))&"' ")
        ElseIf I=UBound(Ary_CodeDesc) Then
          CodeDesc=Cstr(Ary_CodeDesc(I))&"歲以上"
          DonateWhere=replace(Donate_Where,"Where 1=1","Where Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDesc(I))&"' ")
        Else
          CodeDesc=Cstr(Csng(Ary_CodeDesc(I-1))+1)&" ～ "&Cstr(Ary_CodeDesc(I))&"歲"
          DonateWhere=replace(Donate_Where,"Where 1=1","Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDesc(I))&"' And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDesc(I-1))-1&"' ")
        End If
      ElseIf Session("Stat_Kind")="City" Or Session("Stat_Kind")="Invoice_City" Then
        CodeDesc=Ary_City(I)
        DonateWhere=replace(Donate_Where,"Where 1=1","Where "&Session("Stat_Kind")&"='"&Ary_CodeDesc(I)&"'")
      ElseIf Session("Stat_Kind")="Donate_Amt" Then
        If I=0 Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDesc(I),0))&"元以下"
          DonateWhere=replace(Donate_Where,"Where 1=1","Where CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDesc(I)&"' ")
        ElseIf I=UBound(Ary_CodeDesc) Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDesc(I),0))&"元以上"
          DonateWhere=replace(Donate_Where,"Where 1=1","Where CONVERT(numeric,Donate_Amt)>='"&Ary_CodeDesc(I)&"' ")
        Else
          CodeDesc=Cstr(FormatNumber(Csng(Ary_CodeDesc(I-1))+1,0))&" ～ "&Cstr(FormatNumber(Ary_CodeDesc(I),0))&"元"
          DonateWhere=replace(Donate_Where,"Where 1=1","Where CONVERT(numeric,Donate_Amt)>='"&Csng(Ary_CodeDesc(I-1))+1&"' And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDesc(I)&"' ")
        End If
      Else
        CodeDesc=Ary_CodeDesc(I)
        DonateWhere=replace(Donate_Where,"Where 1=1","Where "&Session("Stat_Kind")&"='"&Ary_CodeDesc(I)&"'")
      End If
       
      Total_AmtK=0
      Total_AmtP=0
      SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_AmtK=Csng(RS1("Total_AmtK"))
      RS1.Close
      Set RS1=Nothing
      If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))

      Total_NoK=0
      Total_NoP=0
      SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "

      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_NoK=Csng(RS1("Total_NoK"))
      RS1.Close
      Set RS1=Nothing
      If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
      
      Response.Write "<tr>"&vbcrlf
      Response.Write "  <td align='center' width='32%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&CodeDesc&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf   
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf   
      Response.Write "</tr>"&vbcrlf
      Response.Flush
      Response.Clear    
    Next
    
    '資料不詳
    If Session("Stat_Kind")="Age" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Birthday Is Null ")
    ElseIf Session("Stat_Kind")="Donate_Amt" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Donate_Amt Is Null ")
    Else
      DonateWhere=replace(Donate_Where,"Where 1=1","Where ("&Session("Stat_Kind")&"='' Or "&Session("Stat_Kind")&" Is Null)")
    End If
    
    Total_AmtK=0
    Total_AmtP=0
    SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    Total_AmtK=Csng(RS1("Total_AmtK"))
    RS1.Close
    Set RS1=Nothing
    If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))

    Total_NoK=0
    Total_NoP=0
    SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    Total_NoK=Csng(RS1("Total_NoK"))
    RS1.Close
    Set RS1=Nothing
    If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
          
    If Total_AmtK>0 Then
      Response.Write "<tr>"&vbcrlf
      Response.Write "  <td align='center' width='32%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKind&"不詳</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf
      Response.Write "</tr>"&vbcrlf 
	  End If

    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td align='center' width='32%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>合計</span></td>"&vbcrlf
    Response.Write "    <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_Amt,0)&"</span></td>"&vbcrlf
    Response.Write "    <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(1,0)&"</span></td>"&vbcrlf
    Response.Write "    <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_No,0)&"</span></td>"&vbcrlf
    Response.Write "    <td align='center' width='17%' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(1,0)&"</span></td>"&vbcrlf
    Response.Write "  </tr>"&vbcrlf	  
	  Response.Write "</table>"&vbcrlf
	  
    Session.Contents.Remove("DeptId")
    Session.Contents.Remove("Donate_Date_Begin")
    Session.Contents.Remove("Donate_Date_End")
    Session.Contents.Remove("Act_Id")
    Session.Contents.Remove("Stat_Kind")
    Session.Contents.Remove("Donate_Amt_Kind")
    Session.Contents.Remove("Donate_Purpose_Type")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->