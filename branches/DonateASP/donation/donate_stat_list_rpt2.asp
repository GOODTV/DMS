<!--#include file="../include/dbfunctionJ.asp"-->
<%
'邊界:左:8 右:8 上:8 下:8
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=donate_stat2.xls"
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

CodeDescX=""
If Session("Stat_KindX")="Category" Then
  StatKindX="類別"
ElseIf Session("Stat_KindX")="Sex" Then
  StatKindX="性別"
ElseIf Session("Stat_KindX")="Age" Then
  StatKindX="年齡"
  CodeDescX="20,25,30,35,40,45,50,55,60,65,70,71"
ElseIf Session("Stat_KindX")="Education" Then
  StatKindX="教育程度"
ElseIf Session("Stat_KindX")="Occupation" Then
  StatKindX="職業別"
ElseIf Session("Stat_KindX")="Marriage" Then
  StatKindX="婚姻狀況"
ElseIf Session("Stat_KindX")="Religion" Then
  StatKindX="宗教信仰"
ElseIf Session("Stat_KindX")="City" Or Session("Stat_KindX")="Invoice_City" Then
  If Session("Stat_KindX")="City" Then StatKindX="通訊縣市"
  If Session("Stat_KindX")="Invoice_City" Then StatKindX="收據縣市"
  City=""
  SQL1="Select mCode,mValue From CODECITY Where { fn LENGTH(mCode) }=1 Order By Seq"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  While Not RS1.EOF
    If CodeDescX="" Then
      CodeDescX=RS1("mCode")
      City=RS1("mValue")
    Else
      CodeDescX=CodeDescX&","&RS1("mCode")
      City=City&","&RS1("mValue")
    End If
    RS1.MoveNext
  Wend
  RS1.Close
  Set RS1=Nothing
  Ary_CityX=Split(City,",")
ElseIf Session("Stat_KindX")="Donate_Payment" Then
  StatKindX="捐款方式"
ElseIf Session("Stat_KindX")="Donate_Purpose" Then
  StatKindX="捐款用途"
ElseIf Session("Stat_KindX")="Donate_Amt" Then
  StatKindX="捐款金額"
  CodeDescX=Session("Donate_Amt_Kind")
  If Instr(CodeDescX,",")>0 Then CodeDescX=CodeDescX&","&Cstr(Csng(Split(CodeDescX,",")(UBound(Split(CodeDescX,","))))+1)
End If
If CodeDescX="" Then
  SQL1="Select CodeDesc From CASECODE Where CodeType='"&replace(Session("Stat_KindX"),"Donate_","")&"' Order By Seq"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  While Not RS1.EOF
    If CodeDescX="" Then
      CodeDescX=RS1("CodeDesc")
    Else
      CodeDescX=CodeDescX&","&RS1("CodeDesc")
    End If
    RS1.MoveNext
  Wend
  RS1.Close
  Set RS1=Nothing
End If

CodeDescY=""
If Session("Stat_KindY")="Category" Then
  StatKindY="類別"
ElseIf Session("Stat_KindY")="Sex" Then
  StatKindY="性別"
ElseIf Session("Stat_KindY")="Age" Then
  StatKindY="年齡"
  CodeDescY="20,25,30,35,40,45,50,55,60,65,70,71"
ElseIf Session("Stat_KindY")="Education" Then
  StatKindY="教育程度"
ElseIf Session("Stat_KindY")="Occupation" Then
  StatKindY="職業別"
ElseIf Session("Stat_KindY")="Marriage" Then
  StatKindY="婚姻狀況"
ElseIf Session("Stat_KindY")="Religion" Then
  StatKindY="宗教信仰"
ElseIf Session("Stat_KindY")="City" Or Session("Stat_KindY")="Invoice_City" Then
  If Session("Stat_KindY")="City" Then StatKindY="通訊縣市"
  If Session("Stat_KindY")="Invoice_City" Then StatKindY="收據縣市"
  City=""
  SQL1="Select mCode,mValue From CODECITY Where { fn LENGTH(mCode) }=1 Order By Seq"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  While Not RS1.EOF
    If CodeDescY="" Then
      CodeDescY=RS1("mCode")
      City=RS1("mValue")
    Else
      CodeDescY=CodeDescY&","&RS1("mCode")
      City=City&","&RS1("mValue")
    End If
    RS1.MoveNext
  Wend
  RS1.Close
  Set RS1=Nothing
  Ary_CityY=Split(City,",")
ElseIf Session("Stat_KindY")="Donate_Payment" Then
  StatKindY="捐款方式"
ElseIf Session("Stat_KindY")="Donate_Purpose" Then
  StatKindY="捐款用途"
ElseIf Session("Stat_KindY")="Donate_Amt" Then
  StatKindY="捐款金額"
  CodeDescY=Session("Donate_Amt_Kind")
  If Instr(CodeDescY,",")>0 Then CodeDescY=CodeDescY&","&Cstr(Csng(Split(CodeDescY,",")(UBound(Split(CodeDescY,","))))+1)
End If
If CodeDescY="" Then
  SQL1="Select CodeDesc From CASECODE Where CodeType='"&replace(Session("Stat_KindY"),"Donate_","")&"' Order By Seq"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  While Not RS1.EOF
    If CodeDescY="" Then
      CodeDescY=RS1("CodeDesc")
    Else
      CodeDescY=CodeDescY&","&RS1("CodeDesc")
    End If
    RS1.MoveNext
  Wend
  RS1.Close
  Set RS1=Nothing
End If
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <%If request("action")<>"export" Then%><link REL="stylesheet" type="text/css" HREF="../include/dms.css"><%End If%>
  <title><%=ProgDesc("stat2")%></title>
</head>
<body class=tool>
  <p><div align="center"><center>
  <table border="0" width="980" cellspacing="0" cellpadding="4">
    <tr>
			<td align="center" colspan='2'><span style="font-size: 14pt;font-family: 標楷體">捐款交叉分析</span></td>
		</tr>
		<tr>
			<td width="70" align="right"><span style='font-size: 9pt; font-family: 新細明體'>機構：</span></td>
			<td width="910"><span style='font-size: 9pt; font-family: 新細明體'><%=Comp_ShortName%></span></td>
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
		<tr>
			<td align="right"><span style='font-size: 9pt; font-family: 新細明體'>統計項目X：</span></td>
			<td><span style='font-size: 9pt; font-family: 新細明體'></span><%=StatKindX%></td>
		</tr>
		<tr>
			<td align="right"><span style='font-size: 9pt; font-family: 新細明體'>統計項目Y：</span></td>
			<td><span style='font-size: 9pt; font-family: 新細明體'></span><%=StatKindY%></td>
		</tr>				
  </table>
  <%
    Ary_CodeDescX=Split(CodeDescX,",")
    Ary_CodeDescY=Split(CodeDescY,",")
    Width_Default=Cint(980/(UBound(Ary_CodeDescX)+4))
    
    '捐款金額
    Response.Write "<table width='980' border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0'>"&vbcrlf
    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td align='center' wudth='100' width='"&Width_Default&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'><b>捐款金額</b></span></td>"&vbcrlf
    For I=0 To UBound(Ary_CodeDescX)
      If Session("Stat_KindX")="Age" Then
        If I=0 Then
          CodeDesc=Cstr(Ary_CodeDescX(I))&"歲以下"
        ElseIf I=UBound(Ary_CodeDescX) Then
          CodeDesc=Cstr(Ary_CodeDescX(I))&"歲以上"
        Else
          CodeDesc=Cstr(Csng(Ary_CodeDescX(I-1))+1)&" ～ "&Cstr(Ary_CodeDescX(I))&"歲"
        End If
      ElseIf Session("Stat_KindX")="City" Or Session("Stat_KindX")="Invoice_City" Then
        CodeDesc=Ary_CityX(I)
      ElseIf Session("Stat_KindX")="Donate_Amt" Then
        If I=0 Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDescX(I),0))&"元以下"
        ElseIf I=UBound(Ary_CodeDescX) Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDescX(I),0))&"元以上"
        Else
          CodeDesc=Cstr(FormatNumber(Csng(Ary_CodeDescX(I-1))+1,0))&" ～ "&Cstr(FormatNumber(Ary_CodeDescX(I),0))&"元"
        End If
      Else
        CodeDesc=Ary_CodeDescX(I)
      End If
      Response.Write "  <td align='center' bgcolor='#D7E9F1' width='"&Width_Default&"' height='30' colspan='2'><span style='font-size: 9pt; font-family: 新細明體'>"&CodeDesc&"</span></td>"&vbcrlf
    Next

    Check_X=False
    If Session("Stat_KindX")="Age" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Birthday Is Null ")
    ElseIf Session("Stat_KindX")="Donate_Amt" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Donate_Amt Is Null ")
    Else
      DonateWhere=replace(Donate_Where,"Where 1=1 ","Where ("&Session("Stat_KindX")&"='' Or "&Session("Stat_KindX")&" Is Null) ")
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
    If Total_AmtK>0 Then
      Check_X=True
      Response.Write "<td align='center' bgcolor='#D7E9F1' width='"&Width_Default&"' height='30' colspan='2'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKindX&"不詳</span></td>"&vbcrlf
    End If
    Response.Write "  <td align='center' bgcolor='#FFFF99' width='"&Width_Default&"' height='30' colspan='2'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKindY&"(Y)小計</span></td>"&vbcrlf
    Response.Write "</tr>"&vbcrlf

    For I=0 To UBound(Ary_CodeDescY)
      Response.Write "<tr>"&vbcrlf
      DonateWhere_Y=""
      If Session("Stat_KindY")="Age" Then
        If I=0 Then
          CodeDesc=Cstr(Ary_CodeDescY(I))&"歲以下"
          DonateWhere_Y="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescY(I))&"' "
        ElseIf I=UBound(Ary_CodeDescY) Then
          CodeDesc=Cstr(Ary_CodeDescY(I))&"歲以上"
          DonateWhere_Y="Where Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescY(I))&"' "
        Else
          CodeDesc=Cstr(Csng(Ary_CodeDescY(I-1))+1)&" ～ "&Cstr(Ary_CodeDescY(I))&"歲"
          DonateWhere_Y="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescY(I))&"' And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescY(I-1))-1&"' "
        End If
      ElseIf Session("Stat_KindY")="City" Or Session("Stat_KindY")="Invoice_City" Then
        CodeDesc=Ary_CityY(I)
        DonateWhere_Y="Where "&Session("Stat_KindY")&"='"&Ary_CodeDescY(I)&"' "
      ElseIf Session("Stat_KindY")="Donate_Amt" Then
        If I=0 Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDescY(I),0))&"元以下"
          DonateWhere_Y="Where CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescY(I)&"' "
        ElseIf I=UBound(Ary_CodeDescY) Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDescY(I),0))&"元以上"
          DonateWhere_Y="Where CONVERT(numeric,Donate_Amt)>='"&Ary_CodeDescY(I)&"' "
        Else
          CodeDesc=Cstr(FormatNumber(Csng(Ary_CodeDescY(I-1))+1,0))&" ～ "&Cstr(FormatNumber(Ary_CodeDescY(I),0))&"元"
          DonateWhere_Y="Where CONVERT(numeric,Donate_Amt)>='"&Csng(Ary_CodeDescY(I-1))+1&"' And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescY(I)&"' "
        End If
      Else
        CodeDesc=Ary_CodeDescY(I)
        DonateWhere_Y="Where "&Session("Stat_KindY")&"='"&Ary_CodeDescY(I)&"' "
      End If
      Response.Write "  <td align='center' bgcolor='#D7E9F1' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&CodeDesc&"</span></td>"&vbcrlf	      

      For J=0 To UBound(Ary_CodeDescX)
        If Session("Stat_KindX")="Age" Then
          If J=0 Then
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(J))&"' ") 
          ElseIf J=UBound(Ary_CodeDescX) Then
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(J))&"' ")
          Else
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(J))&"' And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(J-1))-1&"' ")
          End If
        ElseIf Session("Stat_KindX")="Donate_Amt" Then
          If J=0 Then
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(J)&"' ")
          ElseIf J=UBound(Ary_CodeDescX) Then
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And CONVERT(numeric,Donate_Amt)>='"&Ary_CodeDescX(J)&"' ")
          Else
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And CONVERT(numeric,Donate_Amt)>='"&Csng(Ary_CodeDescX(J-1))+1&"' And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(J)&"' ")
          End If
        Else
          DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And "&Session("Stat_KindX")&"='"&Ary_CodeDescX(J)&"'")
        End If
        Total_AmtK=0
        Total_AmtP=0
        SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
        '''Response.Write SQL1&"<br />"
        '''Response.End
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,1
        Total_AmtK=Csng(RS1("Total_AmtK"))
        RS1.Close
        Set RS1=Nothing
        If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf        
        Response.Flush
        Response.Clear
      Next

      If Check_X Then
        If Session("Stat_KindX")="Age" Then
          DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And Birthday Is Null ")
        ElseIf Session("Stat_KindX")="Donate_Amt" Then
          DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&"' And Donate_Amt Is Null ")
        Else
          DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And ("&Session("Stat_KindX")&"='' Or "&Session("Stat_KindX")&" Is Null)")
        End If
        Total_AmtK=0
        Total_AmtP=0
        SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
        '''Response.Write SQL1&"<br />"
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,1
        Total_AmtK=Csng(RS1("Total_AmtK"))
        RS1.Close
        Set RS1=Nothing
        If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))
        Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
        Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf
        Response.Flush
        Response.Clear
      End If

      DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" ")
      Total_AmtK=0
      Total_AmtP=0
      SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
      '''Response.Write SQL1&"<br />"
      '''Response.End
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_AmtK=Csng(RS1("Total_AmtK"))
      RS1.Close
      Set RS1=Nothing
      If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf
      Response.Write "</tr>"&vbcrlf
      Response.Flush
      Response.Clear
    Next
    
    Check_Y=False
    If Session("Stat_KindY")="Age" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Birthday Is Null ")
    ElseIf Session("Stat_KindY")="Donate_Amt" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Donate_Amt Is Null ")
    Else
      DonateWhere=replace(Donate_Where,"Where 1=1","Where ("&Session("Stat_KindY")&"='' Or "&Session("Stat_KindY")&" Is Null)")
    End If
    Total_AmtK=0
    Total_AmtP=0
    SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    Total_AmtK=Csng(RS1("Total_AmtK"))
    RS1.Close
    Set RS1=Nothing
    If Total_AmtK>0 Then
      Check_Y=True
      Response.Write "  <tr>"&vbcrlf
      Response.Write "    <td align='center' bgcolor='#D7E9F1' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKindY&"不詳</span></td>"&vbcrlf
      For I=0 To UBound(Ary_CodeDescX)
        DonateWhere=""
        If Session("Stat_KindX")="Age" Then
          If I=0 Then
            DonateWhere="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' "
          ElseIf I=UBound(Ary_CodeDescX) Then
            DonateWhere="Where Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' "
          Else
            DonateWhere="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(I-1))-1&"' "
          End If
        ElseIf Session("Stat_KindX")="Donate_Amt" Then
          If I=0 Then
            DonateWhere="Where CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(I)&"' "
          ElseIf I=UBound(Ary_CodeDescY) Then
            DonateWhere="Where CONVERT(numeric,Donate_Amt)>='"&Ary_CodeDescX(I)&"' "
          Else
            DonateWhere="Where CONVERT(numeric,Donate_Amt)>='"&Csng(Ary_CodeDescX(I-1))+1&"' And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(I)&"' "
          End If
        Else
          DonateWhere="Where "&Session("Stat_KindX")&"='"&Ary_CodeDescX(I)&"' And ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') "
        End If

        If Session("Stat_KindY")="Age" Then
          DonateWhere=DonateWhere&"And Birthday Is Null "
        ElseIf Session("Stat_KindY")="Donate_Amt" Then
          DonateWhere=DonateWhere&"And Donate_Amt Is Null "
        Else
          DonateWhere=DonateWhere&"And ("&Session("Stat_KindY")&"='' Or "&Session("Stat_KindY")&" Is Null) "
        End If
            
        Total_AmtK=0
        Total_AmtP=0
        DonateWhere=replace(Donate_Where,"Where 1=1",DonateWhere)
        SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
        '''Response.Write "SQL1:"&SQL1&"<br />"
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,1
        Total_AmtK=Csng(RS1("Total_AmtK"))
        RS1.Close
        Set RS1=Nothing
        If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf
        Response.Flush
        Response.Clear
      Next
      
      If Check_X Then
        DonateWhere=""
        If Session("Stat_KindX")="Age" Then
          DonateWhere="Where ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') And Birthday Is Null "
        ElseIf Session("Stat_KindX")="Donate_Amt" Then
          DonateWhere="Where ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') And Donate_Amt Is Null "
        Else
          DonateWhere="Where ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') And ("&Session("Stat_KindX")&" Is Null Or "&Session("Stat_KindX")&"='') "
        End If
        Total_AmtK=0
        Total_AmtP=0
        DonateWhere=replace(Donate_Where,"Where 1=1",DonateWhere)
        SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
        '''Response.Write "SQL1:"&SQL1&"<br />"
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,1
        Total_AmtK=Csng(RS1("Total_AmtK"))
        RS1.Close
        Set RS1=Nothing
        If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf
        Response.Flush
        Response.Clear
      End If

      DonateWhere=""
      If Session("Stat_KindY")="Age" Then
        DonateWhere="Where Birthday Is Null "
      ElseIf Session("Stat_KindX")="Donate_Amt" Then
        DonateWhere="Where Donate_Amt Is Null "
      Else
        DonateWhere="Where ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') "
      End If
      Total_AmtK=0
      Total_AmtP=0
      DonateWhere=replace(Donate_Where,"Where 1=1",DonateWhere)
      SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
      '''Response.Write "SQL1:"&SQL1&"<br />"
      '''Response.End
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_AmtK=Csng(RS1("Total_AmtK"))
      RS1.Close
      Set RS1=Nothing
      If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf
      Response.Flush
      Response.Clear
      
      Response.Write "  </tr>"&vbcrlf
      Response.Flush
      Response.Clear
    End If
    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td align='center' bgcolor='#FFFF99' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKindX&"(X)小計</span></td>"&vbcrlf

    For I=0 To UBound(Ary_CodeDescX)
      DonateWhere=""
      If Session("Stat_KindX")="Age" Then
        If I=0 Then
          DonateWhere="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' "
        ElseIf I=UBound(Ary_CodeDescX) Then
          DonateWhere="Where Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' "
        Else
          DonateWhere="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(I-1))-1&"' "
        End If
      ElseIf Session("Stat_KindX")="Donate_Amt" Then
        If I=0 Then
          DonateWhere="Where CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(I)&"' "
        ElseIf I=UBound(Ary_CodeDescY) Then
          DonateWhere="Where CONVERT(numeric,Donate_Amt)>='"&Ary_CodeDescX(I)&"' "
        Else
          DonateWhere="Where CONVERT(numeric,Donate_Amt)>='"&Csng(Ary_CodeDescX(I-1))+1&"' And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(I)&"' "
        End If
      Else
        DonateWhere="Where "&Session("Stat_KindX")&"='"&Ary_CodeDescX(I)&"' "
      End If
      DonateWhere=replace(Donate_Where,"Where 1=1 ",DonateWhere)
      Total_AmtK=0
      Total_AmtP=0
      SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
      '''Response.Write "SQL1:"&SQL1&"<br />"
      '''Response.End
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_AmtK=Csng(RS1("Total_AmtK"))
      RS1.Close
      Set RS1=Nothing
      If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf
      Response.Flush
      Response.Clear
    Next

    If Check_X Then
      If Session("Stat_KindX")="Age" Then
        DonateWhere=replace(Donate_Where,"Where 1=1","Where Birthday Is Null ")
      ElseIf Session("Stat_KindX")="Donate_Amt" Then
        DonateWhere=replace(Donate_Where,"Where 1=1","Where Donate_Amt Is Null ")
      Else
        DonateWhere=replace(Donate_Where,"Where 1=1","Where ("&Session("Stat_KindX")&"='' Or "&Session("Stat_KindX")&" Is Null)")
      End If
      Total_AmtK=0
      Total_AmtP=0
      SQL1="Select Total_AmtK=Isnull(Sum(Donate_Amt),0) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
      '''Response.Write SQL1&"<br />"
      '''Response.End
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_AmtK=Csng(RS1("Total_AmtK"))
      RS1.Close
      Set RS1=Nothing
      If Total_Amt>0 And Total_AmtK>0 Then Total_AmtP=Csng((Total_AmtK/Total_Amt))
      Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_AmtK,0)&"</span></td>"&vbcrlf
      Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_AmtP,2)&"</span></td>"&vbcrlf
      Response.Flush
      Response.Clear
    End If
    Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'><b>"&FormatNumber(Total_Amt,0)&"</b></span></td>"&vbcrlf
    Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'><b>"&FormatPercent(1,2)&"</b></span></td>"&vbcrlf  
    Response.Write "  </tr>"&vbcrlf
    Response.Write "</table>"&vbcrlf

    Response.Write "<table width='980' border='0' cellspacing='0' cellpadding='0'><tr><td height='30'> </td></tr></table>"&vbcrlf	
      	  
	  '捐款次數
    Response.Write "<table width='980' border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0'>"&vbcrlf
    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td align='center' wudth='100' width='"&Width_Default&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'><b>捐款次數<b/></span></td>"&vbcrlf
    For I=0 To UBound(Ary_CodeDescX)
      If Session("Stat_KindX")="Age" Then
        If I=0 Then
          CodeDesc=Cstr(Ary_CodeDescX(I))&"歲以下"
        ElseIf I=UBound(Ary_CodeDescX) Then
          CodeDesc=Cstr(Ary_CodeDescX(I))&"歲以上"
        Else
          CodeDesc=Cstr(Csng(Ary_CodeDescX(I-1))+1)&" ～ "&Cstr(Ary_CodeDescX(I))&"歲"
        End If
      ElseIf Session("Stat_KindX")="City" Or Session("Stat_KindX")="Invoice_City" Then
        CodeDesc=Ary_CityX(I)
      ElseIf Session("Stat_KindX")="Donate_Amt" Then
        If I=0 Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDescX(I),0))&"元以下"
        ElseIf I=UBound(Ary_CodeDescX) Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDescX(I),0))&"元以上"
        Else
          CodeDesc=Cstr(FormatNumber(Csng(Ary_CodeDescX(I-1))+1,0))&" ～ "&Cstr(FormatNumber(Ary_CodeDescX(I),0))&"元"
        End If
      Else
        CodeDesc=Ary_CodeDescX(I)
      End If
      Response.Write "  <td align='center' bgcolor='#D7E9F1' width='"&Width_Default&"' height='30' colspan='2'><span style='font-size: 9pt; font-family: 新細明體'>"&CodeDesc&"</span></td>"&vbcrlf
    Next

    Check_X=False
    If Session("Stat_KindX")="Age" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Birthday Is Null ")
    ElseIf Session("Stat_KindX")="Donate_Amt" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Donate_Amt Is Null ")
    Else
      DonateWhere=replace(Donate_Where,"Where 1=1 ","Where ("&Session("Stat_KindX")&"='' Or "&Session("Stat_KindX")&" Is Null) ")
    End If
    Total_NoK=0
    Total_NoP=0
    SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    Total_NoK=Csng(RS1("Total_NoK"))
    RS1.Close
    Set RS1=Nothing
    If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
    If Total_NoK>0 Then
      Check_X=True
      Response.Write "<td align='center' bgcolor='#D7E9F1' width='"&Width_Default&"' height='30' colspan='2'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKindX&"不詳</span></td>"&vbcrlf
    End If
    Response.Write "  <td align='center' bgcolor='#FFFF99' width='"&Width_Default&"' height='30' colspan='2'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKindY&"(Y)小計</span></td>"&vbcrlf
    Response.Write "</tr>"&vbcrlf

    For I=0 To UBound(Ary_CodeDescY)
      Response.Write "<tr>"&vbcrlf
      DonateWhere_Y=""
      If Session("Stat_KindY")="Age" Then
        If I=0 Then
          CodeDesc=Cstr(Ary_CodeDescY(I))&"歲以下"
          DonateWhere_Y="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescY(I))&"' "
        ElseIf I=UBound(Ary_CodeDescY) Then
          CodeDesc=Cstr(Ary_CodeDescY(I))&"歲以上"
          DonateWhere_Y="Where Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescY(I))&"' "
        Else
          CodeDesc=Cstr(Csng(Ary_CodeDescY(I-1))+1)&" ～ "&Cstr(Ary_CodeDescY(I))&"歲"
          DonateWhere_Y="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescY(I))&"' And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescY(I-1))-1&"' "
        End If
      ElseIf Session("Stat_KindY")="City" Or Session("Stat_KindY")="Invoice_City" Then
        CodeDesc=Ary_CityY(I)
        DonateWhere_Y="Where "&Session("Stat_KindY")&"='"&Ary_CodeDescY(I)&"' "
      ElseIf Session("Stat_KindY")="Donate_Amt" Then
        If I=0 Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDescY(I),0))&"元以下"
          DonateWhere_Y="Where CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescY(I)&"' "
        ElseIf I=UBound(Ary_CodeDescY) Then
          CodeDesc=Cstr(FormatNumber(Ary_CodeDescY(I),0))&"元以上"
          DonateWhere_Y="Where CONVERT(numeric,Donate_Amt)>='"&Ary_CodeDescY(I)&"' "
        Else
          CodeDesc=Cstr(FormatNumber(Csng(Ary_CodeDescY(I-1))+1,0))&" ～ "&Cstr(FormatNumber(Ary_CodeDescY(I),0))&"元"
          DonateWhere_Y="Where CONVERT(numeric,Donate_Amt)>='"&Csng(Ary_CodeDescY(I-1))+1&"' And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescY(I)&"' "
        End If
      Else
        CodeDesc=Ary_CodeDescY(I)
        DonateWhere_Y="Where "&Session("Stat_KindY")&"='"&Ary_CodeDescY(I)&"' "
      End If
      Response.Write "  <td align='center' bgcolor='#D7E9F1' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&CodeDesc&"</span></td>"&vbcrlf	      

      For J=0 To UBound(Ary_CodeDescX)
        If Session("Stat_KindX")="Age" Then
          If J=0 Then
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(J))&"' ") 
          ElseIf J=UBound(Ary_CodeDescX) Then
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(J))&"' ")
          Else
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(J))&"' And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(J-1))-1&"' ")
          End If
        ElseIf Session("Stat_KindX")="Donate_Amt" Then
          If J=0 Then
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(J)&"' ")
          ElseIf J=UBound(Ary_CodeDescX) Then
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And CONVERT(numeric,Donate_Amt)>='"&Ary_CodeDescX(J)&"' ")
          Else
            DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And CONVERT(numeric,Donate_Amt)>='"&Csng(Ary_CodeDescX(J-1))+1&"' And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(J)&"' ")
          End If
        Else
          DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And "&Session("Stat_KindX")&"='"&Ary_CodeDescX(J)&"'")
        End If
        Total_NoK=0
        Total_NoP=0
        SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
        '''Response.Write SQL1&"<br />"
        '''Response.End
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,1
        Total_NoK=Csng(RS1("Total_NoK"))
        RS1.Close
        Set RS1=Nothing
        If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf        
        Response.Flush
        Response.Clear
      Next

      If Check_X Then
        If Session("Stat_KindX")="Age" Then
          DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And Birthday Is Null ")
        ElseIf Session("Stat_KindX")="Donate_Amt" Then
          DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&"' And Donate_Amt Is Null ")
        Else
          DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" And ("&Session("Stat_KindX")&"='' Or "&Session("Stat_KindX")&" Is Null)")
        End If
        Total_NoK=0
        Total_NoP=0
        SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
        '''Response.Write SQL1&"<br />"
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,1
        Total_NoK=Csng(RS1("Total_NoK"))
        RS1.Close
        Set RS1=Nothing
        If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
        Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
        Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf
        Response.Flush
        Response.Clear
      End If

      DonateWhere=replace(Donate_Where,"Where 1=1",""&DonateWhere_Y&" ")
      Total_NoK=0
      Total_NoP=0
      SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
      '''Response.Write SQL1&"<br />"
      '''Response.End
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_NoK=Csng(RS1("Total_NoK"))
      RS1.Close
      Set RS1=Nothing
      If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf
      Response.Write "</tr>"&vbcrlf
      Response.Flush
      Response.Clear
    Next
    
    Check_Y=False
    If Session("Stat_KindY")="Age" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Birthday Is Null ")
    ElseIf Session("Stat_KindY")="Donate_Amt" Then
      DonateWhere=replace(Donate_Where,"Where 1=1","Where Donate_Amt Is Null ")
    Else
      DonateWhere=replace(Donate_Where,"Where 1=1","Where ("&Session("Stat_KindY")&"='' Or "&Session("Stat_KindY")&" Is Null)")
    End If
    Total_NoK=0
    Total_NoP=0
    SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    Total_NoK=Csng(RS1("Total_NoK"))
    RS1.Close
    Set RS1=Nothing
    If Total_NoK>0 Then
      Check_Y=True
      Response.Write "  <tr>"&vbcrlf
      Response.Write "    <td align='center' bgcolor='#D7E9F1' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKindY&"不詳</span></td>"&vbcrlf
      For I=0 To UBound(Ary_CodeDescX)
        DonateWhere=""
        If Session("Stat_KindX")="Age" Then
          If I=0 Then
            DonateWhere="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' "
          ElseIf I=UBound(Ary_CodeDescX) Then
            DonateWhere="Where Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' "
          Else
            DonateWhere="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(I-1))-1&"' "
          End If
        ElseIf Session("Stat_KindX")="Donate_Amt" Then
          If I=0 Then
            DonateWhere="Where CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(I)&"' "
          ElseIf I=UBound(Ary_CodeDescY) Then
            DonateWhere="Where CONVERT(numeric,Donate_Amt)>='"&Ary_CodeDescX(I)&"' "
          Else
            DonateWhere="Where CONVERT(numeric,Donate_Amt)>='"&Csng(Ary_CodeDescX(I-1))+1&"' And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(I)&"' "
          End If
        Else
          DonateWhere="Where "&Session("Stat_KindX")&"='"&Ary_CodeDescX(I)&"' And ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') "
        End If

        If Session("Stat_KindY")="Age" Then
          DonateWhere=DonateWhere&"And Birthday Is Null "
        ElseIf Session("Stat_KindY")="Donate_Amt" Then
          DonateWhere=DonateWhere&"And Donate_Amt Is Null "
        Else
          DonateWhere=DonateWhere&"And ("&Session("Stat_KindY")&"='' Or "&Session("Stat_KindY")&" Is Null) "
        End If
            
        Total_NoK=0
        Total_NoP=0
        DonateWhere=replace(Donate_Where,"Where 1=1",DonateWhere)
        SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
        '''Response.Write "SQL1:"&SQL1&"<br />"
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,1
        Total_NoK=Csng(RS1("Total_NoK"))
        RS1.Close
        Set RS1=Nothing
        If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf
        Response.Flush
        Response.Clear
      Next
      
      If Check_X Then
        DonateWhere=""
        If Session("Stat_KindX")="Age" Then
          DonateWhere="Where ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') And Birthday Is Null "
        ElseIf Session("Stat_KindX")="Donate_Amt" Then
          DonateWhere="Where ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') And Donate_Amt Is Null "
        Else
          DonateWhere="Where ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') And ("&Session("Stat_KindX")&" Is Null Or "&Session("Stat_KindX")&"='') "
        End If
        Total_NoK=0
        Total_NoP=0
        DonateWhere=replace(Donate_Where,"Where 1=1",DonateWhere)
        SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
        '''Response.Write "SQL1:"&SQL1&"<br />"
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,1
        Total_NoK=Csng(RS1("Total_NoK"))
        RS1.Close
        Set RS1=Nothing
        If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
        Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf
        Response.Flush
        Response.Clear
      End If

      DonateWhere=""
      If Session("Stat_KindY")="Age" Then
        DonateWhere="Where Birthday Is Null "
      ElseIf Session("Stat_KindX")="Donate_Amt" Then
        DonateWhere="Where Donate_Amt Is Null "
      Else
        DonateWhere="Where ("&Session("Stat_KindY")&" Is Null Or "&Session("Stat_KindY")&"='') "
      End If
      Total_NoK=0
      Total_NoP=0
      DonateWhere=replace(Donate_Where,"Where 1=1",DonateWhere)
      SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
      '''Response.Write "SQL1:"&SQL1&"<br />"
      '''Response.End
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_NoK=Csng(RS1("Total_NoK"))
      RS1.Close
      Set RS1=Nothing
      If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf
      Response.Flush
      Response.Clear
      
      Response.Write "  </tr>"&vbcrlf
      Response.Flush
      Response.Clear
    End If
    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td align='center' bgcolor='#FFFF99' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&StatKindX&"(X)小計</span></td>"&vbcrlf

    For I=0 To UBound(Ary_CodeDescX)
      DonateWhere=""
      If Session("Stat_KindX")="Age" Then
        If I=0 Then
          DonateWhere="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' "
        ElseIf I=UBound(Ary_CodeDescX) Then
          DonateWhere="Where Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' "
        Else
          DonateWhere="Where Year(Birthday)>='"&Year(Date())-Csng(Ary_CodeDescX(I))&"' And Year(Birthday)<='"&Year(Date())-Csng(Ary_CodeDescX(I-1))-1&"' "
        End If
      ElseIf Session("Stat_KindX")="Donate_Amt" Then
        If I=0 Then
          DonateWhere="Where CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(I)&"' "
        ElseIf I=UBound(Ary_CodeDescY) Then
          DonateWhere="Where CONVERT(numeric,Donate_Amt)>='"&Ary_CodeDescX(I)&"' "
        Else
          DonateWhere="Where CONVERT(numeric,Donate_Amt)>='"&Csng(Ary_CodeDescX(I-1))+1&"' And CONVERT(numeric,Donate_Amt)<='"&Ary_CodeDescX(I)&"' "
        End If
      Else
        DonateWhere="Where "&Session("Stat_KindX")&"='"&Ary_CodeDescX(I)&"' "
      End If
      DonateWhere=replace(Donate_Where,"Where 1=1 ",DonateWhere)
      Total_NoK=0
      Total_NoP=0
      SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
      '''Response.Write "SQL1:"&SQL1&"<br />"
      '''Response.End
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_NoK=Csng(RS1("Total_NoK"))
      RS1.Close
      Set RS1=Nothing
      If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
      Response.Write "  <td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf
      Response.Flush
      Response.Clear
    Next

    If Check_X Then
      If Session("Stat_KindX")="Age" Then
        DonateWhere=replace(Donate_Where,"Where 1=1","Where Birthday Is Null ")
      ElseIf Session("Stat_KindX")="Donate_Amt" Then
        DonateWhere=replace(Donate_Where,"Where 1=1","Where Donate_Amt Is Null ")
      Else
        DonateWhere=replace(Donate_Where,"Where 1=1","Where ("&Session("Stat_KindX")&"='' Or "&Session("Stat_KindX")&" Is Null)")
      End If
      Total_NoK=0
      Total_NoP=0
      SQL1="Select Total_NoK=Count(*) From Donate Join Donor On Donate.Donor_Id=Donor.Donor_Id "&DonateWhere&" "
      '''Response.Write SQL1&"<br />"
      '''Response.End
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Total_NoK=Csng(RS1("Total_NoK"))
      RS1.Close
      Set RS1=Nothing
      If Total_No>0 And Total_NoK>0 Then Total_NoP=Csng((Total_NoK/Total_No))
      Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total_NoK,0)&"</span></td>"&vbcrlf
      Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'>"&FormatPercent(Total_NoP,2)&"</span></td>"&vbcrlf
      Response.Flush
      Response.Clear
    End If
    Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'><b>"&FormatNumber(Total_No,0)&"</b></span></td>"&vbcrlf
    Response.Write "<td align='center' width='"&Width_Default/2&"' height='30'><span style='font-size: 9pt; font-family: 新細明體'><b>"&FormatPercent(1,2)&"</b></span></td>"&vbcrlf  
    Response.Write "  </tr>"&vbcrlf
    Response.Write "</table>"&vbcrlf
	
    Session.Contents.Remove("DeptId")
    Session.Contents.Remove("Donate_Date_Begin")
    Session.Contents.Remove("Donate_Date_End")
    Session.Contents.Remove("Act_Id")
    Session.Contents.Remove("Stat_KindX")
    Session.Contents.Remove("Stat_KindY")
    Session.Contents.Remove("Donate_Amt_Kind")
    Session.Contents.Remove("Donate_Purpose_Type")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->