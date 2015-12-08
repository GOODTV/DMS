<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "big5"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=會員繳費資料明細表.xls"
End If

Function DonateList (SQL,ReportName)
  Donate_Amt=0
  SQL1="Select Donate_Amt=Sum(Donate_Amt) From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id Where"&Split(Split(SQL,"Where")(1),"Order By")(0)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If Not RS1.EOF Then Donate_Amt=RS1("Donate_Amt")
  RS1.Close
  Set RS1=Nothing
  Response.Write "<center><span style='font-size: 12pt; font-family: 標楷體'>"&ReportName&"</span></center>"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table x:str id=grid border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='98%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For I = 1 To FieldsCount
      If I=4 Then
        Response.Write "<td bgcolor='#FFFFFF' align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(i),0) & "</span></td>"
      Else
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i) & "</span></td>"
	    End If
	  Next
	  Response.Flush
    Response.Clear
    RS1.MoveNext
	  Response.Write "</tr>"
  Wend
  Response.Write "<tr>"
  Response.Write "  <td colspan=""4"" align=""right"">會員繳費合計："&FormatNumber(Donate_Amt,0)&" 元</td>"
  Response.Write "  <td colspan="""&FieldsCount-4&"""></td>"
  Response.Write "</tr>"	
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function 
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>會員繳費資料明細表</title>
</head>
<body class=tool>
<%
  SQL="Select Donate_Id,編號=DONATE.Donor_Id,捐款人=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),捐款日期=CONVERT(VarChar,Donate_Date,111),捐款金額=Donate_Amt,捐款方式=Donate_Payment,捐款用途=Donate_Purpose,勸募活動=Act_ShortName,收據開立=DONATE.Invoice_Type,收據編號=Invoice_Pre+Invoice_No,沖帳日期=CONVERT(VarChar,Accoun_Date,111),列印=(Case When Invoice_Print='1' Then 'V' Else '' End),狀態=(Case When Issue_Type='M' Then '手開' Else Case When Issue_Type='D' Then '作廢' Else '' End End),經手人=DONATE.Create_User " & _
      "From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id Left Join ACT On DONATE.Act_Id=ACT.Act_Id Where IsMember='Y' "
  If request("Dept_Id")<>"" Then
    SQL=SQL&"And DONATE.Dept_Id = '"&request("Dept_Id")&"' "
  Else
    SQL=SQL&"And DONATE.Dept_Id In ("&Session("all_dept_type")&") "
  End If
  If request("Donor_Name")<>"" Then SQL=SQL&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or DONATE.Invoice_Title Like '%"&request("Donor_Name")&"%') "
  If request("Member_No")<>"" Then SQL=SQL&"And Member_No Like '%"&request("Member_No")&"%' "
  If request("Donate_Payment")<>"" Then SQL=SQL&"And Donate_Payment = '"&request("Donate_Payment")&"' "
  If request("Donate_Purpose")<>"" Then
    SQL=SQL&"And Donate_Purpose = '"&request("Donate_Purpose")&"' "
  Else
    If DonatePurpose<>"" Then SQL=SQL&"And Donate_Purpose In ("&DonatePurpose&") "
  End If
  If request("Donate_Date_B")<>"" Then SQL=SQL&"And Donate_Date >= '"&request("Donate_Date_B")&"' "
  If request("Donate_Date_E")<>"" Then SQL=SQL&"And Donate_Date <= '"&request("Donate_Date_E")&"' "
  If request("Invoice_Type")<>"" Then SQL=SQL&"And DONATE.Invoice_Type = '"&request("Invoice_Type")&"' "
  If request("Invoice_No_B")<>"" Then SQL=SQL&"And Invoice_No >= '"&request("Invoice_No_B")&"' "
  If request("Invoice_No_E")<>"" Then SQL=SQL&"And Invoice_No <= '"&request("Invoice_No_E")&"' "
  If request("Accoun_Date_B")<>"" Then SQL=SQL&"And Accoun_Date >= '"&request("Accoun_Date_B")&"' "
  If request("Accoun_Date_E")<>"" Then SQL=SQL&"And Accoun_Date <= '"&request("Accoun_Date_E")&"' "
  If request("Donate_Type")<>"" Then SQL=SQL&"And Donate_Type = '"&request("Donate_Type")&"' "
  If request("Accoun_Bank")<>"" Then SQL=SQL&"And Accoun_Bank = '"&request("Accoun_Bank")&"' "
  If request("Donation_NumberNo")<>"" Then SQL=SQL&"And Donation_NumberNo Like '%"&request("Donation_NumberNo")&"%' "
  If request("Donation_SubPoenaNo")<>"" Then SQL=SQL&"And Donation_SubPoena Like '%"&request("Donation_SubPoenaNo")&"%' "
  If request("Accounting_Title")<>"" Then SQL=SQL&"And Accounting_Title = '"&request("Accounting_Title")&"' "
  If request("Act_Id")<>"" Then SQL=SQL&"And DONATE.Act_Id = '"&request("Act_Id")&"' "
  If request("Create_User")<>"" Then SQL=SQL&"And DONATE.Create_User = '"&request("Create_User")&"' "
  If request("Issue_TypeM")<>"" Then SQL=SQL&"And Issue_Type = '"&request("Issue_TypeM")&"' "
  If request("Issue_TypeD")<>"" Then SQL=SQL&"And Issue_Type = '"&request("Issue_TypeD")&"' "  
  Donate_Purpose_Type=""
  Ary_Purpose_Type=Split(request("Donate_Purpose_Type"),",")
  For I = 0 To UBound(Ary_Purpose_Type)
    If Donate_Purpose_Type="" Then
      Donate_Purpose_Type="'"&Trim(Ary_Purpose_Type(I))&"'"
    Else
      Donate_Purpose_Type=Donate_Purpose_Type&",'"&Trim(Ary_Purpose_Type(I))&"'"
    End If
  Next
  If Donate_Purpose_Type<>"" Then SQL=SQL&"And Donate_Purpose_Type In ("&Donate_Purpose_Type&") "
  ReportName="會員繳費資料明細表"
  call DonateList (SQL,ReportName)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->