<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=contribute.xls"
End If

Function ContributeList (SQL,ReportName)
  Contribute_Amt=0
  SQL1="Select Contribute_Amt=Isnull(Sum(Contribute_Amt),0) From CONTRIBUTE Where"&Split(Split(SQL,"Where")(1),"Order By")(0)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If Not RS1.EOF Then Contribute_Amt=RS1("Contribute_Amt")
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
    If I=3 Then Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>捐贈內容</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For I = 1 To FieldsCount
      If I=3 Then
        Response.Write "<td bgcolor='#FFFFFF' align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(i),0) & "</span></td>"
        Goods_Name=""
        SQL2="Select Goods_Name,Goods_Qty,Goods_Unit From CONTRIBUTEDATA Where Contribute_Id='"&RS1(0)&"' Order By Ser_No"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
        If Not RS2.EOF Then
          While Not RS2.EOF
            If Goods_Name="" Then
              Goods_Name=RS2("Goods_Name")&"("&FormatNumber(RS2("Goods_Qty"),0)&RS2("Goods_Unit")&")"
            Else
              Goods_Name=Goods_Name&"、"&RS2("Goods_Name")&"("&FormatNumber(RS2("Goods_Qty"),0)&RS2("Goods_Unit")&")"
            End If
            RS2.MoveNext
          Wend
          Goods_Name=Goods_Name&"。"  
        End If
        RS2.Close
        Set RS2=Nothing
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & Goods_Name & "</span></td>"
      Else
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & Data_Minus(RS1(I)) & "</span></td>"
	    End If
	  Next
	  Response.Flush
    Response.Clear
    RS1.MoveNext
	  Response.Write "</tr>"
  Wend
  Response.Write "<tr>"
  Response.Write "  <td colspan=""3"" align=""right"">折合現金合計："&FormatNumber(Contribute_Amt,0)&" 元</td>"
  Response.Write "  <td colspan="""&FieldsCount-2&"""></td>"
  Response.Write "</tr>"	
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function 
%>
<%Prog_Id="contribute"%>
<!--#include file="../include/head_rpt.asp"-->
<body class=tool>
<%
  SQL="Select Contribute_id,捐贈人=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),捐贈日期=CONVERT(nvarchar,Contribute_Date,111),折合現金=Isnull(Contribute_Amt,0),捐款方式=Contribute_Payment,捐款用途=Contribute_Purpose,收據開立=CONTRIBUTE.Invoice_Type,收據編號=Invoice_Pre+Invoice_No,列印=(Case When Invoice_Print='1' Then 'V' Else '' End),狀態=(Case When Issue_Type='M' Then '手開' Else Case When Issue_Type='D' Then '作廢' Else '' End End),經手人=CONTRIBUTE.Create_User " & _
      "From CONTRIBUTE Join DONOR On CONTRIBUTE.Donor_Id=DONOR.Donor_Id Left Join ACT On CONTRIBUTE.Act_Id=ACT.Act_Id Where 1=1 "
  If request("Dept_Id")<>"" Then SQL=SQL&"And CONTRIBUTE.Dept_Id = '"&request("Dept_Id")&"' "
  If request("Donor_Name")<>"" Then SQL=SQL&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or CONTRIBUTE.Invoice_Title Like '%"&request("Donor_Name")&"%') "
  If request("Contribute_Payment")<>"" Then SQL=SQL&"And Contribute_Payment = '"&request("Contribute_Payment")&"' "
  If request("Contribute_Purpose")<>"" Then SQL=SQL&"And Contribute_Purpose = '"&request("Contribute_Purpose")&"' "
  If request("Contribute_Date_B")<>"" Then SQL=SQL&"And Contribute_Date >= '"&request("Contribute_Date_B")&"' "
  If request("Contribute_Date_E")<>"" Then SQL=SQL&"And Contribute_Date <= '"&request("Contribute_Date_E")&"' "
  If request("Invoice_Type")<>"" Then SQL=SQL&"And CONTRIBUTE.Invoice_Type = '"&request("Invoice_Type")&"' "
  If request("Invoice_No_B")<>"" Then SQL=SQL&"And Invoice_No >= '"&request("Invoice_No_B")&"' "
  If request("Invoice_No_E")<>"" Then SQL=SQL&"And Invoice_No <= '"&request("Invoice_No_E")&"' "
  If request("Accoun_Date_B")<>"" Then SQL=SQL&"And Accoun_Date >= '"&request("Accoun_Date_B")&"' "
  If request("Accoun_Date_E")<>"" Then SQL=SQL&"And Accoun_Date <= '"&request("Accoun_Date_E")&"' "
  If request("Issue_TypeM")<>"" Then SQL=SQL&"And Issue_Type = '"&request("Issue_TypeM")&"' "
  If request("Issue_TypeD")<>"" Then SQL=SQL&"And Issue_Type = '"&request("Issue_TypeD")&"' "
  If request("Accounting_Title")<>"" Then SQL=SQL&"And Accounting_Title = '"&request("Accounting_Title")&"' "
  If request("Act_Id")<>"" Then SQL=SQL&"And Act_Id = '"&request("Act_Id")&"' "
  If request("Create_User")<>"" Then SQL=SQL&"And CONTRIBUTE.Create_User = '"&request("Create_User")&"' "
  SQL=SQL&"Order By CONVERT(VarChar,Contribute_Date,111) Desc,Contribute_id Desc"
  call ContributeList (SQL,Prog_Desc)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->