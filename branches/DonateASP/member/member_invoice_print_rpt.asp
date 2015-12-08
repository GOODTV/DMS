<%
If Request("action")="help" Then Response.Redirect "invoice_help.htm"
'搜尋條件一
Donate_Where=""
If Request("Invoice_Type")<>"" Then
  invoicetype=""
  Ary_Invoice_Type=Split(Request("Invoice_Type"),",")
  If UBound(Ary_Invoice_Type)=0 Then
    invoicetype=invoicetype&"('"&Trim(Ary_Invoice_Type(0))&"')"
  Else
    For I = 0 To UBound(Ary_Purpose)
      If I=0 Then
        invoicetype=invoicetype&"('"&Trim(Ary_Invoice_Type(I))&"',"
      ElseIf I=UBound(Ary_Purpose) Then
        invoicetype=invoicetype&"'"&Trim(Ary_Invoice_Type(I))&"')" 
      Else
        invoicetype=invoicetype&"'"&Trim(Ary_Invoice_Type(I))&"',"
      End If
    Next
  End If
  If invoicetype<>"" Then Donate_Where=Donate_Where&"And Donate.Invoice_Type In "&invoicetype&" "
End If
If Request("Dept_Id")<>"" Then
  Donate_Where=Donate_Where&"And Donate.Dept_Id = '"&Request("Dept_Id")&"' "
Else
  Donate_Where=Donate_Where&"And Donate.Dept_Id In ("&Session("all_dept_type")&") "
End If
If Request("Create_User")<>"" Then Donate_Where=Donate_Where&"And Donate.Create_User = '"&Request("Create_User")&"' "
If Request("Donate_Id")<>"" Then Donate_Where=Donate_Where&"And Donate.Donate_Id = '"&Request("Donate_Id")&"' "
If Request("Donate_Date_Begin")<>"" Then Donate_Where=Donate_Where&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "
If Request("Donate_Date_End")<>"" Then Donate_Where=Donate_Where&"And Donate_Date<='"&Request("Donate_Date_End")&"' "
If Request("action")="report" Then
  If Request("RePrint")<>"" Then
    Donate_Where=Donate_Where&"And Invoice_Print='1' "
  Else
    Donate_Where=Donate_Where&"And (Invoice_Print='0' Or Invoice_Print Is Null) "
  End If
End If
If Request("action")="address" Then
  If Request("ReAddPrint")<>"" Then 
    Donate_Where=Donate_Where&"And Invoice_Print_Add='1' "
  Else
    Donate_Where=Donate_Where&"And (Invoice_Print_Add='0' Or Invoice_Print_Add Is Null Or Issue_Type='M') "
  End If
End If
If Request("Act_Id")<>"" Then Donate_Where=Donate_Where&"And DONATE.Act_Id='"&Request("Act_Id")&"' "
If Request("Invoice_No_Begin")<>"" Then Donate_Where=Donate_Where&"And Invoice_No>='"&Request("Invoice_No_Begin")&"' "
If Request("Invoice_No_End")<>"" Then Donate_Where=Donate_Where&"And Invoice_No<='"&Request("Invoice_No_End")&"' "
If Request("Donate_Payment")<>"" Then
  donatepayment=""
  Ary_Payment=Split(Request("Donate_Payment"),",")
  If UBound(Ary_Payment)=0 Then
    donatepayment=donatepayment&"('"&Trim(Ary_Payment(0))&"')"
  Else
    For I = 0 To UBound(Ary_Payment)
      If I=0 Then
        donatepayment=donatepayment&"('"&Trim(Ary_Payment(I))&"',"
      ElseIf I=UBound(Ary_Payment) Then
        donatepayment=donatepayment&"'"&Trim(Ary_Payment(I))&"')" 
      Else
        donatepayment=donatepayment&"'"&Trim(Ary_Payment(I))&"',"
      End If
    Next
  End If
  If donatepayment<>"" Then Donate_Where=Donate_Where&"And Donate_Payment In "&donatepayment&" "
End If
If Request("Donate_Purpose")<>"" Then
  donatepurpose=""
  Ary_Purpose=Split(Request("Donate_Purpose"),",")
  If UBound(Ary_Purpose)=0 Then
    donatepurpose=donatepurpose&"('"&Trim(Ary_Purpose(0))&"')"
  Else
    For I = 0 To UBound(Ary_Purpose)
      If I=0 Then
        donatepurpose=donatepurpose&"('"&Trim(Ary_Purpose(I))&"',"
      ElseIf I=UBound(Ary_Purpose) Then
        donatepurpose=donatepurpose&"'"&Trim(Ary_Purpose(I))&"')" 
      Else
        donatepurpose=donatepurpose&"'"&Trim(Ary_Purpose(I))&"',"
      End If
    Next
  End If
  If donatepurpose<>"" Then Donate_Where=Donate_Where&"And Donate_Purpose In "&donatepurpose&" "
End If
If Request("Issue_TypeM")<>"" Or Request("Issue_TypeD")<>"" Then
  If Request("Issue_TypeM")<>"" Then Donate_Where=Donate_Where&"And Issue_Type='M' "
  If Request("Issue_TypeD")<>"" Then Donate_Where=Donate_Where&"And Issue_Type='D' "
Else
  Donate_Where=Donate_Where&"And (Issue_Type='M' Or Issue_Type='' Or Issue_Type Is null) "
End If
If Request("Print_Pint")="1" Then
  Donate_Where=Donate_Where&"And Donate_Payment Not In ('"&Request("DonateDesc")&"') "
ElseIf Request("Print_Pint")="2" Then
  Donate_Where=Donate_Where&"And Donate_Payment In ('"&Request("DonateDesc")&"') "
End If

'搜尋條件二
Donor_Where=""
If Request("Donor_Name")<>"" Then Donor_Where=Donor_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "

'排序方式
OrderBy_Where=""
If Request("OrderBy_Type")="1" Then
  OrderBy_Where=OrderBy_Where&"Order By Donate.Dept_Id,Invoice_No"
ElseIf Request("OrderBy_Type")="2" Then
  OrderBy_Where=OrderBy_Where&"Order By Member_No,Donate.Donor_Id,Invoice_No"
ElseIf Request("OrderBy_Type")="3" Then
  If Request("Print_Desc")="1" Then
    OrderBy_Where=OrderBy_Where&"Order By ZipCode,Donate.Dept_Id,Invoice_No"
  Else
    OrderBy_Where=OrderBy_Where&"Order By Invoice_ZipCode,Donate.Dept_Id,Invoice_No"
  End If
End If

'搜尋單筆收據資料
SQL1="Select Donate_Id,Invoice_No=(Case When Donate.Invoice_No_Old<>'' Then Donate.Invoice_No_Old Else Invoice_Pre+Invoice_No End),Donate_Date=CONVERT(VarChar,Donate_Date,111),Donor_Name=(Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End),Title,Title2,ZipCode,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Address2=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+DONOR.ZipCode+C.mValue+Address Else B.mValue+DONOR.ZipCode+Address End End),Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+DONOR.Invoice_ZipCode+E.mValue+Invoice_Address Else D.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End),Donate_Amt,Donate_Amt2,Donate_Desc=CONVERT(nvarchar(4000),Donate_Desc),Donate.Dept_Id,Invoice_IDNo,DONATE.Donor_Id,Donate_Purpose,Donate_Payment,Invoice_Print,Invoice_Print_Add,Post_Donor_Name=Donor_Name,TEL=(Case When DONOR.Cellular_Phone<>'' Then DONOR.Cellular_Phone Else Case When DONOR.Tel_Office<>'' Then DONOR.Tel_Office Else DONOR.Tel_Home End End), " & _
     "Invoice_PrintComment=CONVERT(nvarchar(4000),Invoice_PrintComment),Comment=CONVERT(nvarchar(4000),Comment),Invoice_No_Old,Donate.Invoice_Type,Donate.Create_User,Donate_Forign " & _
     "From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id " & _
     "Left Join ACT On DONATE.Act_Id=ACT.Act_Id " & _
     "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
     "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode Where IsMember='Y' And Donate_Purpose_Type In ('M','A') And Invoice_No<>'' "
If Request("action")="address" Then
  If Request("Print_Desc")="1" Then
    SQL1=SQL1&"And Invoice_Address<>'' "
  Else
    SQL1=SQL1&"And Address<>'' "
  End If
End If
SQL1=SQL1&" "&Donate_Where&Donor_Where&OrderBy_Where&" "

Session("Rept_Licence")=Request("Rept_Licence")
Session("Print_Desc")=Request("Print_Desc")
Session("SQL1")=SQL1
If Request("action")="report" Then
  If Request("Print_Pint")="1" Then
    '捐款收據
    Response.Redirect "../include/movebar.asp?URL=../donation/invoice_print_style"&Request("Invoice_Prog")&".asp"
  ElseIf Request("Print_Pint")="2" Then
    '捐物收據
    Response.Redirect "../include/movebar.asp?URL=../donation/invoice_print_style"&Request("Invoice_Prog")&".asp"
  End If
ElseIf Request("action")="address" Then
  Response.Redirect "../include/movebar.asp?URL=../donation/address_list_rpt_style"&request("Label")&".asp"
ElseIf Request("action")="post" Then
  Session("DeptId")=Request("Dept_Id")
  Response.Redirect "../include/movebar.asp?URL=../donation/invoice_print_post.asp"  
End If
%>
