<%
'搜尋單筆收據資料
SQL1="Select Contribute_Id,Invoice_No=(Case When Contribute.Invoice_No_Old<>'' Then Contribute.Invoice_No_Old Else Invoice_Pre+Invoice_No End),Contribute_Date=CONVERT(nvarchar,Contribute_Date,111),Donor_Name=(Case When Contribute.Invoice_Title<>'' Then Contribute.Invoice_Title Else Donor_Name End),Title,Title2,ZipCode,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Address2=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+DONOR.ZipCode+C.mValue+Address Else B.mValue+DONOR.ZipCode+Address End End),Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+DONOR.Invoice_ZipCode+E.mValue+Invoice_Address Else D.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End),Contribute_Amt,CONTRIBUTE.Dept_Id,Invoice_IDNo,CONTRIBUTE.Donor_Id,Contribute_Purpose,Contribute_Payment,Invoice_Print,Invoice_Print_Add,Post_Donor_Name=Donor_Name,TEL=(Case When DONOR.Cellular_Phone<>'' Then DONOR.Cellular_Phone Else Case When DONOR.Tel_Office<>'' Then DONOR.Tel_Office Else DONOR.Tel_Home End End), " & _
     "Invoice_PrintComment=CONVERT(nvarchar(4000),Invoice_PrintComment),Comment=CONVERT(nvarchar(4000),Comment),Invoice_No_Old,CONTRIBUTE.Invoice_Type,CONTRIBUTE.Create_User " & _
     "From CONTRIBUTE Join DONOR On CONTRIBUTE.Donor_Id=DONOR.Donor_Id " & _
     "Left Join ACT On CONTRIBUTE.Act_Id=ACT.Act_Id " & _
     "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
     "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode Where Invoice_No<>'' And Contribute_Id='"&Request("Contribute_Id")&"' " & _
     "Order By Invoice_No "
Session("SQL1")=SQL1
If Request("action")="report" Then
  Response.Redirect "../include/movebar.asp?URL=../contribute/contribute_invoice_print_style"&Request("Invoice_Prog")&".asp"
ElseIf Request("action")="address" Then
  Response.Redirect "../include/movebar.asp?URL=../contribute/address_list_rpt_style"&request("Label")&".asp"
End If
%>
