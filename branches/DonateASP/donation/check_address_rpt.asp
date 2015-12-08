<%
SQL1="Select Donor_Id,Donor_Name,Invoice_Title, " & _
     "Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End), " & _
     "Invoice_Address2=(Case When DONOR.Invoice_City='' Then DONOR.Invoice_Address Else Case When C.mValue<>B.mValue Then C.mValue+DONOR.Invoice_ZipCode+D.mValue+DONOR.Invoice_Address Else C.mValue+DONOR.Invoice_ZipCode+DONOR.Invoice_Address End End), " & _
     "City,Area,Address,Invoice_City,Invoice_Area,Invoice_Address " & _
     "From DONOR " & _
     "Left Join CodeCity As A On DONOR.City=A.mCode " & _
     "Left Join CodeCity As B On DONOR.Area=B.mCode " & _ 
     "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
     "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
     "Where CONVERT(int,DONOR.Donor_Id)>='"&Request("Donor_Id_Begin")&"' And CONVERT(int,DONOR.Donor_Id)<='"&Request("Donor_Id_End")&"' "
If Request("Check_Type")="IsSendNews" Then SQL1=SQL1&"And IsSendNews='Y' "
If Request("Check_Type")="IsSendYNews" Then SQL1=SQL1&"And IsSendYNews='Y' "
If Request("Check_Type")="IsBirthday" Then SQL1=SQL1&"And IsBirthday='Y' "
If Request("Check_Type")="IsXmas" Then SQL1=SQL1&"And IsXmas='Y' "
SQL1=SQL1&"Order By Donor_Id Desc"

Session("Check_Type")=Request("Check_Type")
Session("Address_Type")=Request("Address_Type")
Session("SQL1")=SQL1
If Request("Check_Type")="IsDonor" Then
  Response.Redirect "../include/movebar.asp?URL=../donation/check_address_rpt_style1.asp"
Else
  Response.Redirect "../include/movebar.asp?URL=../donation/check_address_rpt_style2.asp"
End If  
%>
