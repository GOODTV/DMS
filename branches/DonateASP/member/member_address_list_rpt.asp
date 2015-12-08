<%
If Request("action")="help" Then Response.Redirect "member_address_list_help.htm"

'搜尋條件一
Donate_Where=""
If Request("Dept_Id")<>"" Then
  Donate_Where=Donate_Where&"And A.Dept_Id = '"&Request("Dept_Id")&"' "
Else
  Donate_Where=Donate_Where&"And A.Dept_Id In ("&Session("all_dept_type")&") "
End If
If Request("Donate_Date_Begin")<>"" Then Donate_Where=Donate_Where&"And A.Donate_Date>='"&Request("Donate_Date_Begin")&"' "
If Request("Donate_Date_End")<>"" Then Donate_Where=Donate_Where&"And A.Donate_Date<='"&Request("Donate_Date_End")&"' "
If Request("Act_Id")<>"" Then Donate_Where=Donate_Where&"And A.Act_Id='"&Request("Act_Id")&"' "
If Request("Invoice_No_Begin")<>"" Then Donate_Where=Donate_Where&"And A.Invoice_No>='"&Request("Invoice_No_Begin")&"' "
If Request("Invoice_No_End")<>"" Then Donate_Where=Donate_Where&"And A.Invoice_No<='"&Request("Invoice_No_End")&"' "
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
  If donatepayment<>"" Then Donate_Where=Donate_Where&"And A.Donate_Payment In "&donatepayment&" "
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
  If donatepurpose<>"" Then Donate_Where=Donate_Where&"And A.Donate_Purpose In "&donatepurpose&" "
End If
If Request("Invoice_Type")<>"" Then
  donateinvoice=""
  Ary_Invoice_Type=Split(Request("Invoice_Type"),",")
  If UBound(Ary_Invoice_Type)=0 Then
    donateinvoice=donateinvoice&"('"&Trim(Ary_Invoice_Type(0))&"')"
  Else
    For I = 0 To UBound(Ary_Invoice_Type)
      If I=0 Then
        donateinvoice=donateinvoice&"('"&Trim(Ary_Invoice_Type(I))&"',"
      ElseIf I=UBound(Ary_Invoice_Type) Then
        donateinvoice=donateinvoice&"'"&Trim(Ary_Invoice_Type(I))&"')" 
      Else
        donateinvoice=donateinvoice&"'"&Trim(Ary_Invoice_Type(I))&"',"
      End If
    Next
  End If
  If donateinvoice<>"" Then Donate_Where=Donate_Where&"And A.Invoice_Type In "&donateinvoice&" "
End If
If Request("Donate_Type")<>"" Then
  donatetype=""
  Ary_Donate_Type=Split(Request("Donate_Type"),",")
  If UBound(Ary_Donate_Type)=0 Then
    donatetype=donatetype&"('"&Trim(Ary_Donate_Type(0))&"')"
  Else
    For I = 0 To UBound(Ary_Donate_Type)
      If I=0 Then
        donatetype=donatetype&"('"&Trim(Ary_Donate_Type(I))&"',"
      ElseIf I=UBound(Ary_Donate_Type) Then
        donatetype=donatetype&"'"&Trim(Ary_Donate_Type(I))&"')" 
      Else
        donatetype=donatetype&"'"&Trim(Ary_Donate_Type(I))&"',"
      End If
    Next
  End If
  If donatetype<>"" Then Donate_Where=Donate_Where&"And A.Donate_Type In "&donatetype&" "
End If

'搜尋條件二
Donor_Where="And IsMember='Y' "
If Request("Category")<>"" Then
  donorcategory=""
  Ary_Category=Split(Request("Category"),",")
  If UBound(Ary_Category)=0 Then
    donorcategory=donorcategory&"('"&Trim(Ary_Category(0))&"')"
  Else
    For I = 0 To UBound(Ary_Category)
      If I=0 Then
        donorcategory=donorcategory&"('"&Trim(Ary_Category(I))&"',"
      ElseIf I=UBound(Ary_Category) Then 
        donorcategory=donorcategory&"'"&Trim(Ary_Category(I))&"')" 
      Else
        donorcategory=donorcategory&"'"&Trim(Ary_Category(I))&"',"
      End If
    Next
  End If
  If donorcategory<>"" Then Donor_Where=Donor_Where&"And Category In "&donorcategory&" "
End If
If Request("Donor_Type")<>"" Then
  donortype=""
  Ary_Donor_Type=Split(Request("Donor_Type"),",")
  If UBound(Ary_Donor_Type)=0 Then
    donortype=donortype&"And Donor_Type Like '%"&Trim(Ary_Donor_Type(0))&"%' "
  Else
    For I = 0 To UBound(Ary_Donor_Type)
      If I=0 Then
        donortype=donortype&"And (Donor_Type Like '%"&Trim(Ary_Donor_Type(I))&"%' "
      ElseIf I=UBound(Ary_Donor_Type) Then 
        donortype=donortype&"Or Donor_Type Like '%"&Trim(Ary_Donor_Type(I))&"%') "
      Else
        donortype=donortype&"Or Donor_Type Like '%"&Trim(Ary_Donor_Type(I))&"%' "
      End If
    Next
  End If
  If donortype<>"" Then Donor_Where=Donor_Where&donortype
End If
If Request("Member_Type")<>"" Then SQL=SQL&"And Member_Type='"&Request("Member_Type")&"' "
If Request("Member_Status")<>"" Then SQL=SQL&"And Member_Status='"&Request("Member_Status")&"' "
If Request("Member_JoinYear")<>"" Then SQL=SQL&"And Year(Member_JoinDate)='"&Request("Member_JoinYear")&"' "
If Request("Member_Group")<>"" Then SQL=SQL&"And Member_Group Like '%"&Request("Member_Group")&"%' "
If Request("City")<>"" Then Donor_Where=Donor_Where&"And City='"&Request("City")&"' "
If Request("Area")<>"" Then Donor_Where=Donor_Where&"And Area='"&Request("Area")&"' "
If Request("Invoice_City")<>"" Then Donor_Where=Donor_Where&"And City='"&Request("Invoice_City")&"' "
If Request("Invoice_Area")<>"" Then Donor_Where=Donor_Where&"And Invoice_Area='"&Request("Invoice_Area")&"' "
If Request("Donate_Total_Begin")<>"" And Request("Print_Type")<>"1" Then Donor_Where=Donor_Where&"And CONVERT(numeric,A.Donate_Total)>='"&Request("Donate_Total_Begin")&"' "
If Request("Donate_Total_End")<>"" And Request("Print_Type")<>"1" Then Donor_Where=Donor_Where&"And CONVERT(numeric,A.Donate_Total)<='"&Request("Donate_Total_End")&"' "
If Request("Donate_No_Begin")<>"" And Request("Print_Type")<>"1" Then Donor_Where=Donor_Where&"And A.Donate_No>='"&Request("Donate_No_Begin")&"' "
If Request("Donate_No_End")<>"" And Request("Print_Type")<>"1" Then Donor_Where=Donor_Where&"And A.Donate_No>='"&Request("Donate_No_End")&"' "
If Request("Donor_Name")<>"" Then Donor_Where=Donor_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "
If Request("Donor_Id_Begin")<>"" Then SQL=SQL&"And Donor_Id>='"&Request("Donor_Id_Begin")&"' "
If Request("Donor_Id_End")<>"" Then SQL=SQL&"And Donor_Id<='"&Request("Donor_Id_End")&"' "
If Request("Birthday_Month")<>"" Then SQL=SQL&"And Month(Birthday)='"&Request("Birthday_Month")&"' "
If Request("Donate_Over")<>"" Then SQL=SQL&"And Last_DonateDateD<'"&DateAdd("M",Cint(Request("Donate_Over"))*-1,Date())&"' "                              
If Request("Print_Type")="2" Then Donor_Where=Donor_Where&"And IsSendNews='Y' "  '郵寄會訊
If Request("Print_Type")="3" Then Donor_Where=Donor_Where&"And IsSendYNews='Y' " '郵寄年報
If Request("Print_Type")="4" Then Donor_Where=Donor_Where&"And IsBirthday='Y' " '郵寄生日卡
If Request("Print_Type")="5" Then Donor_Where=Donor_Where&"And IsXmas='Y' " '郵寄賀年(耶誕)卡

'排序方式
OrderBy_Where=""
If Request("OrderBy_Type")="1" Then
  If Request("Print_Desc")="1" Then
    OrderBy_Where=OrderBy_Where&"Order By Invoice_ZipCode,Invoice_Address,Donor.Donor_Id Desc"
  Else
    OrderBy_Where=OrderBy_Where&"Order By ZipCode,Address,Donor.Donor_Id Desc"
  End If
ElseIf Request("OrderBy_Type")="2" Then
  OrderBy_Where=OrderBy_Where&"Order By Donor.Donor_Id"
ElseIf Request("OrderBy_Type")="3" Then
  If Request("Print_Type")="1" Then
    OrderBy_Where=OrderBy_Where&"Order By Donor.Dept_Id,Invoice_No"
  Else
    If Request("Print_Desc")="1" Then
      OrderBy_Where=OrderBy_Where&"Order By Invoice_ZipCode,Invoice_Address,Donor.Donor_Id Desc"
    Else
      OrderBy_Where=OrderBy_Where&"Order By ZipCode,Address,Donor.Donor_Id Desc"
    End If
  End If  
End If

'搜尋地址名條資料
If Request("Print_Type")="1" Then                    
  SQL1="Select Distinct Donor.Dept_Id,Donor.Donor_Id,Donor_Name,Title=(Case When Title='' Then '君' Else Title End),ZipCode,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),Invoice_Pre,Invoice_No " & _
       "From Donor Join DONATE A On Donor.Donor_Id=A.Donor_Id " & _
       "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
       "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode " & _
       "Where Invoice_No<>'' And A.Invoice_Type Not In ('"&Request("InvoiceTypeN")&"','"&Request("InvoiceTypeY")&"') "
  If Request("Print_Desc")="1" Then
    SQL1=SQL1&"And Invoice_Address<>''" & Donate_Where & Donor_Where & OrderBy_Where
  Else
    SQL1=SQL1&"And Address<>''" & Donate_Where & Donor_Where & OrderBy_Where
  End If
Else
  Join_Type="Left Join"
  If Request("Donate_Date_Begin")<>"" Or Request("Donate_Date_End")<>"" Or Request("Donate_Total_Begin")<>"" Or Request("Donate_Total_End")<>"" Or Request("Act_Id")<>"" Or Request("Donate_No_Begin")<>"" Or Request("Donate_No_End")<>"" Or Request("Invoice_No_Begin")<>"" Or Request("Invoice_No_End")<>"" Or Request("Donate_Payment")<>"" Or Request("Donate_Purpose")<>"" Or Request("Donate_Type")<>"" Then Join_Type="Join"
  SQL1="Select Distinct Donor.Dept_Id,Donor.Donor_Id,Donor_Name,Title=(Case When Title='' Then '君' Else Title End),ZipCode,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End) " & _
       "From Donor "&Join_Type&" (Select Dept_Id,Donor_Id,Donate_No=Count(*),Donate_Total=Sum(Donate_Amt) From Donate Where 1=1 "&replace(Donate_Where,"A.","Donate.")&" " &_
       "Group By Dept_Id,Donor_Id) As A On Donor.Donor_Id=A.Donor_Id " & _
       "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
       "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode "
  If Request("Print_Desc")="1" Then
    SQL1=SQL1&"Where Invoice_Address<>'' " & Donor_Where & OrderBy_Where
  Else
    SQL1=SQL1&"Where Address<>'' " & Donor_Where & OrderBy_Where
  End If
End If
Session("Print_Type")=Request("Print_Type")
Session("Print_Desc")=Request("Print_Desc")
Session("SQL1")=SQL1
If Request("action")="report" Then
  Response.Redirect "../include/movebar.asp?URL=../donation/address_list_rpt_style"&request("Label")&".asp"
ElseIf Request("action")="post" Then
  Session("DeptId")=Request("Dept_Id")
  Response.Redirect "../include/movebar.asp?URL=../donation/address_list_post.asp"
End If  
%>
