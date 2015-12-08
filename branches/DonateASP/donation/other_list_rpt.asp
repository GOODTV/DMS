<%
'搜尋條件
SQL_Where=""
If Request("Donor_Name")<>"" Then SQL_Where=SQL_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "
If Request("Donor_Id_Begin")<>"" Then SQL_Where=SQL_Where&"And Donor.Donor_Id>='"&Request("Donor_Id_Begin")&"' "
If Request("Donor_Id_End")<>"" Then SQL_Where=SQL_Where&"And Donor.Donor_Id<='"&Request("Donor_Id_End")&"' "
If Request("Birthday_Month")<>"" Then SQL_Where=SQL_Where&"And Month(Birthday)='"&Request("Birthday_Month")&"' "
If Request("Donate_Over")<>"" Then SQL_Where=SQL_Where&"And Last_DonateDate<'"&DateAdd("M",Cint(Request("Donate_Over"))*-1,Date())&"' "
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
  If donorcategory<>"" Then SQL_Where=SQL_Where&"And Category In "&donorcategory&" "
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
  If donortype<>"" Then SQL_Where=SQL_Where&donortype
End If
If Request("City")<>"" Then SQL_Where=SQL_Where&"And City='"&Request("City")&"' "
If Request("Area")<>"" Then SQL_Where=SQL_Where&"And Area='"&Request("Area")&"' "
If Request("Invoice_City")<>"" Then SQL_Where=SQL_Where&"And City='"&Request("Invoice_City")&"' "
If Request("Invoice_Area")<>"" Then SQL_Where=SQL_Where&"And Invoice_Area='"&Request("Invoice_Area")&"' "
If Request("IsAbroad")<>"" Then SQL_Where=SQL_Where&"And (IsAbroad = 'Y' Or IsAbroad_Invoice = 'Y') "
If Request("IsSendNews")<>"" Then SQL_Where=SQL_Where&"And IsSendNews='Y' "
If Request("IsSendEpaper")<>"" Then SQL_Where=SQL_Where&"And IsSendEpaper='Y' "
If Request("IsSendYNews")<>"" Then SQL_Where=SQL_Where&"And IsSendYNews='Y' "
If Request("IsBirthday")<>"" Then SQL_Where=SQL_Where&"And IsBirthday='Y' "
If Request("IsXmas")<>"" Then SQL_Where=SQL_Where&"And IsXmas='Y' "
If Request("Donate_Total_Begin")<>"" Then SQL_Where=SQL_Where&"And CONVERT(numeric,A.Donate_Total)>='"&Request("Donate_Total_Begin")&"' "
If Request("Donate_Total_End")<>"" Then SQL_Where=SQL_Where&"And CONVERT(numeric,A.Donate_Total)<='"&Request("Donate_Total_End")&"' "
If Request("Donate_No_Begin")<>"" Then SQL_Where=SQL_Where&"And A.Donate_No>='"&Request("Donate_No_Begin")&"' "
If Request("Donate_No_End")<>"" Then SQL_Where=SQL_Where&"And A.Donate_No>='"&Request("Donate_No_End")&"' "

'排序方式
OrderBy_Where=""
If Request("OrderBy_Type")="1" Then
  If Request("Print_Desc")="1" Then
    OrderBy_Where=OrderBy_Where&"Order By ZipCode,A.Donor_Id"
  Else
    OrderBy_Where=OrderBy_Where&"Order By Invoice_ZipCode,A.Donor_Id"
  End If
ElseIf Request("OrderBy_Type")="2" Then
  OrderBy_Where=OrderBy_Where&"Order By A.Donor_Id"
End If

SQL1="Select Distinct A.Donor_Id,Donor_Name,Title=(Case When Title='' Then '君' Else Title End),ZipCode,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),Email,總次數=A.Donate_No,總金額=CONVERT(numeric,A.Donate_Total) " & _
     "From Donor Join (Select Donor_Id,Donate_No=Count(*),Donate_Total=Sum(Donate_Amt) From Donate Where 1=1 "
If Request("Dept_Id")<>"" Then
  SQL1=SQL1&"And Dept_Id = '"&Request("Dept_Id")&"' "
Else
  SQL1=SQL1&"And Dept_Id In ("&Session("all_dept_type")&")"
End If
If Request("Donate_Date_Begin")<>"" Then SQL1=SQL1&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "
If Request("Donate_Date_End")<>"" Then SQL1=SQL1&"And Donate_Date<='"&Request("Donate_Date_End")&"' "
If Request("Act_Id")<>"" Then SQL1=SQL1&"And Act_Id='"&Request("Act_Id")&"' "
SQL1=SQL1&"Group By Donor_Id) As A On DONOR.Donor_Id=A.Donor_Id " & _
"Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
"Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode Where (Address<>'' Or Invoice_Address<>'') " & _  
" "&SQL_Where&OrderBy_Where&" "

Session("SQL1")=SQL1
If Request("action")="report" Then
  Session("EmailMgr_Id")=Request("EmailMgr_Id")
  Response.Redirect "../include/movebar.asp?URL=../donation/other_list_rpt_style1.asp"
ElseIf Request("action")="address" Then
  Session("Print_Desc")=Request("Print_Desc")
  Response.Redirect "../include/movebar.asp?URL=../donation/address_list_rpt_style"&request("Label")&".asp"
End If
%>
