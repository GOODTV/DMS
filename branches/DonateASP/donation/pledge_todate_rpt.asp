<%
'搜尋條件
SQL_Where=""
If Request("Donor_Name")<>"" Then SQL_Where=SQL_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "
If Request("Dept_Id")<>"" Then
  SQL_Where=SQL_Where&"And A.Dept_Id = '"&Request("Dept_Id")&"' "
Else
  SQL_Where=SQL_Where&"And A.Dept_Id In ("&Session("all_dept_type")&") "
End If
If request("To_Year")<>"" Then SQL_Where=SQL_Where & "And Year(Donate_ToDate) = '"&request("To_Year")&"' "
If request("To_Month")<>"" Then SQL_Where=SQL_Where & "And Month(Donate_ToDate) = '"&request("To_Month")&"' "
If request("Status")<>"" Then SQL_Where=SQL_Where & "And Status = '"&request("Status")&"' "
If request("Donate_Type")<>"" Then SQL_Where=SQL_Where & "And Donate_Type = '"&request("Donate_Type")&"' "
If request("Pledge_Id_Begin")<>"" Then SQL_Where=SQL_Where & "And Pledge_Id >= '"&request("Pledge_Id_Begin")&"' "
If request("Pledge_Id_End")<>"" Then SQL_Where=SQL_Where & "And Pledge_Id <= '"&request("Pledge_Id_End")&"' "
  
'排序方式
OrderBy_Where=""
If Request("OrderBy_Type")="1" Then
  If Request("Print_Desc")="1" Then
    OrderBy_Where=OrderBy_Where&"Order By ZipCode,A.Dept_Id,A.Donor_Id"
  Else
    OrderBy_Where=OrderBy_Where&"Order By Invoice_ZipCode,A.Dept_Id,A.Donor_Id"
  End If
ElseIf Request("OrderBy_Type")="2" Then
  OrderBy_Where=OrderBy_Where&"Order By A.Donor_Id"
ElseIf Request("OrderBy_Type")="3" Then
  OrderBy_Where=OrderBy_Where&"Order By Pledge_Id Desc"
End If

SQL1="Select Distinct A.Dept_Id,A.Donor_Id,Donor_Name,Title=(Case When Title='' Then '君' Else Title End),ZipCode,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),Pledge_Id,Email,Donate_ToDate " & _
     "From Donor Join PLEDGE A On Donor.Donor_Id=A.Donor_Id " & _
     "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
     "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode " & _
     "Where (Address<>'' Or Invoice_Address<>'') "&SQL_Where&OrderBy_Where&" "

Session("SQL1")=SQL1
If Request("action")="report" Then
  Session("EmailMgr_Id")=Request("EmailMgr_Id")
  Response.Redirect "../include/movebar.asp?URL=../donation/pledge_todate_rpt_style1.asp"
ElseIf Request("action")="address" Then
  Session("Print_Desc")=Request("Print_Desc")
  Response.Redirect "../include/movebar.asp?URL=../donation/address_list_rpt_style"&request("Label")&".asp"
End If
%>
