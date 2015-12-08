<%
'搜尋條件
SQL_Where=""
If Request("Donor_Name")<>"" Then SQL_Where=SQL_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "
If Request("Donor_Id_Begin")<>"" Then SQL_Where=SQL_Where&"And Donor.Donor_Id >= '"&Request("Donor_Id_Begin")&"' "
If Request("Donor_Id_End")<>"" Then SQL_Where=SQL_Where&"And Donor.Donor_Id <= '"&Request("Donor_Id_End")&"' "
If Request("action")="report" Then
  If Request("RePrint")<>"" Then
    SQL_Where=SQL_Where&"And IsThanks='1' "
  Else
    SQL_Where=SQL_Where&"And (IsThanks='0' Or IsThanks Is Null) "
  End If
End If
If Request("action")="address" Then
  If Request("ReAddPrint")<>"" Then 
    SQL_Where=SQL_Where&"And IsThanks_Add='1' "
  Else
    SQL_Where=SQL_Where&"And (IsThanks_Add='0' Or IsThanks Is Null) "
  End If
End If
If Request("Dept_Id")<>"" Then
  SQL_Where=SQL_Where&"And A.Dept_Id = '"&Request("Dept_Id")&"' "
Else
  SQL_Where=SQL_Where&"And A.Dept_Id In ("&Session("all_dept_type")&") "
End If
If Request("Donate_Date_Begin")<>"" Then SQL_Where=SQL_Where&"And A.Donate_Date>='"&Request("Donate_Date_Begin")&"' "
If Request("Donate_Date_End")<>"" Then SQL_Where=SQL_Where&"And A.Donate_Date<='"&Request("Donate_Date_End")&"' "
If Request("Act_Id")<>"" Then SQL_Where=SQL_Where&"And A.Act_Id='"&Request("Act_Id")&"' "

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
End If

SQL1="Select Distinct A.Dept_Id,A.Donor_Id,Donor_Name,Title=(Case When Title='' Then '君' Else Title End),ZipCode,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),IsThanks,IsThanks_Add " & _
     "From Donor Join DONATE A On Donor.Donor_Id=A.Donor_Id " & _
     "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
     "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode " & _
     "Where (Address<>'' Or Invoice_Address<>'') "&SQL_Where&OrderBy_Where&" "

Session("Print_Desc")=Request("Print_Desc")
Session("SQL1")=SQL1
If Request("action")="report" Then
  Session("EmailMgr_Id")=Request("EmailMgr_Id")
  Response.Redirect "../include/movebar.asp?URL=../donation/thanks_list_rpt_style1.asp"
ElseIf Request("action")="address" Then
  Response.Redirect "../include/movebar.asp?URL=../donation/address_list_rpt_style"&request("Label")&".asp"
End If
%>
