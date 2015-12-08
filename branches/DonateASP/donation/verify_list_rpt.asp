<%
'搜尋條件
SQL_Where=""
If Request("Donor_Name")<>"" Then SQL_Where=SQL_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "
If Request("Dept_Id")<>"" Then
  SQL_Where=SQL_Where&"And Donate.Dept_Id = '"&Request("Dept_Id")&"' "
Else
  SQL_Where=SQL_Where&"And Donate.Dept_Id In ("&Session("all_dept_type")&") "
End If
If Request("Donate_Date_Begin")<>"" Then SQL_Where=SQL_Where&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "
If Request("Donate_Date_End")<>"" Then SQL_Where=SQL_Where&"And Donate_Date<='"&Request("Donate_Date_End")&"' "
If Request("Act_Id")<>"" Then SQL_Where=SQL_Where&"And Donate.Act_Id='"&Request("Act_Id")&"' "

'排序方式
OrderBy_Where=""
If Request("List_Type")="1" Then
  OrderBy_Where=OrderBy_Where&"Order By Donate_Date Desc,Donate_Id Desc"
ElseIf Request("List_Type")="2" Then
  OrderBy_Where=OrderBy_Where&"Order By Donate_Amt Desc,編號抬頭,Donate_Date Desc,Donate_Id Desc"
ElseIf Request("List_Type")="3" Then
  OrderBy_Where=OrderBy_Where&"Order By CASECODE.Seq,Donate_Date Desc,Donate_Id Desc"
ElseIf Request("List_Type")="4" Then
  OrderBy_Where=OrderBy_Where&"Order By Donate_Amt Desc,捐款人,Donate_Date Desc,Donate_Id Desc"
End If

Donate_Purpose_Type=""
Ary_Purpose_Type=Split(request("Donate_Purpose_Type"),",")
For I = 0 To UBound(Ary_Purpose_Type)
  If Donate_Purpose_Type="" Then
     Donate_Purpose_Type="'"&Trim(Ary_Purpose_Type(I))&"'"
  Else
    Donate_Purpose_Type=Donate_Purpose_Type&",'"&Trim(Ary_Purpose_Type(I))&"'"
  End If
Next

SQL1="Select Donate_Id,捐款用途=Donate_Purpose,捐款日期=Donate_Date,捐款人=(Case When IsAnonymous='Y' Then Case When NickName<>'' Then NickName Else Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End End Else Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End End),捐款金額=Donate_Amt,捐物內容=CONVERT(nvarchar(4000),Donate_Desc),編號抬頭=CONVERT(nvarchar(10),Donate.Donor_Id)+Donate.Invoice_Title " & _
     "From DONATE Join Donor On Donor.Donor_Id=Donate.Donor_Id " & _
     "Join CASECODE On Donate.Donate_Purpose=CASECODE.CodeDesc Where CASECODE.CodeType='Purpose' "
If Donate_Purpose_Type<>"" Then SQL1=SQL1&"And Donate_Purpose_Type In ("&Donate_Purpose_Type&") "
If Request("Donate_Payment")="1" Then
  'SQL1=SQL1&"And Donate_Payment Not In ('"&Request("DonateDesc")&"') "&SQL_Where&OrderBy_Where&""
  SQL1=SQL1&SQL_Where&OrderBy_Where&""
Else
  'SQL1=SQL1&"And Donate_Payment In ('"&Request("DonateDesc")&"') "&SQL_Where&OrderBy_Where&""
  SQL1=SQL1&SQL_Where&OrderBy_Where&""
End If

Session("SQL1")=SQL1
Session("action")=request("action")
Session("Donate_Date_Begin")=request("Donate_Date_Begin")
Session("Donate_Date_End")=request("Donate_Date_End")
Session("Act_Id")=request("Act_Id")
Session("Donate_Payment")=request("Donate_Payment")
If request("action")="report" Then
  If request("Donate_Payment")="2" And request("List_Type2")="2" Then
    Response.Redirect "../include/movebar.asp?URL=../donation/verify_list_rpt_style1.asp"
  Else
    Response.Redirect "../include/movebar.asp?URL=../donation/verify_list_rpt_style"&Request("List_Type")&".asp"
  End If
ElseIf request("action")="export" Then
  Response.Redirect "verify_list_rpt_style"&Request("List_Type")&".asp"
End If  
%>