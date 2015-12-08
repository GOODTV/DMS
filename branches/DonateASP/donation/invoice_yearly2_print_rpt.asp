<%
If Request("action")="help" Then Response.Redirect "invoice_help.htm"

'搜尋條件
'20140214 Modify by GoodTV Tanya:收據資料勾選「不寄」不列印年度收據
Donate_Where="Where Year(Donate_Date)='"&request("Donate_Year")&"' And Donate_Payment<>'"&request("DonateDesc")&"' And Issue_Type<>'D' And Donate.Invoice_Type <>'不寄' "
If Request("Dept_Id")<>"" Then
  Donate_Where=Donate_Where&"And Donate.Dept_Id = '"&Request("Dept_Id")&"' "
Else
  Donate_Where=Donate_Where&"And Donate.Dept_Id In ("&Session("all_dept_type")&") "
End If
If Request("Donor_Id_Begin")<>"" Then Donate_Where=Donate_Where&"And Donate.Donor_Id>='"&Request("Donor_Id_Begin")&"' "
If Request("Donor_Id_End")<>"" Then Donate_Where=Donate_Where&"And Donate.Donor_Id<='"&Request("Donor_Id_End")&"' "
If Request("Donor_Name")<>"" Then Donate_Where=Donate_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "
'20140320 增加地址查詢：國內或國外或全部
If Request("Address")="1" Then
	Donate_Where=Donate_Where&"And IsAbroad = 'N' "
ElseIf Request("Address")="2" Then
	Donate_Where=Donate_Where&"And IsAbroad = 'Y' "
Else
	Donate_Where=Donate_Where&"And IsAbroad <> '' "
End If
'20140213 Add by GoodTV Tanya:增加「收據重印」勾選
If Request("action")="report" Then
  If Request("RePrint")<>"" Then
    Donate_Where=Donate_Where&"And Invoice_Print_Yearly='1' "
  Else
    Donate_Where=Donate_Where&"And (Invoice_Print_Yearly='0' Or Invoice_Print_Yearly Is Null) "
  End If
End If

'排序方式
OrderBy_Where=""
If Request("OrderBy_Type")="1" Then
  If Request("Print_Desc")="1" Then
    OrderBy_Where=OrderBy_Where&"Order By ZipCode,Address,Donate.Donor_Id,Donate.Invoice_Title,Donate.Dept_Id "
  Else
    OrderBy_Where=OrderBy_Where&"Order By Invoice_ZipCode,Invoice_Address,Donate.Donor_Id,Donate.Invoice_Title,Donate.Dept_Id "
  End If
ElseIf Request("OrderBy_Type")="2" Then
  OrderBy_Where=OrderBy_Where&"Order By Donate.Donor_Id "
End If

SQL1="Select Distinct Donate.Dept_Id,Donate.Donor_Id,Donor_Name=Donate.Invoice_Title,Title=(Case When Title<>'' Then Title Else '' End),ZipCode,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then DONOR.Invoice_ZipCode+D.mValue+E.mValue+Invoice_Address Else DONOR.Invoice_ZipCode+D.mValue+Invoice_Address End End),Invoice_IDNo,Donate_Total=Isnull(Sum(Donate_Amt),0) " & _
     "From Donate Join DONOR On Donate.Donor_Id=DONOR.Donor_Id " & _
     "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
     "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode " & _  
     " "&Donate_Where&" " & _
     "Group By Donate.Dept_Id,Donate.Donor_Id,Donate.Invoice_Title,(Case When Title<>'' Then Title Else '' End),ZipCode,(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Invoice_ZipCode,(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then DONOR.Invoice_ZipCode+D.mValue+E.mValue+Invoice_Address Else DONOR.Invoice_ZipCode+D.mValue+Invoice_Address End End),Invoice_IDNo " & _
     " "&OrderBy_Where&" "

Session("DeptId")=Request("Dept_Id")
Session("DonateDesc")=Request("DonateDesc")
Session("Rept_Licence")=Request("Rept_Licence")
Session("Print_Desc")=Request("Print_Desc")
Session("SQL1")=SQL1
Session("Invoice_Title_New")=Request("Invoice_Title_New")
Session("Donate_Where")=Donate_Where
If Request("action")="report" Or Request("action")="preview" Then
  Session("action")=Request("action")
  Session("Donate_Year")=Request("Donate_Year")
  Response.Redirect "../include/movebar.asp?URL=../donation/invoice_yearly2_print_style"&Request("Invoice_Prog")&".asp"
ElseIf Request("action")="address" Then
  Response.Redirect "../include/movebar.asp?URL=../donation/address_list_rpt_style"&request("Label")&".asp"
ElseIf Request("action")="post" Then
  Response.Redirect "../include/movebar.asp?URL=../donation/invoice_yearly_print_post.asp"  
End If
%>
