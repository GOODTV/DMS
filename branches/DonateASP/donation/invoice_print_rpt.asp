<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body>
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
    For I = 0 To UBound(Ary_Invoice_Type)
      If I=0 Then
        invoicetype=invoicetype&"('"&Trim(Ary_Invoice_Type(I))&"',"
      ElseIf I=UBound(Ary_Invoice_Type) Then
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
'If Request("Donor_Name")<>"" Then Donor_Where=Donor_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "
If Request("Donor_Id")<>"" Then Donor_Where=Donor_Where&"And DONATE.Donor_Id = '"&request("Donor_Id")&"' "
If Request("Print_Address")="1" Then
	Donor_Where=Donor_Where&"And IsAbroad = 'N' And IsAbroad_Invoice = 'N' "
ElseIf Request("Print_Address")="2" Then
	Donor_Where=Donor_Where&"And IsAbroad = 'Y' And IsAbroad_Invoice = 'Y' "
End If
'20140122 Add by GoodTV Tanya:開放身份別「主知名」可選擇是否列印
If Request("Print_NoName")="Y" Then Donor_Where=Donor_Where&"And IsNull(DONOR.Donor_Type,'') Not Like '%主知名%' "

'排序方式
OrderBy_Where=""
If Request("OrderBy_Type")="1" Then
  OrderBy_Where=OrderBy_Where&"Order By Donate.Dept_Id,Invoice_No"
ElseIf Request("OrderBy_Type")="2" Then
  OrderBy_Where=OrderBy_Where&"Order By Donate.Donor_Id,Invoice_No"
ElseIf Request("OrderBy_Type")="3" Then
  If Request("Print_Desc")="1" Then
  	'20140218 Modify by GoodTV Tanya:因新資料庫Table CODECITY也有ZipCode欄位，因此SQL加上指定Table DONOR.ZipCode
    OrderBy_Where=OrderBy_Where&"Order By DONOR.ZipCode,Donate.Dept_Id,Invoice_No"
  Else
    OrderBy_Where=OrderBy_Where&"Order By Invoice_ZipCode,Donate.Dept_Id,Invoice_No"
  End If
End If

'搜尋單筆收據資料
'20131001 Modify by GoodTV Tanya:查詢排除「錯址」、「不主動聯絡」、「歿」及「主知名」
'20131213 Modify by GoodTV Tanya:收據地址加上「Attn」
'20140107 Modify by GoodTV Tanya:地址City是否為空時的判斷加上IsNull
'20140122 Modify by GoodTV Tanya:開放身份別「主知名」可選擇是否列印
'20140212 Modify by GoodTV Tanya:收據地址(鄉鎮市區)可能為空值查詢
'20140218 Modify by GoodTV Tanya:因新資料庫Table CODECITY也有ZipCode欄位，因此SQL加上指定Table DONOR.ZipCode
SQL1="Select Donor.Attn,Donor.Invoice_Attn,Donate_Id,Invoice_No=(Case When Donate.Invoice_No_Old<>'' Then Donate.Invoice_No_Old Else Invoice_Pre+Invoice_No End), " & _
     "Donate_Date=CONVERT(VarChar,Donate_Date,111),Donor_Name=(Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End),Donor_Name1=(Case When Donor_Name<>'' Then Donor_Name Else Donate.Invoice_Title End), " & _
     "Title,Title2,DONOR.ZipCode,Address=(Case When IsNull(DONOR.City,'')='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+IsNull(C.mValue,'')+Address Else B.mValue+Address End End), " & _
     "Address2=(Case When IsNull(DONOR.City,'')='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+DONOR.ZipCode+IsNull(C.mValue,'')+Address Else B.mValue+DONOR.ZipCode+Address End End), " & _
     "Invoice_ZipCode,Invoice_Address=(Case When IsAbroad_Invoice='Y' Then Invoice_OverseasAddress Else Case When D.mValue<>E.mValue Then D.mValue+IsNull(E.mValue,'')+Invoice_Address Else D.mValue+Invoice_Address End End), " & _ 
     "Invoice_Address2=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then DONOR.Invoice_ZipCode+D.mValue+IsNull(E.mValue,'')+Invoice_Address Else DONOR.Invoice_ZipCode+D.mValue+Invoice_Address End End), " & _
     "Invoice_OverseasCountry,IsAbroad,IsAbroad_Invoice,Donate_Amt,Donate_Amt2,Donate_Desc=CONVERT(nvarchar(4000),Donate_Desc),Donate.Dept_Id,Invoice_IDNo,DONATE.Donor_Id,Donate_Purpose,Donate_Payment,Invoice_Print,Invoice_Print_Add, " & _
     "Post_Donor_Name=Donor_Name,TEL=(Case When DONOR.Cellular_Phone<>'' Then DONOR.Cellular_Phone Else Case When DONOR.Tel_Office<>'' Then DONOR.Tel_Office Else DONOR.Tel_Home End End), " & _
     "Invoice_PrintComment=CONVERT(nvarchar(4000),Invoice_PrintComment),Comment=CONVERT(nvarchar(4000),Comment),Invoice_No_Old,Donate.Invoice_Type,Donate.Create_User,Donate_Forign " & _
     "From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id " & _
     "Left Join ACT On DONATE.Act_Id=ACT.Act_Id " & _
     "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
     "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode Where Donate_Purpose_Type In ('D') And Invoice_No<>'' " & _
     " And DONOR.Donor_Name <> '主知名' And IsNull(DONATE.Invoice_Title,'') <> '主知名' " & _
     " And IsNull(DONOR.IsErrAddress,'') <> 'Y' And IsNull(DONOR.IsContact,'') <> 'N' And IsNull(DONOR.Sex,'') <> '歿' "
If Request("action")="address" Then
  If Request("Print_Desc")="1" Then
    SQL1=SQL1&"And Invoice_Address<>'' "         
  Else
    SQL1=SQL1&"And Address<>'' "       
  End If
End If
SQL1=SQL1&" "&Donate_Where&Donor_Where&OrderBy_Where&" "
'response.write SQL1
'response.end()
Session("Rept_Licence")=Request("Rept_Licence")
Session("Print_Desc")=Request("Print_Desc")
Session("SQL1")=SQL1

'If Request("action")="report" Then
'  If Request("Print_Pint")="1" Then
    '捐款收據
    'Response.Redirect "../include/movebar.asp?URL=../donation/invoice_print_style"&Request("Invoice_Prog")&".asp"
	'Response.Redirect server.URLEncode("http://donategoodtv.sino1.com.tw/report/printreceipt.aspx?sql="&SQL1)
'	URL="http://donategoodtv.sino1.com.tw/report/printreceipt.aspx?sql="&SQL1
'	server.Transfer(URL)
'  ElseIf Request("Print_Pint")="2" Then
'    '捐物收據
'    Response.Redirect "../include/movebar.asp?URL=../donation/invoice_print_style"&Request("Invoice_Prog")&".asp"
'  End If
'ElseIf Request("action")="address" Then
'  Response.Redirect "../include/movebar.asp?URL=../donation/address_list_rpt_style"&request("Label")&".asp"
'ElseIf Request("action")="post" Then
'  Session("DeptId")=Request("Dept_Id")
'  Response.Redirect "../include/movebar.asp?URL=../donation/invoice_print_post.asp"  
'End If

'20131119 Add by GoodTV Tanya:增加更新收據列印紀錄
SQL2=""
If Request("action")="report" Then
	UpdateItem=""
  If Request("RePrint")<>"" Then '重印收據  	
   UpdateItem = "Invoice_RePrint_Date=GETDATE()"
  Else
   UpdateItem = "Invoice_Print='1',Invoice_Print_Date=GETDATE()"
  End If  	
'20140212 Modify by GoodTV Tanya:修正因身份別為主知名，列印紀錄無法更新問題
	SQL2="Update DONATE SET "&UpdateItem&" Where Donate_Id IN " & _	      
			 "(Select Donate_Id " & _
	     " From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id " & _
       " Left Join ACT On DONATE.Act_Id=ACT.Act_Id " & _
       " Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
       " Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode " & _ 
       " Where Donate_Purpose_Type In ('D') And Invoice_No<>'' " & _
       " 	 And DONOR.Donor_Name <> '主知名' And IsNull(DONATE.Invoice_Title,'') <> '主知名' " & _
       " 	 And IsNull(DONOR.IsErrAddress,'') <> 'Y' And IsNull(DONOR.IsContact,'') <> 'N' And IsNull(DONOR.Sex,'') <> '歿' " &Donate_Where&Donor_Where&")"  
'	response.write SQL2
End If  	

%>
<form name="form" action="http://donate3.goodtv.tv/report/printreceipt.aspx" method="post" >
<input type="hidden" name="SQL" value="<%=Server.URLEncode(SQL1)%>" />
<input type="hidden" name="UpdateSQL" value="<%=Server.URLEncode(SQL2)%>" />
</form>
<script type="text/javascript">
	document.form.submit();		
</script>
</body>
</html>