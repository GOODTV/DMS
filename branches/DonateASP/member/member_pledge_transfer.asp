<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Session("Transfer_AspName")=request("Transfer_AspName")
  Server.Transfer "creditcard/"&request("Transfer_AspName")
End If

If request("action")="input" Then
  Row=0                '授權筆數
  Danate_Total=0       '授權金額
  Dept_Id_Pre=""       '前一筆轉帳機構代碼
  Ary_Pledge_Id=Split(request("pledge_id"),",")
  For I = 0 To UBound(Ary_Pledge_Id)
    SQL1="Select * From PLEDGE Where Pledge_Id='"&Trim(Ary_Pledge_Id(I))&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    If Not RS1.EOF Then      
      '取捐款編號
      Invoice_Pre=""
      Invoice_No=""
      Act_id=""
      If RS1("Act_id")<>"" Then Act_id=Cstr(RS1("Act_id"))
      InvoiceNo=Get_Invoice_No2("1",Cstr(RS1("Dept_Id")),Cstr(request("Donate_Date")),Cstr(RS1("Invoice_Type")),Act_id)
      If InvoiceNo<>"" Then
        Invoice_Pre=Split(InvoiceNo,"/")(0)
        Invoice_No=Split(InvoiceNo,"/")(1)
      End If
      
      '新增捐款資料
      donate_id=""
      SQL2="Select * From DONATE Where Pledge_Id='"&RS1("Pledge_Id")&"' And Donate_Date='"&request("Donate_Date")&"' "
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,3
      If RS2.EOF Then 
        RS2.Addnew
        RS2("Donor_Id")=RS1("Donor_Id")
        RS2("Pledge_Id")=RS1("Pledge_Id")
        RS2("Donate_Date")=request("Donate_Date")
        RS2("Invoice_Pre")=Invoice_Pre
        RS2("Invoice_No")=Invoice_No
        RS2("Create_Date")=Date()
        RS2("Create_DateTime")=Now()
        If RS1("Create_User")<>"" Then
          RS2("Create_User")=RS1("Create_User")
        Else
          RS2("Create_User")=session("user_name")
        End If
        RS2("Create_IP")=Request.ServerVariables("REMOTE_HOST")
      Else
        donate_id=RS2("donate_id")
        RS2("LastUpdate_Date")=Date()
        RS2("LastUpdate_DateTime")=Now()
        RS2("LastUpdate_User")=session("user_name")
        RS2("LastUpdate_IP")=Request.ServerVariables("REMOTE_HOST")
      End If
      RS2("Donate_Payment")=RS1("Donate_Payment")
      RS2("Donate_Purpose")=RS1("Donate_Purpose")
      RS2("Donate_Type")=RS1("Donate_Type")
      If RS1("Donate_FirstAmt")>0 Then
        RS2("Donate_Amt")=RS1("Donate_FirstAmt")
        RS2("Donate_Accou")=RS1("Donate_FirstAmt")
      ElseIf RS1("Donate_Amt")>0 Then
        RS2("Donate_Amt")=RS1("Donate_Amt")
        RS2("Donate_Accou")=RS1("Donate_Amt")
      Else
        RS2("Donate_Amt")="0"
        RS2("Donate_Accou")="0"
      End If
      RS2("Donate_Fee")="0"
      RS2("Donate_Forign")=""
      'RS2("Donate_Desc")=""
      'RS2("IsBeductible")="N"
      'RS2("Donate_Amt2")="0"
      RS2("Card_Bank")=RS1("Card_Bank")
      RS2("Card_Type")=RS1("Card_Type")
      RS2("Account_No")=RS1("Account_No")
      RS2("Valid_Date")=RS1("Valid_Date")
      RS2("Card_Owner")=RS1("Card_Owner")
      RS2("Owner_IDNo")=RS1("Owner_IDNo")
      RS2("Relation")=RS1("Relation")
      RS2("Authorize")=RS1("Authorize")
      RS2("Check_No")=""
      RS2("Check_ExpireDate")=null
      RS2("Post_Name")=RS1("Post_Name")
      RS2("Post_IDNo")=RS1("Post_IDNo")
      RS2("Post_SavingsNo")=RS1("Post_SavingsNo")
      RS2("Post_AccountNo")=RS1("Post_AccountNo")
      RS2("Dept_Id")=RS1("Dept_Id")
      RS2("Invoice_Title")=RS1("Invoice_Title")  
      RS2("Invoice_Pre")=Invoice_Pre
      RS2("Invoice_No")=Invoice_No
      RS2("Invoice_Print")="0"
      RS2("Invoice_Print_Add")="0"
      RS2("Invoice_Print_Yearly")="0"
      RS2("Invoice_Print_Yearly_Add")="0"
      RS2("Request_Date")=null
      RS2("Accoun_Bank")=RS1("Accoun_Bank")
      RS2("Accoun_Date")=null
      RS2("Invoice_Type")=RS1("Invoice_Type")
      RS2("Accounting_Title")=RS1("Accounting_Title")
      If RS1("Act_id")<>"" Then
        RS2("Act_id")=RS1("Act_id")
      Else
        RS2("Act_id")=null
      End If
      RS2("Comment")=RS1("Comment")
      RS2("Issue_Type")=""
      RS2("Issue_Type_Keep")=""
      RS2("Export")="N"
      RS2.Update
      RS2.Close
      Set RS2=Nothing
      
      '取捐款PK
      If donate_id="" Then
        SQL2="Select @@IDENTITY As donate_id"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
        donate_id=RS2("donate_id")
        RS2.Close
        Set RS2=Nothing
      End If
  
      '修改捐款人捐款紀錄
      SQL="DECLARE @last_donatedate datetime " & _
          "DECLARE @begin_donatedate datetime " & _
          "DECLARE @donate_no numeric " & _
          "DECLARE @donate_total numeric " & _
          "Select Top 1 @last_donatedate=Donate_Date From DONATE Where Donor_Id='"&RS1("donor_id")&"' And Issue_Type<>'D' Order By Donate_Date Desc " & _
          "Select Top 1 @begin_donatedate=Donate_Date From DONATE Where Donor_Id='"&RS1("donor_id")&"' And Issue_Type<>'D' Order By Donate_Date " & _
          "Select @donate_no=Count(*) From DONATE Where Donor_Id='"&RS1("donor_id")&"' And Issue_Type<>'D' " & _
          "Select @donate_total=Isnull(Sum(Donate_Amt),0) From DONATE Where Donor_Id='"&RS1("donor_id")&"' And Issue_Type<>'D' " & _
          "Update DONOR Set Last_DonateDate=@last_donatedate,Begin_DonateDate=@begin_donatedate,Donate_No=@donate_no,Donate_Total=@donate_total Where Donor_Id='"&RS1("donor_id")&"'"
      Set RS=Conn.Execute(SQL)
      
      '機構最後轉帳日
      If Dept_Id_Pre<>RS1("Dept_Id") Then
        SQL2="Update DEPT Set LastTransfer_Date='"&request("Donate_Date")&"' Where Dept_Id='"&RS1("Dept_Id")&"'"
        Set RS2=Conn.Execute(SQL2)
        Dept_Id_Pre=RS1("Dept_Id")
      End If
      
      RS1("Donate_FirstAmt")="0"
      RS1("Transfer_Date")=request("Donate_Date")
      '下次轉帳日期
      If Cint(Year(RS1("Donate_ToDate")))<Cint(Year(request("Donate_Date"))) Or (Cint(Year(RS1("Donate_ToDate")))=Cint(Year(request("Donate_Date"))) And Cint(Month(RS1("Donate_ToDate")))<=Cint(Month(request("Donate_Date")))) Then
        RS1("Status")="停止"
      Else
        If RS1("Donate_Period")="單筆" Then
          RS1("Status")="停止"
        Else  
          If RS1("Donate_Period")="月繳" Then
            Next_DonateDate=DateAdd("M",1,Year(request("Donate_Date"))&"/"&Month(request("Donate_Date"))&"/"&request("Transfer_Date"))
          ElseIf RS1("Donate_Period")="隔月繳" Then
            Next_DonateDate=DateAdd("M",2,Year(request("Donate_Date"))&"/"&Month(request("Donate_Date"))&"/"&request("Transfer_Date"))
          ElseIf RS1("Donate_Period")="季繳" Then
            Next_DonateDate=DateAdd("M",3,Year(request("Donate_Date"))&"/"&Month(request("Donate_Date"))&"/"&request("Transfer_Date"))
          ElseIf RS1("Donate_Period")="半年繳" Then
            Next_DonateDate=DateAdd("M",6,Year(request("Donate_Date"))&"/"&Month(request("Donate_Date"))&"/"&request("Transfer_Date"))
          ElseIf RS1("Donate_Period")="年繳" Then
            Next_DonateDate=DateAdd("M",12,Year(request("Donate_Date"))&"/"&Month(request("Donate_Date"))&"/"&request("Transfer_Date"))
          End If
          
          If Cint(Month(Next_DonateDate))=1 And Cint(Day(Next_DonateDate))=1 Then Next_DonateDate=DateAdd("D",1,Next_DonateDate)
          If Cint(Month(Next_DonateDate))=10 And Cint(Day(Next_DonateDate))=10 Then Next_DonateDate=DateAdd("D",1,Next_DonateDate)
          Check_NextDate=False
          While Check_NextDate=False
            If WeekDay(Next_DonateDate)="1" Or WeekDay(Next_DonateDate)="7" Then
              If WeekDay(Next_DonateDate)="1" Then
                Next_DonateDate=DateAdd("D",1,Next_DonateDate)
              Else
                Next_DonateDate=DateAdd("D",2,Next_DonateDate)
              End If
            Else
              Check_NextDate=True
            End If
          Wend
          
          If CDate(Next_DonateDate)<=CDate(RS1("Donate_ToDate")) Then
            RS1("Next_DonateDate")=Next_DonateDate
          Else
            RS1("Status")="停止"
          End If
        End If
      End If
      Danate_Total=Cdbl(Danate_Total)+Cdbl(RS1("Donate_Amt"))
      RS1.Update
      Row=Row+1
      
      
      '確認收據編號無重覆
      If Invoice_No<>"" Then
        Invoice_Pre_Old=Invoice_Pre
        Invoice_No_Old=Invoice_No
        Check_InvoiceNo=False
        While Check_InvoiceNo=False
          Check_Contribute=False
          SQL_Check="Select * From CONTRIBUTE Where Invoice_Pre='"&Invoice_Pre_Old&"' And Invoice_No='"&Invoice_No_Old&"'"
          Call QuerySQL(SQL_Check,RS_Check)
          If RS_Check.EOF Then Check_Contribute=True
          RS_Check.Close
          Set RS_Check=Nothing

          Check_Donate=False
          SQL_Check="Select * From DONATE Where Invoice_Pre='"&Invoice_Pre_Old&"' And Invoice_No='"&Invoice_No_Old&"' And Donate_Id<>'"&donate_id&"' "
          Call QuerySQL(SQL_Check,RS_Check)
          If RS_Check.EOF Then Check_Donate=True
          RS_Check.Close
          Set RS_Check=Nothing

          If Check_Contribute And Check_Donate Then
            Check_InvoiceNo=True
          Else
            Act_id=""
            If RS1("Act_id")<>"" Then Act_id=Cstr(RS1("Act_id"))
            InvoiceNo=Get_Invoice_No2("1",Cstr(RS1("Dept_Id")),Cstr(request("Donate_Date")),Cstr(RS1("Invoice_Type")),Act_id)
            If InvoiceNo<>"" Then
              Invoice_Pre_Old=Split(InvoiceNo,"/")(0)
              Invoice_No_Old=Split(InvoiceNo,"/")(1)
              SQL="Update DONATE Set Invoice_Pre='"&Invoice_Pre_Old&"',Invoice_No='"&Invoice_No_Old&"' Where Donate_Id='"&donate_id&"' "
              Set RS=Conn.Execute(SQL)
            End If
          End If
        Wend
      End If

      Response.Flush
      Response.Clear
    End If
    RS1.Close
    Set RS1=Nothing
  Next
  SQL2="Update PLEDGE Set Status='停止' Where Status='授權中' And ((Year(Donate_ToDate)<'"&Year(request("Donate_Date"))&"') Or (Year(Donate_ToDate)='"&Year(request("Donate_Date"))&"' And Month(Donate_ToDate)<'"&Month(request("Donate_Date"))&"')) "
  Set RS2=Conn.Execute(SQL2)
  
  session("errnumber")=1
  session("msg")="授權轉入捐款成功 ！\n轉入筆數："&FormatNumber(Row,0)&"\n轉入金額："&FormatNumber(Danate_Total,0)&""
End If


SQL1="Select * From DEPT Where Dept_Id='"&session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
Transfer_Date=RS1("Transfer_Date")
LastTransfer_Date=RS1("LastTransfer_Date")
RS1.Close
Set RS1=Nothing
If request("Donate_Date")="" Then
  If LastTransfer_Date<>"" Then
    DonateDate=DateAdd("M",1,LastTransfer_Date)  
  Else
    DonateDate=Year(Date())&"/"&Month(Date())&"/"&Transfer_Date
  End If
Else
  DonateDate=request("Donate_Date")
End If

LastTransfer_Check="0"
If LastTransfer_Date<>"" Then
  If Cstr(Year(LastTransfer_Date))=Cstr(Year(Date())) And Cstr(Month(LastTransfer_Date))=Cstr(Month(Date())) Then LastTransfer_Check="1"
End If
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>每月固定轉帳作業</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="Transfer_Date" value="<%=Transfer_Date%>">	
      <input type="hidden" name="LastTransfer_Check" value="<%=LastTransfer_Check%>">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="780" border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"> </td>
                <td width="95%">
  		            <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                <tr>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                      <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                    </tr>
                    <tr>
                      <td class="table62-bg">　</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">每月固定轉帳作業</td>
                      <td class="table63-bg">　</td>
                    </tr>  
    	            </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td> 
	          <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">　</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="1" cellspacing="1">
                    <tr>
                      <td class="td02-c" width="65" align="right">機構：</td>
                      <td class="td02-c" width="95">
                      <%
                        If Session("comp_label")="1" Then
                          SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' And Dept_Id In ("&Session("all_dept_type")&") Order By Comp_Label,Dept_Id"
                        Else
                          SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id In ("&Session("all_dept_type")&") Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                        End If
                        FName="Dept_Id"
                        Listfield="Comp_ShortName"
                        menusize="1"
                        BoundColumn=request("Dept_Id")
                        call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                      %>
                      </td>
                      <td class="td02-c" width="65" align="right">捐款人：</td>
                      <td class="td02-c" width="70"><input type="text" name="Donor_Name" size="8" class="font9" value="<%=request("Donor_Name")%>"></td>
                      <td class="td02-c" width="65" align="right">TXT格式：</td>
                      <td class="td02-c" width="110">
                      <%
                        SQL="Select Transfer_AspName,Transfer_BankName From DONATE_TRANSFER Where Transfer_StopFlg='N' Order By Transfer_Seq,Ser_No Desc"
                        FName="Transfer_AspName"
                        Listfield="Transfer_BankName"
                        menusize="1"
                        BoundColumn=request("Transfer_AspName")
                        call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                      %>
                      </td>
                      <td class="td02-c" width="240" align="right">
                        <input type="button" value="1.TXT匯出" name="export" class="addbutton" style="cursor:hand" onClick="Export_OnClick()">
                        <input type="button" value="2.授權轉捐款" name="input" class="addbutton" style="cursor:hand" onClick="Input_OnClick()">	
                        <input type="button" value="查詢" name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">扣款日期：</td>
                      <td class="td02-c"><%call Calendar("Donate_Date",DonateDate)%></td>
                      <td class="td02-c" align="right">轉帳週期：</td>
                      <td class="td02-c">
                        <SELECT Name='Donate_Period' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      	  <OPTION  value=''> </OPTION>
                      		<OPTION  value='單筆' <%If request("Donate_Period")="單筆" Then Response.Write "selected" End If%> >單筆</OPTION>
                      		<OPTION  value='月繳' <%If request("Donate_Period")="月繳" Then Response.Write "selected" End If%> >月繳</OPTION>
                      		<OPTION  value='季繳' <%If request("Donate_Period")="季繳" Then Response.Write "selected" End If%> >季繳</OPTION>
                      		<OPTION  value='半年繳' <%If request("Donate_Period")="半年繳" Then Response.Write "selected" End If%> >半年繳</OPTION>
                      		<OPTION  value='年繳' <%If request("Donate_Period")="年繳" Then Response.Write "selected" End If%> >年繳</OPTION>
                      	</SELECT>
                      </td>
                      <td class="td02-c" align="right">授權方式：</td>
                      <td class="td02-c">
                      <%
                        SQL="Select Donate_Payment=CodeDesc From CASECODE Where CodeType='Payment' And CodeDesc Like '%授權書%' Order By Seq"
                        FName="Donate_Payment"
                        Listfield="Donate_Payment"
                        menusize="1"
                        BoundColumn=request("Donate_Payment")
                        call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                      %>
                      </td>
                      <td class="td02-c"> </td>	
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <%
                      SQL_P="Select Donate_No=Count(*),Donate_Amt=Sum(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End) From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And Post_SavingsNo<>'' And Post_AccountNo<>'' "
                      SQL_C="Select Donate_No=Count(*),Donate_Amt=Sum(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End) From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And Account_No<>'' And (Convert(int,RIGHT(Valid_Date,2))>"&Right(Year(DonateDate),2)&" Or (Convert(int,RIGHT(Valid_Date,2))="&Right(Year(DonateDate),2)&" And Convert(int,Left(Valid_Date,2))>="&Month(DonateDate)&")) "
                      SQL1="Select Pledge_Id,授權編號=Pledge_Id,捐款人=DONOR.Donor_Name,授權方式=Donate_Payment,扣款金額=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),授權起日=CONVERT(VarChar,Donate_FromDate,111),授權迄日=CONVERT(VarChar,Donate_ToDate,111),轉帳週期=Donate_Period,下次扣款日=Next_DonateDate,有效月年=(Case When Valid_Date<>'' Then Left(Valid_Date,2)+'/'+Right(Valid_Date,2) Else '' End),最後扣款日=Transfer_Date " & _
                           "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And ((Post_SavingsNo<>'' And Post_AccountNo<>'') Or Account_No<>'') "
                      If request("Dept_Id")<>"" Then 
                        WhereSQL=WhereSQL & "And PLEDGE.Dept_Id = '"&request("Dept_Id")&"' "
                      Else
                        If request("action")="" Then WhereSQL=WhereSQL & "And PLEDGE.Dept_Id In ("&Session("all_dept_type")&") "
                      End If
                      If request("Donor_Name")<>"" Then WhereSQL=WhereSQL & "And (DONOR.Donor_Name Like '%"&request("Donor_Name")&"%' Or DONOR.NickName Like '%"&request("Donor_Name")&"%' Or DONOR.Contactor Like '%"&request("Donor_Name")&"%' Or DONOR.Invoice_Title Like '%"&request("Donor_Name")&"%') "
                      If request("Donate_Date")<>"" Then 
                        WhereSQL=WhereSQL & "And Donate_FromDate <= '"&request("Donate_Date")&"' And Donate_ToDate>= '"&request("Donate_Date")&"' And ((Year(Next_DonateDate)<'"&Year(request("Donate_Date"))&"') Or (Year(Next_DonateDate)='"&Year(request("Donate_Date"))&"' And Month(Next_DonateDate)<='"&Month(request("Donate_Date"))&"')) "
                        WhereSQL_PC=WhereSQL_PC & "And Donate_FromDate <= '"&request("Donate_Date")&"' And Donate_ToDate>= '"&request("Donate_Date")&"' And ((Year(Next_DonateDate)<'"&Year(request("Donate_Date"))&"') Or (Year(Next_DonateDate)='"&Year(request("Donate_Date"))&"' And Month(Next_DonateDate)<='"&Month(request("Donate_Date"))&"')) "
                      Else
                        WhereSQL=WhereSQL & "And Donate_FromDate <= '"&DonateDate&"' And Donate_ToDate>= '"&DonateDate&"' And ((Year(Next_DonateDate)<'"&Year(DonateDate)&"') Or (Year(Next_DonateDate)='"&Year(DonateDate)&"' And Month(Next_DonateDate)<='"&Month(DonateDate)&"')) "
                        WhereSQL_PC=WhereSQL_PC & "And Donate_FromDate <= '"&DonateDate&"' And Donate_ToDate>= '"&DonateDate&"' And ((Year(Next_DonateDate)<'"&Year(DonateDate)&"') Or (Year(Next_DonateDate)='"&Year(DonateDate)&"' And Month(Next_DonateDate)<='"&Month(DonateDate)&"')) "
                      End If
                      If request("Donate_Payment")<>"" Then WhereSQL=WhereSQL & "And Donate_Payment = '"&request("Donate_Payment")&"' "
                      If request("Donate_Period")<>"" Then WhereSQL=WhereSQL & "And Donate_Period = '"&request("Donate_Period")&"' "
                      
                      '帳戶轉帳授權筆數/金額
                      Donate_No_P=0
                      Donate_Amt_P=0
                      SQL_P=SQL_P&WhereSQL&WhereSQL_PC
                      Set RS_P = Server.CreateObject("ADODB.RecordSet")
                      RS_P.Open SQL_P,Conn,1,1
                      If Not RS_P.EOF Then
                        If RS_P("Donate_No")<>"" Then Donate_No_P=Cdbl(RS_P("Donate_No"))
                        If RS_P("Donate_Amt")<>"" Then Donate_Amt_P=Cdbl(RS_P("Donate_Amt"))
                      End If
                      RS_P.Close
                      Set RS_P=Nothing
                      
                      '信用卡扣款授權筆數/金額
                      Donate_No_C=0
                      Donate_Amt_C=0  
                      SQL_C=SQL_C&WhereSQL&WhereSQL_PC
                      Set RS_C = Server.CreateObject("ADODB.RecordSet")
                      RS_C.Open SQL_C,Conn,1,1
                      If Not RS_C.EOF Then
                        If RS_C("Donate_No")<>"" Then Donate_No_C=Cdbl(RS_C("Donate_No"))
                        If RS_C("Donate_Amt")<>"" Then Donate_Amt_C=Cdbl(RS_C("Donate_Amt"))
                      End If
                      RS_C.Close
                      Set RS_C=Nothing
                    %>
                    <tr>
                      <td class="td02-c" colspan="7">
                         <table width="100%" border="0" cellpadding="3" cellspacing="3">
                           <tr>
                    	       <td width="100">信用卡有效月年：</td>
                    	       <td width="15" height="15" bgcolor='#66FF99'> </td>
                    	       <td width="30">2個月</td>
                    	       <td width="15" height="15" bgcolor='#FFFF99'> </td>
                    	       <td width="30">1個月</td>
                    	       <td width="15" height="15" bgcolor='#FFCCFF'> </td>
                    	       <td width="25">本月</td>
                    	       <td width="15" height="15" bgcolor='#FF6666'> </td>
                    	       <td width="25">逾期</td>
                    	       <td align="right">本月應執行&nbsp;帳戶轉帳<%=FormatNumber(Donate_No_P,0)%>筆&nbsp;&nbsp;金額<%=FormatNumber(Donate_Amt_P,0)%>元&nbsp;╱&nbsp;信用卡扣款<%=FormatNumber(Donate_No_C,0)%>筆&nbsp;&nbsp;金額<%=FormatNumber(Donate_Amt_C,0)%>&nbsp;元</td>
                           </tr>
                         </table>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" colspan="7" width="100%" height="400" valign="top">
                      <%					
                        SQL1=SQL1&WhereSQL&"Order By Next_DonateDate,Pledge_Id Desc "
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        FieldsCount = RS1.Fields.Count-1
                        Dim I, J
                        Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;background-color: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
                        Response.Write "<tr>"
                        If RS1.Recordcount>0 Then
                          Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'><input type='checkbox' name='pledge_id' id='pledge_id_0' value='0' checked OnClick='PledgeId_OnClick()'></span></font></td>"
                        Else
                          Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>選擇</span></font></td>"
                        End If
                        For I = 1 To FieldsCount
	                        Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
                        Next
                        Response.Write "</tr>"
                        Row=0
                        While Not RS1.EOF
                          Row=Row+1
                          StrChecked="checked"
                          bgColor="#FFFFFF"
                          If RS1(9)<>"" Then                            
                            Valid_Year_N=Right(Year(Date()),2)
                            Valid_Month_N=Month(Date())
                            If Len(Valid_Month_N)=1 Then Valid_Month_N="0"&Valid_Month_N
                            Valid_Date_N=Cint(Cstr(Valid_Year_N&Valid_Month_N))
                                                        
                            Valid_Year_2=Right(Year(DateAdd("M",2,"20"&Valid_Year_N&"/"&Valid_Month_N&"/1")),2)
                            Valid_Month_2=Month(DateAdd("M",2,"20"&Valid_Year_N&"/"&Valid_Month_N&"/1"))
                            If Len(Valid_Month_2)=1 Then Valid_Month_2="0"&Valid_Month_2
                            Valid_Date_2=Cint(Cstr(Valid_Year_2&Valid_Month_2))
                            
                            Valid_Year_1=Right(Year(DateAdd("M",1,"20"&Valid_Year_N&"/"&Valid_Month_N&"/1")),2)
                            Valid_Month_1=Month(DateAdd("M",1,"20"&Valid_Year_N&"/"&Valid_Month_N&"/1"))
                            If Len(Valid_Month_1)=1 Then Valid_Month_1="0"&Valid_Month_1
                            Valid_Date_1=Cint(Cstr(Valid_Year_1&Valid_Month_1))
                            
                            Valid_Year=Right(RS1(9),2)
                            Valid_Month=Left(RS1(9),2)
                            Valid_Date=Cint(Cstr(Valid_Year&Valid_Month))
                            
                            If Valid_Date=Valid_Date_2 Then
                              bgColor="#66FF99"
                            ElseIf Valid_Date=Valid_Date_1 Then
                              bgColor="#FFFF99"
                            ElseIf Valid_Date=Valid_Date_N Then
                              bgColor="#FFCCFF"
                            ElseIf Valid_Date<Valid_Date_N Then
                              StrChecked=""
                              bgColor="#FF6666"
                            End If
                          End If
                          
                          'If request("Donate_Date")<>"" Then
                          '  If Cint(Year(RS1(8)))<>Cint(Year(request("Donate_Date"))) Or Cint(Month(RS1(8)))<>Cint(Month(request("Donate_Date"))) Then StrChecked=""
                          'ElseIf DonateDate<>"" Then
                          '  If Cint(Year(RS1(8)))<>Cint(Year(DonateDate)) Or Cint(Month(RS1(8)))<>Cint(Month(DonateDate)) Then StrChecked=""
                          'End If
                          
                          Response.Write "<tr "&showhand&" bgcolor='"&bgColor&"' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor="""&bgColor&"""'>"
                          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'><input type='checkbox' name='pledge_id' id='pledge_id_"&Row&"' value='"&RS1(0)&"' "&StrChecked&" >" & "</span></td>"
                          For J = 1 To FieldsCount
                            If J=4 Then
                              Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(J),0) & "</span></td>"
                            Else
                              Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
                            End If
                          Next
                          Response.Write "</tr>"
                          Response.Flush
                          Response.Clear
                          RS1.MoveNext
                        Wend  
                        Response.Write "</table>"
                        RS1.Close
                        Set RS1=Nothing
                      %>
                      <input type="hidden" name="Total_Row" value="<%=Row%>">		
                      </td>
                    </tr>
                  </table>
                </td>
                <td class="table63-bg">　</td>
              </tr>
              <tr>
                <td style="background-color:#EEEEE3"><img src="../images/table06_06.gif" width="10" height="10"></td>
                <td class="table64-bg"><img src="../images/table06_07.gif" width="1" height="10"></td>
                <td style="background-color:#EEEEE3"><img src="../images/table06_08.gif" width="10" height="10"></td>
              </tr>
            </table>
           </td>
          </tr>
      </table>
    </form>
  </center></div></p>
 <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Export_OnClick(){
  <%call CheckStringJ("Donate_Date","扣款日期")%>
  <%call CheckDateJ("Donate_Date","扣款日期")%>
  <%call CheckStringJ("Transfer_AspName","TXT格式")%>
  if(confirm('您是否確定要將TXT授權資料匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
function Input_OnClick(){
  <%call CheckStringJ("Donate_Date","扣款日期")%>
  <%call CheckDateJ("Donate_Date","扣款日期")%>
  if(document.form.Total_Row.value>'0'){
    if(document.form.LastTransfer_Check.value=='1'){
      if(confirm('本月份已執行過授權資料轉入捐款資料，\n\n您是否確定要『再次』執行？')){
        document.form.action.value='input';
        document.form.submit();
      }  
    }else{
       if(confirm('您是否確定要將授權資料轉入捐款資料？\n\n※請注意轉入過程中請勿『關閉視窗』')){
        document.form.action.value='input';
        document.form.submit();
      }  
    }
  }else{
    alert('查無相關授權資料無法轉入！');
    return;
  }  
}
function Query_OnClick(){
  document.form.action.value="query"
  document.form.submit();
}	
function PledgeId_OnClick(){
  if(document.form.pledge_id[0].checked){
    for(var i=1;i<=Number(document.form.Total_Row.value);i++){
      document.form.pledge_id[i].checked=true;
    }
  }else{
    for(var i=1;i<=Number(document.form.Total_Row.value);i++){
      document.form.pledge_id[i].checked=false;
    }
  }
}
--></script>	