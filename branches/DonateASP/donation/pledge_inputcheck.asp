<% Response.Buffer = True %>
<!--#include file="../include/dbfunctionJ.asp"-->
<%	
pledge_id=Session("pledge_id")
Donate_Date=Session("Donate_Date")
Transfer_Date=Session("Transfer_Date")
If request("act")="input" Then
  Row=0                '授權筆數
  Danate_Total=0       '授權金額
  Dept_Id_Pre=""       '前一筆轉帳機構代碼
  Ary_Pledge_Id=Split(Session("pledge_id"),",")
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
      'InvoiceNo=Get_Invoice_No2("1",Cstr(RS1("Dept_Id")),Cstr(Donate_Date),Cstr(RS1("Invoice_Type")),Act_id)
      InvoiceNo=Get_Invoice_No(Cstr(RS1("Dept_Id")),Cstr(Donate_Date),Act_id)
      If InvoiceNo<>"" Then
        Invoice_Pre=Split(InvoiceNo,"/")(0)
        Invoice_No=Split(InvoiceNo,"/")(1)
      End If

      CodeKind="D"
      SQL2="Select CodeKind From CASECODE Where CodeType='Purpose' And CodeDesc='"&RS1("Donate_Purpose")&"' "
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,1
      CodeKind=RS2("CodeKind")
      RS2.Close
      Set RS2=Nothing
      
      '新增捐款資料
      donate_id=""
      SQL2="Select * From DONATE Where Pledge_Id='"&RS1("Pledge_Id")&"' And Donate_Date='"&Donate_Date&"' "
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,3
      If RS2.EOF Then 
        RS2.Addnew
        RS2("Donor_Id")=RS1("Donor_Id")
        RS2("Pledge_Id")=RS1("Pledge_Id")
        RS2("Donate_Date")=Donate_Date
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
      RS2("Donate_Purpose_Type")=CodeKind
      RS2("Donate_Type")=RS1("Donate_Type")  
      If RS1("Donate_FirstAmt")>0 Then
        RS2("Donate_Amt")=RS1("Donate_FirstAmt")
        RS2("Donate_Accou")=RS1("Donate_FirstAmt")
        RS2("Donate_Fee")="0"
        RS2("Donate_Amt"&CodeKind&"")=RS1("Donate_FirstAmt")
        RS2("Donate_Accou"&CodeKind&"")=RS1("Donate_FirstAmt")
        RS2("Donate_Fee"&CodeKind&"")="0"
      ElseIf RS1("Donate_Amt")>0 Then
        RS2("Donate_Amt")=RS1("Donate_Amt")
        RS2("Donate_Accou")=RS1("Donate_Amt")
        RS2("Donate_Fee")="0"
        RS2("Donate_Amt"&CodeKind&"")=RS1("Donate_Amt")
        RS2("Donate_Accou"&CodeKind&"")=RS1("Donate_Amt")
        RS2("Donate_Fee"&CodeKind&"")="0"
      Else
        RS2("Donate_Amt")="0"
        RS2("Donate_Accou")="0"
        RS2("Donate_Fee")="0"
        RS2("Donate_Amt"&CodeKind&"")="0"
        RS2("Donate_Accou"&CodeKind&"")="0"
        RS2("Donate_Fee"&CodeKind&"")="0"
      End If
      If CodeKind="D" Then
        RS2("Donate_AmtM")="0"
        RS2("Donate_FeeM")="0"
        RS2("Donate_AccouM")="0"
        RS2("Donate_AmtA")="0"
        RS2("Donate_FeeA")="0"
        RS2("Donate_AccouA")="0"
        RS2("Donate_AmtS")="0"
        RS2("Donate_RateS")="0"
        RS2("Donate_FeeS")="0"
        RS2("Donate_AccouS")="0"
      ElseIf CodeKind="M" Then
        RS2("Donate_AmtD")="0"
        RS2("Donate_FeeD")="0"
        RS2("Donate_AccouD")="0"
        RS2("Donate_AmtA")="0"
        RS2("Donate_FeeA")="0"
        RS2("Donate_AccouA")="0"
        RS2("Donate_AmtS")="0"
        RS2("Donate_RateS")="0"
        RS2("Donate_FeeS")="0"
        RS2("Donate_AccouS")="0"
      ElseIf CodeKind="A" Then
        RS2("Donate_AmtD")="0"
        RS2("Donate_FeeD")="0"
        RS2("Donate_AccouD")="0"
        RS2("Donate_AmtM")="0"
        RS2("Donate_FeeM")="0"
        RS2("Donate_AccouM")="0"
        RS2("Donate_AmtS")="0"
        RS2("Donate_RateS")="0"
        RS2("Donate_FeeS")="0"
        RS2("Donate_AccouS")="0"
      ElseIf CodeKind="S" Then
        RS2("Donate_AmtD")="0"
        RS2("Donate_FeeD")="0"
        RS2("Donate_AccouD")="0"
        RS2("Donate_AmtM")="0"
        RS2("Donate_FeeM")="0"
        RS2("Donate_AccouM")="0"
        RS2("Donate_AmtA")="0"
        RS2("Donate_FeeA")="0"
        RS2("Donate_AccouA")="0"
        RS2("Donate_RateS")="0"
      End If
      RS2("Donate_Forign")=""
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
        SQL2="Update DEPT Set LastTransfer_Date='"&Donate_Date&"' Where Dept_Id='"&RS1("Dept_Id")&"'"
        Set RS2=Conn.Execute(SQL2)
        Dept_Id_Pre=RS1("Dept_Id")
      End If
      
      RS1("Donate_FirstAmt")="0"
      RS1("Transfer_Date")=Donate_Date
      '下次轉帳日期
      If Cint(Year(RS1("Donate_ToDate")))<Cint(Year(Donate_Date)) Or (Cint(Year(RS1("Donate_ToDate")))=Cint(Year(Donate_Date)) And Cint(Month(RS1("Donate_ToDate")))<=Cint(Month(Donate_Date))) Then
        RS1("Status")="停止"
      Else
        If RS1("Donate_Period")="單筆" Then
          RS1("Status")="停止"
        Else  
          If RS1("Donate_Period")="月繳" Then
            Next_DonateDate=DateAdd("M",1,Year(Donate_Date)&"/"&Month(Donate_Date)&"/"&Transfer_Date)
          ElseIf RS1("Donate_Period")="隔月繳" Then
            Next_DonateDate=DateAdd("M",2,Year(Donate_Date)&"/"&Month(Donate_Date)&"/"&Transfer_Date)
          ElseIf RS1("Donate_Period")="季繳" Then
            Next_DonateDate=DateAdd("M",3,Year(Donate_Date)&"/"&Month(Donate_Date)&"/"&Transfer_Date)
          ElseIf RS1("Donate_Period")="半年繳" Then
            Next_DonateDate=DateAdd("M",6,Year(Donate_Date)&"/"&Month(Donate_Date)&"/"&Transfer_Date)
          ElseIf RS1("Donate_Period")="年繳" Then
            Next_DonateDate=DateAdd("M",12,Year(Donate_Date)&"/"&Month(Donate_Date)&"/"&Transfer_Date)
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
            'InvoiceNo=Get_Invoice_No2("1",Cstr(RS1("Dept_Id")),Cstr(Donate_Date),Cstr(RS1("Invoice_Type")),Act_id)
            InvoiceNo=Get_Invoice_No(Cstr(RS1("Dept_Id")),Cstr(Donate_Date),Act_id)
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
	  '將已成功授權轉捐款的資料從Table dbo.PLEDGE_SEND_RETURN刪除 2014/1/20 update
	  SQL="delete from dbo.PLEDGE_SEND_RETURN where Pledge_Id = '"&Trim(Ary_Pledge_Id(I))&"'"
	  Set RS=Conn.Execute(SQL)
    End If
    RS1.Close
    Set RS1=Nothing
  Next
  SQL2="Update PLEDGE Set Status='停止' Where Status='授權中' And ((Year(Donate_ToDate)<'"&Year(Donate_Date)&"') Or (Year(Donate_ToDate)='"&Year(Donate_Date)&"' And Month(Donate_ToDate)<'"&Month(Donate_Date)&"')) "
  Set RS2=Conn.Execute(SQL2)
 
   Response.Write("<script>alert('授權轉入捐款成功 ！\n轉入筆數："&FormatNumber(Row,0)&"\n轉入金額："&FormatNumber(Danate_Total,0)&"'); location.href='pledge_transfer.asp'; </script>")

End If
  %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
    <script language="JavaScript">
        function Check_OnClick() {
            if (confirm('※請注意轉入過程中請勿『關閉視窗』')) {
                document.Form.act.value = 'input';
                document.Form.submit();
                //location.href = "pledge_transfer.asp";
            }
        }
    </script>
</head> 
  <body>
   <p><div align="center"><center>
	<form name="Form" method="post" action=""> 
	<input type="hidden" name="act" id="act">
<%
	SQL1="Select Pledge_Id,授權編號=Pledge_Id,捐款人=DONOR.Donor_Name,授權方式=Donate_Payment,扣款金額=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),授權起日=CONVERT(VarChar,Donate_FromDate,111),授權迄日=CONVERT(VarChar,Donate_ToDate,111),轉帳週期=Donate_Period,下次扣款日=Next_DonateDate,有效月年=(Case When Valid_Date<>'' Then Left(Valid_Date,2)+'/'+Right(Valid_Date,2) Else '' End),最後扣款日=Transfer_Date " & _
		 "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And ((Post_SavingsNo<>'' And Post_AccountNo<>'') Or Account_No<>'' Or P_RCLNO<>'') " & _
		 "And Pledge_Id in ("&pledge_id&") Order By Pledge_Id "
	'response.write SQL1
	'response.end()
	Set RS1 = Server.CreateObject("ADODB.RecordSet")
	RS1.Open SQL1,Conn,1,3
	SQL2="Select COUNT(Pledge_Id) AS Count,SUM(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End) as Donate_Total " & _
		 "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And ((Post_SavingsNo<>'' And Post_AccountNo<>'') Or Account_No<>'' Or P_RCLNO<>'') " & _
		 "And Pledge_Id in ("&pledge_id&") "
	'response.write SQL2
	'response.end()
	Set RS2 = Server.CreateObject("ADODB.RecordSet")
	RS2.Open SQL2,Conn,1,3
	FieldsCount = RS1.Fields.Count-1
	Dim I, J
	Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;background-color: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
	Response.Write "<tr>"
	For I = 1 To FieldsCount
		Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
	Next
	Response.Write "</tr>"
	Row=0
	While Not RS1.EOF
	  Row=Row+1	
	Response.Write "<tr "&showhand&" bgcolor='"&bgColor&"' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor="""&bgColor&"""'>"
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
	
	Response.Write "<tr>"
	Response.Write "<td bgcolor='#FFEBEB' colspan='10'><span style='font-size: 10pt; font-family: 新細明體'>總筆數：" & RS2("Count") & " 筆　總金額：" & FormatNumber(RS2("Donate_Total"))& "</span></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	RS1.Close
	Set RS1=Nothing
%>	
    <input type="button" name="check" value="確認送出" class="cbutton" style="cursor:hand" onClick="Check_OnClick()">
	</form>
   </center></div></p>
  </body>
</html>

