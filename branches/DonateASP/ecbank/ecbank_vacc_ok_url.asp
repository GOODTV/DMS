<!--#include file="include/dbinclude_ok.asp"-->
<!--#include file="include/email.inc"-->
<%
  If Request.Form("mer_id")<>"" And Request.Form("od_sob")<>"" And Request.Form("succ")<>"" Then
    '修改綠界交易資料(DONATE_ECBANK)
    Set RS1=Server.CreateObject("ADODB.RecordSet")
    RS1.Open "Select * From DONATE_ECBANK Where mer_id='"&Request.Form("mer_id")&"' And payment_type='"&Request.Form("payment_type")&"' And od_sob='"&Request.Form("od_sob")&"'", Conn, 1, 3
    If Not RS1.EOF Then
      '交易壓碼驗證
      valid=0
      WebUrl="https://ecbank.com.tw/web_service/get_outmac_valid.php?key="&ECBank_Key&"&serial="&Request.Form("proc_date")&Request.Form("proc_time")&Request.Form("tsr")&"&tac="&Request.Form("tac")&"" 
      tac_valid=WinHttp(WebUrl,"GET")
      If (tac_valid="valid=1") Then valid="1"

      '修改綠界交易資料(DONATE_ECBANK)
      RS1("get_tsr")=Request.Form("tsr")
      RS1("amt")=Cdbl(Request.Form("amt"))
      RS1("get_expire_date")=Request.Form("expire_date")      
      RS1("post_succ")=Request.Form("succ")
      RS1("post_payfrom")=Request.Form("payfrom")
      RS1("post_payer_bank")=Request.Form("payer_bank")
      RS1("post_payer_acc")=Request.Form("payer_acc")
      RS1("post_proc_date")=Request.Form("paid_date")
      RS1("post_proc_time")=Request.Form("proc_time")
      RS1("post_tac")=Request.Form("tac")
      RS1("get_valid")=valid
      If valid=1 Then RS1("close_type")="ok"
      RS1("return_date")=now()
      RS1("return_ip")=Request.ServerVariables("REMOTE_HOST")
      RS1("return_url1")=Cstr(Request.ServerVariables("HTTP_REFERER"))
      RS1("return_url2")=Cstr(Request.ServerVariables("SERVER_NAME"))
      RS1.Update 
      
      '判斷是否新增捐款資料(DONATE) 
      If valid=1 And Request.Form("succ")="1" Then 
        Set RS2=Server.CreateObject("ADODB.RecordSet")
        RS2.Open "Select * From DONATE_WEB Where od_sob='"&Request.Form("od_sob")&"'", Conn, 1, 1
        If Not RS2.EOF Then
          Set RS3=Server.CreateObject("ADODB.RecordSet")
          RS3.Open "Select * From DONATE Where od_sob='"&Request.Form("od_sob")&"'", Conn, 1, 1
          If RS3.EOF Then
            '捐款日期
            Donate_Date=Date()
            If Request.Form("paid_date")<>"" Then Donate_Date=Left(Request.Form("paid_date"),4)&"/"&Mid(Request.Form("paid_date"),5,2)&"/"&Right(Request.Form("paid_date"),2)
          
            '捐款收據編號
            'Invoice_No=Get_InvoiceNo(RS2("Dept_Id"),Donate_Date,"1",RS2("Donate_Invoice_Type"))

            '捐款前置碼
            'Invoice_Pre=""
            'If Invoice_No<>"" Then
            '  Set RS4=Server.CreateObject("ADODB.RecordSet")
            '  RS4.Open "Select * From DEPT Where Dept_Id='"&RS2("Dept_Id")&"'", Conn, 1, 1
            '  Invoice_Pre=RS4("Invoice_Pre")
            '  RS4.Close
            '  Set RS4=Nothing
            'End If

            Invoice_Pre=""
            Invoice_No=""
            Act_id=""
            If RS2("Donate_Act_Id")<>"" Then Act_id=Cstr(RS2("Donate_Act_Id"))
            InvoiceNo=Get_Invoice_No2("1",Cstr(RS2("Dept_Id")),Cstr(Date()),Cstr(RS2("Donate_Invoice_Type")),Act_id)
            If InvoiceNo<>"" Then
              Invoice_Pre=Split(InvoiceNo,"/")(0)
              Invoice_No=Split(InvoiceNo,"/")(1)
            End If
            
            '捐款資料
            Donate_Fee=0
            Donate_Accou=0
            If Request.Form("amt")<>"" Then
              Donate_Fee=10
              Donate_Accou=CDbl(Request.Form("amt"))-CDbl(Donate_Fee)
            End If
            SQL4="DONATE"
            Set RS4=Server.CreateObject("ADODB.RecordSet")
            RS4.Open SQL4,Conn,1,3
            RS4.Addnew
            RS4("od_sob")=Request.Form("od_sob")
            RS4("Donor_Id")=RS2("Donor_Id")
            RS4("Donate_Date")=Donate_Date
            RS4("Donate_Amt")=Cdbl(Request.Form("amt"))
            RS4("Donate_AmtD")=Cdbl(Request.Form("amt"))
            RS4("Donate_Fee")=Donate_Fee
            RS4("Donate_FeeD")=Donate_Fee
            RS4("Donate_Accou")=Donate_Accou
            RS4("Donate_AccouD")=Donate_Accou
            RS4("Donate_AmtM")="0"
            RS4("Donate_FeeM")="0"
            RS4("Donate_AccouM")="0"
            RS4("Donate_AmtA")="0"
            RS4("Donate_FeeA")="0"
            RS4("Donate_AccouA")="0"
            RS4("Donate_AmtS")="0"
            RS4("Donate_RateS")="0"
            RS4("Donate_FeeS")="0"
            RS4("Donate_AccouS")="0"
            RS4("Donate_Payment")="銀行虛擬帳號"
            RS4("Donate_Purpose")=RS2("Donate_Purpose")
            RS4("Donate_Purpose_Type")="D"
            RS4("Donate_Type")="單次捐款"
            RS4("Donate_Forign")=""
            RS4("Donate_Desc")=""
            RS4("IsBeductible")="N"
            RS4("Donate_Amt2")="0"
            RS4("Card_Bank")=""
            RS4("Card_Type")=""
            RS4("Account_No")=""
            RS4("Valid_Date")=""
            RS4("Card_Owner")=""
            RS4("Owner_IDNo")=""
            RS4("Relation")=""
            RS4("Authorize")=""
            RS4("Check_No")=""
            RS4("Check_ExpireDate")=null
            RS4("Post_Name")=""
            RS4("Post_IDNo")=""
            RS4("Post_SavingsNo")=""
            RS4("Post_AccountNo")=""
            RS4("Dept_Id")=RS2("Dept_Id")
            RS4("Invoice_Title")=RS2("Donate_Invoice_Title")
            RS4("Invoice_Pre")=Invoice_Pre
            RS4("Invoice_No")=Invoice_No
            RS4("Invoice_Print")="0"
            RS4("Invoice_Print_Add")="0"
            RS4("Invoice_Print_Yearly_Add")="0"
            RS4("Request_Date")=null
            RS4("Accoun_Bank")=""
            RS4("Accoun_Date")=null
            RS4("Invoice_type")=RS2("Donate_Invoice_Type")
            RS4("Accounting_Title")=""
            If RS2("Donate_Act_Id")<>"" Then
              RS4("Act_id")=RS2("Donate_Act_Id")
            Else
              RS4("Act_id")=null
            End If
            RS4("Comment")=""
            RS4("Invoice_PrintComment")=""
            RS4("Issue_Type")=""
            RS4("Issue_Type_Keep")=""
            RS4("Export")="N"
            RS4("Create_Date")=Date()
            RS4("Create_DateTime")=Now()
            RS4("Create_User")="線上金流"
            RS4("Create_IP")=Request.ServerVariables("REMOTE_ADDR")
            RS4.Update
            RS4.Close
            Set RS4=Nothing 
          End If
          RS3.Close
          Set RS3=Nothing
        End If
        RS2.Close
        Set RS2=Nothing
      End If
      
      '回覆己收到資料
      If valid=1 Then Server.Transfer "ecbank_ok_write.asp"
    End If
    RS1.Close
    Set RS1=Nothing
  End If
%>
<!--#include file="include/dbclose.asp"-->