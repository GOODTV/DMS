<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/dbinclude_ok.asp"-->
<!--#include file="include/email.inc"-->
<%
  Session.Contents.Remove("get_od_sob")
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
      RS1("get_original_tsr")=Request.Form("original_tsr")
      RS1("get_payno")=Request.Form("payno")
      RS1("amt")=Cdbl(Request.Form("amt"))
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
            RS4("Donate_Payment")="WEB-ATM"
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
            
            '收據寄送地址
            Donate_Invoice_Address=""
            If RS2("Donate_Invoice_AreaCode")<>"" Then
              SQL4="Select Invoice_Address=B.mValue+Left(A.mCode,3)+A.mValue From CODECITY A Join CODECITY B On Substring(A.codeMetaID,7,1)=B.mCode Where A.mCode='"&RS2("Donate_Invoice_AreaCode")&"'"
              Set RS4=Server.CreateObject("ADODB.RecordSet")
              RS4.Open SQL4,Conn,1,1
              If Not RS4.EOF Then Donate_Invoice_Address=RS4("Invoice_Address")&RS2("Donate_Invoice_Address")
              RS4.Close
              Set RS4=Nothing
            Else
              Donate_Invoice_Address=RS2("Donate_Invoice_Address")
            End If
          
            '寄發E-Mail
            If RS2("Donate_Email")<>"" Then
              ToName=RS2("Donate_DonorName")
              ToEmail=RS2("Donate_Email")
              cc=""
              AttachFile=""
              MailType="html"
              For I=1 To 6
                Randomize
                mc=mc&LCase(Chr(int(26*Rnd)+65))
              Next
              MailSubject="線上捐款明細"
              MailBody = ""  & "<br>"
              MailBody = MailBody & "親愛的【" & RS2("Donate_DonorName") & "】：" & "<br>"
              MailBody = MailBody & "本次線上即時捐款結果，您的金融卡已成功取得扣款。" & "<br>"
              MailBody = MailBody & "以下是您的個人資料及本次捐款明細，請核對並妥善保存。" &  "<br>"
              MailBody = MailBody & "" & "<br>"
              MailBody = MailBody & "捐款人：" & RS2("Donate_DonorName") & "<br>"
              MailBody = MailBody & "捐款用途：" & RS2("Donate_Purpose") & "<br>"
              MailBody = MailBody & "捐款金額：" & FormatNumber(Request.Form("amt"),0)&"元" & "<br>"
              MailBody = MailBody & "交易序號：" & Request.Form("od_sob")& "<br>"
              If Cstr(RS2("Donate_Invoice_Type"))<>InvoiceTypeN Then
                MailBody = MailBody & "收據抬頭：" & RS2("Donate_Invoice_Title") & "<br>"
                MailBody = MailBody & "收據身分證/統編：" & RS2("Donate_Invoice_IDNO") & "<br>"
                MailBody = MailBody & "收據地址：" & Donate_Invoice_Address & "<br>"
              End If
              MailBody = MailBody & "" & "<br>"
              MailBody = MailBody & "再次感謝您的支持與愛護！" & "<br>"
              MailBody = MailBody & "--------------------------------------------------------------------------------------------------<br>"
              MailBody = MailBody & "此信件由系統自動發送，請勿直接回應！<br>"
              call Send_EMail(SendMailType,Comp_Name,SendMailDefault,ToName,ToEmail,cc,MailSubject,MailType,MailBody,AttachFile,mc)
            End If
      
            '寄發簡訊
          
          End If
          RS3.Close
          Set RS3=Nothing
        End If
        RS2.Close
        Set RS2=Nothing
      End If
    End If
    RS1.Close
    Set RS1=Nothing
  End If
  
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open "Select * From DONATE_WEB Where od_sob='"&Request.Form("od_sob")&"'", Conn, 1, 1
  If Not RS1.EOF Then
    Donate_DonorName=RS1("Donate_DonorName")
    Donate_Purpose=RS1("Donate_Purpose")
    Donate_Invoice_Title=RS1("Donate_Invoice_Title")
    Donate_Invoice_IDNo=RS1("Donate_Invoice_IDNo")
    Donate_Invoice_No=""
    Donate_Invoice_Address=""
    If Cstr(RS1("Donate_Invoice_Type"))<>InvoiceTypeN Then
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open "Select Invoice_No=Invoice_Pre+Invoice_No From DONATE Where od_sob = '"&Request.Form("od_sob")&"'", Conn, 1, 1
      If Not RS2.EOF Then Donate_Invoice_No=RS2("Invoice_No")
      RS2.Close
      Set RS2=Nothing
        
      If RS1("Donate_Invoice_AreaCode")<>"" Then
        SQL2="Select Invoice_Address=B.mValue+Left(A.mCode,3)+A.mValue From CODECITY A Join CODECITY B On Substring(A.codeMetaID,7,1)=B.mCode Where A.mCode='"&RS1("Donate_Invoice_AreaCode")&"'"
        Set RS2=Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
        If Not RS2.EOF Then Donate_Invoice_Address=RS2("Invoice_Address")&RS1("Donate_Invoice_Address")
        RS2.Close
        Set RS2=Nothing
      End If
    End If  
  End If
  RS1.Close
  Set RS1=Nothing  
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <link href="include/donate.css" rel="stylesheet" type="text/css">
  <script src="js/npois.js" type="text/javascript"></script>
  <script src="js/ecpayb.js" type="text/javascript"></script>
  <!--[if lt IE 8]>  <script src="js/IE8.js"></script>  <![endif]-->  
  <title><%=Comp_Name%></title>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onselect="document.selection.empty()" oncopy="document.selection.empty()" onbeforecopy="return false">
  <noscript><iframe src=*.html></iframe></noscript>	
  <div id="wrapper">
    <div id="header"><img src="image/banner.jpg" alt="<%=Comp_Name%>" />  </div>
    <div id="top"><a href="ecpay.asp">首頁</a> / 線上捐款 / 授權結果</div>
    <div id="mid">
      <div id="mid-1">
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="page_form">
        <%If Request.Form("succ")="1" Then%>
          <tr>
            <td align="right" class="contents">&nbsp;</td>
            <td class="contents"><span class="mustcolumn">親愛的【<%=Donate_DonorName%>】：</span></td>
          </tr>
          <tr>
            <td align="right" class="contents">&nbsp;</td>
            <td class="contents"><span class="mustcolumn">本次線上即時捐款結果，您的金融卡已成功取得扣款。</span></td>
          </tr>
          <tr>
            <td align="right" class="contents">&nbsp;</td>
            <td class="contents"><span class="mustcolumn">以下是您的個人資料及本次捐款明細。</span></td>
          </tr>
          <tr>
            <td width="180" align="right" class="contents">捐款金額：&nbsp;</td>
            <td width="470" class="contents"><%=FormatNumber(Request.Form("amt"),0)%>元(WEB-ATM捐款)&nbsp;</td>
          </tr>
          <tr>
            <td align="right" class="contents">捐款用途：&nbsp;</td>
            <td class="contents"><%=Donate_Purpose%>&nbsp;</td>
          </tr>
          <tr>
            <td align="right" class="contents">交易序號：&nbsp;</td>
            <td class="contents"><%=Request.Form("od_sob")%>&nbsp;</td>
          </tr>
          <%If Donate_Invoice_No<>"" Then%>
          <tr>
            <td align="right" class="contents">收據抬頭：&nbsp;</td>
            <td class="contents"><%=Donate_Invoice_Title%>&nbsp;</td>
          </tr>
          <tr>
            <td align="right" class="contents">收據身份證：&nbsp;</td>
            <td class="contents"><%=Donate_Invoice_IDNo%>&nbsp;</td>
          </tr>
          <tr>
            <td align="right" class="contents">收據寄送地址：&nbsp;</td>
            <td class="contents"><%=Donate_Invoice_Address%>&nbsp;</td>
          </tr>
          <%End If%>
        <%Else%>
          <tr>
            <td colspan="2" class="contents"><span class="mustcolumn">親愛的【<%=Donate_DonorName%>】：</span></td>
          </tr>
          <tr>
            <td colspan="2" class="contents"><span class="mustcolumn">很抱歉，本次線上捐款並未取得發卡銀行的授權。</span></td>
          </tr>
          <tr>
            <td colspan="2" class="contents"><span class="mustcolumn">誠摯地請您再試一次，並填寫正確資料，謝謝！</span></td>
          </tr>
        <%End If%>
          <tr>
            <%If Request.Form("succ")="1" Then%>
            <td align="center" colspan="2"><input type="button" name="But_Exit" value=" 確  定 " class="cbutton" style="cursor:hand;" onclick="javascript:Exit_OnClick();"></td>
            <%Else%>
            <td align="center" colspan="2">
              <input type="button" value=" 再試一次 " name="But_Dn" class="cbutton" style="cursor:hand;" onClick="javascript:Next_OnClick();">&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="button" value=" 離  開 " name="But_Exit" class="cbutton" style="cursor:hand;" onClick="javascript:Exit_OnClick();">
            </td>
            <%End If%>
          </tr>
        </table>     
      </div>
    </div>
    <div id="bottom"></div>
  <%Message()%>
</body>
</html>
<!--#include file="include/dbclose.asp"-->