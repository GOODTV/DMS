<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/dbinclude_ok.asp"-->
<!--#include file="include/email.inc"-->
<%
  Session.Contents.Remove("get_od_sob")
  If Request.Form("od_sob")<>"" And Request.Form("succ")<>"" Then
    Set RS1=Server.CreateObject("ADODB.RecordSet")
    RS1.Open "Select * From DONATE_WEB Where od_sob='"&Request.Form("od_sob")&"'", Conn, 1, 1
    If Not RS1.EOF Then
      '新增綠界交易資料(DONATE_ECPAY)
      Set RS2=Server.CreateObject("ADODB.RecordSet")
      RS2.Open "Select * From DONATE_ECPAY Where od_sob='"&Request.Form("od_sob")&"'", Conn, 1, 1
      If RS2.EOF Then
        SQL3="DONATE_ECPAY"
        Set RS3=Server.CreateObject("ADODB.RecordSet")
        RS3.Open SQL3,Conn,1,3
        RS3.Addnew
        RS3("ecbank_type")="creditcard"
        RS3("od_sob")=Request.Form("od_sob")
        RS3("succ")=Request.Form("succ")
        RS3("gwsr")=Request.Form("gwsr")
        RS3("response_code")=Request.Form("response_code")
        RS3("response_msg")=Request.Form("response_msg")
        RS3("process_date")=Request.Form("process_date")
        RS3("process_time")=Request.Form("process_time")
        RS3("auth_code")=Request.Form("auth_code")
        RS3("amount")=Request.Form("amount")
        RS3("od_hoho")=Request.Form("od_hoho")
        RS3("eci")=Request.Form("eci")
        RS3("inspect")=Request.Form("inspect")
        RS3("spcheck")=Request.Form("spcheck")
        RS3("card4no")=Request.Form("card4no")
        RS3("return_date")=Now()
        RS3("return_ip")=Request.ServerVariables("REMOTE_HOST")
        RS3("return_url1")=Cstr(Request.ServerVariables("HTTP_REFERER"))
        RS3("return_url2")=Cstr(Request.ServerVariables("SERVER_NAME"))
        RS3("close_type")=""
        RS3("close_url")=""
        RS3.Update
        RS3.Close
        Set RS3=Nothing
      End If
      RS2.Close
      Set RS2=Nothing
      
      If Request.Form("succ")="1" Then   
        '判斷是否新增捐款資料(DONATE) 
        Set RS2=Server.CreateObject("ADODB.RecordSet")
        RS2.Open "Select * From DONATE Where od_sob='"&Request.Form("od_sob")&"'", Conn, 1, 1
        If RS2.EOF Then
          '捐款收據編號
          'Invoice_No=Get_InvoiceNo(RS1("Dept_Id"),Date(),"1",RS1("Donate_Invoice_Type"))
          
          '捐款前置碼
          'Invoice_Pre=""
          'If Invoice_No<>"" Then
          '  Set RS3=Server.CreateObject("ADODB.RecordSet")
          '  RS3.Open "Select * From DEPT Where Dept_Id='"&RS1("Dept_Id")&"'", Conn, 1, 1
          '  Invoice_Pre=RS3("Invoice_Pre")
          '  RS3.Close
          '  Set RS3=Nothing
          'End If

          Invoice_Pre=""
          Invoice_No=""
          Act_id=""
          If RS1("Donate_Act_Id")<>"" Then Act_id=Cstr(RS1("Donate_Act_Id"))
          InvoiceNo=Get_Invoice_No2("1",Cstr(RS1("Dept_Id")),Cstr(Date()),Cstr(RS1("Donate_Invoice_Type")),Act_id)
          If InvoiceNo<>"" Then
            Invoice_Pre=Split(InvoiceNo,"/")(0)
            Invoice_No=Split(InvoiceNo,"/")(1)
          End If
          
          '新增捐款紀錄
          Donate_Fee=0
          Donate_Accou=0
          If Request.Form("amount")<>"" Then
            Donate_Fee=CDbl(Request.Form("amount"))*0.03
            Donate_Accou=CDbl(Request.Form("amount"))-CDbl(Donate_Fee)
          End If
          SQL3="DONATE"
          Set RS3=Server.CreateObject("ADODB.RecordSet")
          RS3.Open SQL3,Conn,1,3
          RS3.Addnew
          RS3("od_sob")=Request.Form("od_sob")
          RS3("Donor_Id")=RS1("Donor_Id")
          RS3("Donate_Date")=Date()
          RS3("Donate_Amt")=Request.Form("amount")
          RS3("Donate_AmtD")=Request.Form("amount")
          RS3("Donate_Fee")=Donate_Fee
          RS3("Donate_FeeD")=Donate_Fee
          RS3("Donate_Accou")=Donate_Accou
          RS3("Donate_AccouD")=Donate_Accou
          RS3("Donate_AmtM")="0"
          RS3("Donate_FeeM")="0"
          RS3("Donate_AccouM")="0"
          RS3("Donate_AmtA")="0"
          RS3("Donate_FeeA")="0"
          RS3("Donate_AccouA")="0"
          RS3("Donate_AmtS")="0"
          RS3("Donate_RateS")="0"
          RS3("Donate_FeeS")="0"
          RS3("Donate_AccouS")="0"
          RS3("Donate_Payment")="網路信用卡"
          RS3("Donate_Purpose")=RS1("Donate_Purpose")
          RS3("Donate_Purpose_Type")="D"
          RS3("Donate_Type")="單次捐款"
          RS3("Donate_Forign")=""
          RS3("Donate_Desc")=""
          RS3("IsBeductible")="N"
          RS3("Donate_Amt2")="0"
          RS3("Card_Bank")=""
          RS3("Card_Type")=""
          RS3("Account_No")=""
          RS3("Valid_Date")=""
          RS3("Card_Owner")=""
          RS3("Owner_IDNo")=""
          RS3("Relation")=""
          RS3("Authorize")=""
          RS3("Check_No")=""
          RS3("Check_ExpireDate")=null
          RS3("Post_Name")=""
          RS3("Post_IDNo")=""
          RS3("Post_SavingsNo")=""
          RS3("Post_AccountNo")=""
          RS3("Dept_Id")=RS1("Dept_Id")
          RS3("Invoice_Title")=RS1("Donate_Invoice_Title")
          RS3("Invoice_Pre")=Invoice_Pre
          RS3("Invoice_No")=Invoice_No
          RS3("Invoice_Print")="0"
          RS3("Invoice_Print_Add")="0"
          RS3("Invoice_Print_Yearly_Add")="0"
          RS3("Request_Date")=null
          RS3("Accoun_Bank")=""
          RS3("Accoun_Date")=null
          RS3("Invoice_type")=RS1("Donate_Invoice_Type")
          RS3("Accounting_Title")=""
          If RS1("Donate_Act_Id")<>"" Then
            RS3("Act_id")=RS1("Donate_Act_Id")
          Else
            RS3("Act_id")=null
          End If
          RS3("Comment")=""
          RS3("Invoice_PrintComment")=""
          RS3("Issue_Type")=""
          RS3("Issue_Type_Keep")=""
          RS3("Export")="N"
          RS3("Create_Date")=Date()
          RS3("Create_DateTime")=Now()
          RS3("Create_User")="線上金流"
          RS3("Create_IP")=Request.ServerVariables("REMOTE_ADDR")
          RS3.Update
          RS3.Close
          Set RS3=Nothing
        End If
        RS2.Close
        Set RS2=Nothing
      
        '交易成功關帳
        close_url=get_close_url(Request.Form("gwsr"),Request.Form("amount"))
        If close_url<>"3" Then
          'For I=1 To 10
            close_type=WinHttp(close_url,"POST")
          '  If close_type="ok" Then Exit For
          'Next
        Else
          close_type="ok"
        End If
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open "Select * From DONATE_ECPAY Where od_sob = '"&Request.Form("od_sob")&"'", Conn, 1, 3
        If Not RS2.EOF Then
          RS2("succ")="1"
          RS2("close_type") = close_type
          RS2("close_url") = close_url
          RS2.Update
        End If
        RS2.Close
        Set RS2 = Nothing
        
        '收據寄送地址
        Donate_Invoice_Address=""
        If RS1("Donate_Invoice_AreaCode")<>"" Then
          SQL2="Select Invoice_Address=B.mValue+Left(A.mCode,3)+A.mValue From CODECITY A Join CODECITY B On Substring(A.codeMetaID,7,1)=B.mCode Where A.mCode='"&RS1("Donate_Invoice_AreaCode")&"'"
          Set RS2=Server.CreateObject("ADODB.RecordSet")
          RS2.Open SQL2,Conn,1,1
          If Not RS2.EOF Then Donate_Invoice_Address=RS2("Invoice_Address")&RS1("Donate_Invoice_Address")
          RS2.Close
          Set RS2=Nothing
        Else
          Donate_Invoice_Address=RS1("Donate_Invoice_Address")
        End If 
            
        '寄發E-Mail
        If Request.Form("email")<>"" Then
          ToName=RS1("Donate_DonorName")
          ToEmail=Request.Form("email")
          cc=""
          AttachFile=""
          MailType="html"
          For I=1 To 6
            Randomize
            mc=mc&LCase(Chr(int(26*Rnd)+65))
          Next
          MailSubject="線上捐款明細"
          MailBody = ""  & "<br>"
          MailBody = MailBody & "親愛的【" & RS1("Donate_DonorName") & "】：" & "<br>"
          MailBody = MailBody & "本次線上即時捐款結果，您的信用卡已成功取得授權。" & "<br>"
          MailBody = MailBody & "以下是您的個人資料及本次刷卡明細，請核對並妥善保存。" &  "<br>"
          MailBody = MailBody & "" & "<br>"
          MailBody = MailBody & "捐款人：" & RS1("Donate_DonorName") & "<br>"
          MailBody = MailBody & "捐款用途：" & RS1("Donate_Purpose") & "<br>"
          MailBody = MailBody & "捐款金額：" & FormatNumber(Request.Form("amount"),0)&"元" & "<br>"
          MailBody = MailBody & "交易序號：" & Request.Form("od_sob")& "<br>"
          If Cstr(RS1("Donate_Invoice_Type"))<>InvoiceTypeN Then
            MailBody = MailBody & "收據抬頭：" & RS1("Donate_Invoice_Title") & "<br>"
            MailBody = MailBody & "收據身分證/統編：" & RS1("Donate_Invoice_IDNO") & "<br>"
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
    Else
      session("errnumber")=1
      session("msg")="很抱歉您輸入的訊息不足，誠摯地請您再試一次！"
      Response.Redirect "ecpay.asp"
    End If
    RS1.Close
    Set RS1=Nothing
  Else
    session("errnumber")=1
    session("msg")="很抱歉您輸入的訊息不足，誠摯地請您再試一次！"
    Response.Redirect "ecpay.asp"
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
  <script src="js/ecpayb.js" type="text/javascript"></script>
  <!--[if lt IE 8]>  <script src="js/IE8.js"></script>  <![endif]-->
  <title><%=Comp_Name%></title>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onbeforecopy="return false" onselect="document.selection.empty()" oncopy="document.selection.empty()">
  <noscript><iframe src=*.html></iframe></noscript>	
  <div id="wrapper">
    <div id="header"><img src="image/banner.jpg" alt="<%=Comp_Name%>" />  </div>
    <div id="top"><a href="ecpay.asp">首頁</a> / 線上捐款 / 授權結果</div>
    <div id="mid">
      <div id="mid-1">
        <table width="100%" height="270" border="0" align="center" cellpadding="0" cellspacing="0" class="page_form">
        <%If Request.Form("succ")="1" Then%>
          <tr>
            <td align="right" class="contents">&nbsp;</td>
            <td class="contents"><span class="verifycode">親愛的【<%=Donate_DonorName%>】：</span></td>
          </tr>
          <tr>
            <td align="right" class="contents">&nbsp;</td>
            <td class="contents"><span class="verifycode">本次線上即時捐款結果，您的信用卡已成功取得授權。</span></td>
          </tr>
          <tr>
            <td align="right" class="contents">&nbsp;</td>
            <td class="contents"><span class="verifycode">以下是您的個人資料及本次捐款明細。</span></td>
          </tr>
          <tr>
            <td width="180" align="right" class="contents">捐款金額：&nbsp;</td>
            <td width="470" class="contents"><%=FormatNumber(Request.Form("amount"),0)%>元(信用卡捐款)&nbsp;</td>
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
            <td colspan="2" class="contents"><span class="verifycode">親愛的【<%=Donate_DonorName%>】：</span></td>
          </tr>
          <tr>
            <td colspan="2" class="contents"><span class="verifycode">很抱歉，本次線上捐款並未取得發卡銀行的授權。</span></td>
          </tr>
          <tr>
            <td colspan="2" class="contents"><span class="verifycode">誠摯地請您再試一次，並填寫正確資料，謝謝！</span></td>
          </tr>
          <tr>
            <td colspan="2" class="contents"><span class="verifycode">失敗原因：<%=Request.Form("response_msg")%></span></td>
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
      <div id="bottom"></div>
    </div>
  </div>
  <%Message()%>
</body>
</html>
<!--#include file="include/dbclose.asp"-->
