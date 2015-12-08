<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/dbinclude_d.asp"-->
<!--#include file="include/email.inc"-->
<%
  If Request.Form("action")="donationvacc" Then
    Donate_DonorName=Request.Form("donatename")
    payment_type=Request.Form("payment_type")
    od_sob=Request.Form("od_sob")
    amt=Request.Form("amt")
    prd_desc=Request.Form("purpose")
    If Len(prd_desc)>20 Then prd_desc=Left(prd_desc,20)
    If Instr(Cstr(Request.ServerVariables("HTTP_REFERER")),"https")>0 Then
      ok_url="https://"&Replace(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL"),"ecbank_vacc_vaccno.asp","ecbank_vacc_ok_url.asp")
    Else
      ok_url="http://"&Replace(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL"),"ecbank_vacc_vaccno.asp","ecbank_vacc_ok_url.asp")
    End If
    vacc_url="https://ecbank.com.tw/gateway.php?mer_id="&ECBank_Code&"&payment_type="&payment_type&"&setbank="&Request.Form("setbank")&"&enc_key="&ECBank_Key&"&od_sob="&od_sob&"&amt="&amt&"&expire_day="&ECBank_BarcodeDate&"&ok_url="&ok_url&""    
    vacc_return=WinHttp(vacc_url,"POST")
    If Instr(vacc_return,"error=0")>0 Then
      ary_vacc_return=Split(vacc_return,"&")
      For I = 0 To UBound(ary_vacc_return)
        ary_return=Split(ary_vacc_return(I),"=")
        If ary_return(0)="error" Then error=ary_return(1)
        If ary_return(0)="mer_id" Then mer_id=ary_return(1)
        If ary_return(0)="tsr" Then tsr=ary_return(1)
        If ary_return(0)="bankcode" Then bankcode=ary_return(1)
        If ary_return(0)="vaccno" Then vaccno=ary_return(1)
        If ary_return(0)="amt" Then amt=ary_return(1)
        If ary_return(0)="expire_date" Then expire_date=ary_return(1)
      Next
      expire_time="235959"
    Else
      error=Split(vacc_return,"=")(1)
      tsr=""
      bankcode=""
      vaccno=""
      expire_date=""
      expire_time=""
      Set RS1=Server.CreateObject("ADODB.RecordSet")
      RS1.Open "Select * From DONATE_API Where API_CODE='"&error&"'",Conn,1,1
      If Not RS1.EOF Then API_DESC=RS1("API_DESC")
      RS1.Close
      Set RS1=Nothing
    End If
    
    '新增綠界交易資料(DONATE_ECBANK)
    Set RS2=Server.CreateObject("ADODB.RecordSet")
    RS2.Open "Select * From DONATE_ECBANK Where mer_id='"&ECBank_Code&"' And payment_type='"&payment_type&"' And od_sob='"&Request.Form("od_sob")&"'", Conn, 1, 1
    If RS2.EOF Then
      SQL3="DONATE_ECBANK"
      Set RS3=Server.CreateObject("ADODB.RecordSet")
      RS3.Open SQL3,Conn,1,3
      RS3.Addnew
      RS3("mer_id")=ECBank_Code
      RS3("payment_type")=payment_type
      RS3("enc_key")=ECBank_Key
      RS3("od_sob")=od_sob
      RS3("amt")=amt
      RS3("prd_desc")=""
      RS3("desc1")=""
      RS3("desc2")=""
      RS3("desc3")=""
      RS3("desc4")=""
      RS3("expire_day")=""
      RS3("setbank")=""
      RS3("ok_url")=ok_url
      RS3("get_error")=error
      RS3("get_original_tsr")=""
      RS3("get_tsr")=tsr
      RS3("get_bankcode")=bankcode
      RS3("get_payno")=vaccno
      RS3("get_barcode1")=""
      RS3("get_barcode2")=""
      RS3("get_barcode3")=""
      RS3("get_expire_date")=expire_date
      RS3("get_expire_time")=expire_time
      RS3.Update
      RS3.Close
      Set RS3=Nothing
    End If
    RS2.Close
    Set RS2=Nothing
  End If
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
    <div id="top"><a href="ecpay.asp">首頁</a> / 線上捐款 / 授權取號</div>
    <div id="mid">
      <div id="mid-1">
        <%If error="0" Then%>
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="page_form">
          <tr>
            <td align="right" class="contents" width="15%">&nbsp;</td>
            <td class="contents">親愛的【<%=Donate_DonorName%>】：</td>
          </tr>
          <tr>
            <td align="right" class="contents" width="15%">&nbsp;</td>
            <td class="contents"><span class="contents">本次線上即時捐款結果，您已經成功取得玉山銀行虛擬帳號。</span></td>
          </tr>
          <tr>
            <td align="right" class="contents" width="15%">&nbsp;</td>
            <td class="contents">繳費銀行代碼：&nbsp;<span class="mustcolumn"><%=bankcode%>&nbsp;(玉山銀行)</span></td>
          </tr>
          <tr>
            <td align="right" class="contents" width="15%">&nbsp;</td>
            <td class="contents">繳費銀行帳號：&nbsp;<span class="mustcolumn"><%=vaccno%></span></td>
          </tr>
          <tr>
            <td align="right" class="contents" width="15%">&nbsp;</td>
            <td class="contents">本次捐款金額：&nbsp;<span class="mustcolumn"><%=FormatNumber(amt,0)%>元</span></td>
          </tr>
          <tr>
            <td align="right" class="contents" width="15%">&nbsp;</td>
            <td class="contents">繳費最後期限：&nbsp;<span class="mustcolumn"><%=Left(expire_date,4)%>/<%=Mid(expire_date,5,2)%>/<%=Right(expire_date,2)%></td>
          </tr>
          <tr>
            <td align="right" class="contents" width="15%">&nbsp;</td>
            <td class="contents">
            	繳費方式說明：<br />
            	1.您可以列印本頁，透過全省ATM自動櫃員機轉帳至繳費銀行帳號。<br />
            	2.您也可透過您的網路銀行或WEB-ATM轉出交易至繳費銀行帳號。<br />
            	3.持玉山銀行金融卡轉帳可免付跨行交易手續費，其它銀行依該行跨行手續費規定扣繳，玉山銀行 WEB-ATM 付款網址 <a href="https://netbank.esunbank.com.tw/webatm/" target="_blank">https://netbank.esunbank.com.tw/webatm/</a>
            </td>
          </tr>

          <tr>
            <td align="center" colspan="2"><br /></td>
          </tr>
          <tr>
            <td align="center" colspan="2">
            	<input type="button" value=" 列  印 " name="But_Exit" class="cbutton" style="cursor:hand;" onClick="javascript:print();">
              <input type="button" value=" 離  開 " name="But_Exit" class="cbutton" style="cursor:hand;" onClick="javascript:Exit_OnClick();">
            </td>
          </tr>
        </table>
        <%Else%>
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="page_form">
          <tr>
            <td align="center"><span class="mustcolumn">銀行虛擬帳號取號執行失敗<br />錯誤訊息:<%=API_DESC%>(<%=error%>)</td>
          </tr>
          <tr>
            <td align="center">
              <input type="button" value=" 再試一次 " name="But_Dn" class="cbutton" style="cursor:hand;" onClick="javascript:Next_OnClick();">&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="button" value=" 離  開 " name="But_Exit" class="cbutton" style="cursor:hand;" onClick="javascript:Exit_OnClick();">
            </td>
          </tr>
        </table>
        <%End If%>
        </table>     
      </div>
    </div>
    <div id="bottom"></div>
  <%Message()%>
</body>
</html>
<!--#include file="include/dbclose.asp"-->