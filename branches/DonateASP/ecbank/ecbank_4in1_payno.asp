<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/dbinclude_d.asp"-->
<!--#include file="include/email.inc"-->
<%
  If Request.Form("action")="donation4in1" Then
    Donate_DonorName=Request.Form("donatename")
    payment_type=Request.Form("payment_type")
    od_sob=Request.Form("od_sob")
    amt=Request.Form("amt")
    prd_desc=Request.Form("purpose")
    If Len(prd_desc)>20 Then prd_desc=Left(prd_desc,20)
    If Instr(Cstr(Request.ServerVariables("HTTP_REFERER")),"https")>0 Then
      ok_url="https://"&Replace(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL"),"ecbank_4in1_payno.asp","ecbank_ok_url.asp")
    Else
      ok_url="http://"&Replace(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL"),"ecbank_4in1_payno.asp","ecbank_ok_url.asp")
    End If
    ibon_url="https://ecbank.com.tw/gateway.php?mer_id="&ECBank_Code&"&payment_type="&payment_type&"&enc_key="&ECBank_Key&"&od_sob="&od_sob&"&amt="&amt&"&prd_desc="&prd_desc&"&ok_url="&ok_url&""
    ibon_return=WinHttp(ibon_url,"POST")
    If Instr(ibon_return,"error=0")>0 Then
      ary_ibon_return=Split(ibon_return,"&")
      For I = 0 To UBound(ary_ibon_return)
        ary_return=Split(ary_ibon_return(I),"=")
        If ary_return(0)="error" Then error=ary_return(1)
        If ary_return(0)="mer_id" Then mer_id=ary_return(1)
        If ary_return(0)="tsr" Then tsr=ary_return(1)
        If ary_return(0)="od_sob" Then od_sob=ary_return(1)
        If ary_return(0)="amt" Then amt=ary_return(1)
        If ary_return(0)="payno" Then payno=ary_return(1)
        If ary_return(0)="expire_date" Then expire_date=ary_return(1)
        If ary_return(0)="expire_time" Then expire_time=ary_return(1)
      Next
    Else
      error=Split(ibon_return,"=")(1)
      tsr=""
      payno=""
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
      RS3("prd_desc")=prd_desc
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
      RS3("get_payno")=payno
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
            <td class="contents"><span class="contents">本次線上即時捐款結果，您已經成功取得繳費代碼。</span></td>
          </tr>
          <tr>
            <td align="right" class="contents" width="15%">&nbsp;</td>
            <td class="contents">繳費代碼：&nbsp;<span class="mustcolumn"><%=payno%></span></td>
          </tr>
          <tr>
            <td align="right" class="contents" width="15%">&nbsp;</td>
            <td class="contents">煩請您牢記繳費代碼並於&nbsp;<span class="mustcolumn"><%=Left(expire_date,4)%>/<%=Mid(expire_date,5,2)%>/<%=Right(expire_date,2)%>&nbsp;<%=Left(expire_time,2)%>:<%=Right(expire_time,2)%></span>&nbsp;之前於便利商店列印繳費</td>
          </tr>
          <tr>
            <td align="center" colspan="2"><br /></td>
          </tr>
          <tr>
            <td align="center" colspan="2">
              <input type="button" value=" 離  開 " name="But_Exit" class="cbutton" style="cursor:hand;" onClick="javascript:Exit_OnClick();">
            </td>
          </tr>
        </table>
        <%Else%>
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="page_form">
          <tr>
            <td align="center"><span class="mustcolumn">繳費代碼取號執行失敗<br />錯誤訊息:<%=API_DESC%>(<%=error%>)</td>
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