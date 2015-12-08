<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/dbinclude_d.asp"-->
<!--#include file="include/email.inc"-->
<%
  Session.Contents.Remove("get_od_sob")
  If Request.Form("action")="donationbarcode" Then
    payment_type=Request.Form("payment_type")
    od_sob=Request.Form("od_sob")
    amt=Request.Form("amt")
    If Instr(Cstr(Request.ServerVariables("HTTP_REFERER")),"https")>0 Then
      ok_url="https://"&Replace(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL"),"ecbank_barcode_print.asp","ecbank_ok_url.asp")
    Else
      ok_url="http://"&Replace(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL"),"ecbank_barcode_print.asp","ecbank_ok_url.asp")
    End If
    barcode_url="https://ecbank.com.tw/gateway.php?mer_id="&ECBank_Code&"&payment_type="&payment_type&"&enc_key="&ECBank_Key&"&od_sob="&od_sob&"&amt="&amt&"&expire_day="&ECBank_BarcodeDate&"&ok_url="&ok_url&""
    barcode_return=WinHttp(barcode_url,"POST")
    If Instr(barcode_return,"error=0")>0 Then
      ary_barcode_return=Split(barcode_return,"&")
      For I = 0 To UBound(ary_barcode_return)
        ary_return=Split(ary_barcode_return(I),"=")
        If ary_return(0)="error" Then error=ary_return(1)
        If ary_return(0)="mer_id" Then mer_id=ary_return(1)
        If ary_return(0)="tsr" Then tsr=ary_return(1)
        If ary_return(0)="amt" Then amt=ary_return(1)
        If ary_return(0)="expire_date" Then expire_date=ary_return(1)
        If ary_return(0)="barcode1" Then barcode1=ary_return(1)
        If ary_return(0)="barcode2" Then barcode2=ary_return(1)
        If ary_return(0)="barcode3" Then barcode3=ary_return(1)
      Next
    Else
      error=Split(barcode_return,"=")(1)
      tsr=""
      expire_date=""
      barcode1=""
      barcode2=""
      barcode3=""
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
      RS3("expire_day")=expire_day
      RS3("setbank")=""
      RS3("ok_url")=ok_url
      RS3("get_error")=error
      RS3("get_original_tsr")=""
      RS3("get_tsr")=tsr
      RS3("get_payno")=""
      RS3("get_barcode1")=barcode1
      RS3("get_barcode2")=barcode2
      RS3("get_barcode3")=barcode3
      RS3("get_expire_date")=expire_date
      RS3("get_expire_time")="235959"
      RS3.Update
      RS3.Close
      Set RS3=Nothing
    End If
    RS2.Close
    Set RS2=Nothing
    If tsr<>"" Then Response.Redirect "https://ecbank.com.tw/order/barcode_print.php?mer_id="&ECBank_Code&"&tsr="&tsr&""
  End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <link href="include/donate.css" rel="stylesheet" type="text/css">
  <!--[if lt IE 8]>  <script src="js/IE8.js"></script>  <![endif]-->  	
  <title><%=Comp_Name%></title>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onselect="document.selection.empty()" oncopy="document.selection.empty()" onbeforecopy="return false">
  <noscript><iframe src=*.html></iframe></noscript>	
  <div id="wrapper">
    <div id="header"><img src="image/banner.jpg" alt="<%=Comp_Name%>" />  </div>
    <div id="top"><a href="ecpay.asp">首頁</a> / 線上捐款 / 列印捐款單失敗</div>
    <div id="mid">
      <div id="mid-1">
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="page_form">
          <tr>
            <td align="center"><span class="mustcolumn">列印捐款單執行失敗<br />錯誤訊息:<%=API_DESC%></td>
          </tr>
        </table>     
      </div>
    </div>
    <div id="bottom"></div>
  <%Message()%>
</body>
</html>
<!--#include file="include/dbclose.asp"-->