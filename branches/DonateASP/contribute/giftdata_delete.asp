<!--#include file="../include/dbfunctionJ.asp"-->
<%
  
  SQL="Delete GIFTDATA Where Ser_No='"&request("Ser_No")&"'"
  Set RS=Conn.Execute(SQL)
        
  session("errnumber")=1
  session("msg")="『"&Goods_Name&"』 刪除成功 ！"
   
  response.redirect "gift_edit.asp?ctype="&request("ctype")&"&gift_id="&request("gift_id")&""
%><!--#include file="../include/dbclose.asp"-->