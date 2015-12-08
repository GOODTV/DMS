<!--#include file="../include/dbfunctionJ.asp"-->
<%
  SQL="Delete From CASECODE Where Ser_No='"&request("ser_no")&"'"
  Set RS=Conn.Execute(SQL)
  session("errnumber")=1
  session("msg")="代碼選項刪除成功 ！"
  response.redirect "casecode_add.asp?codetype="&session("codetype")
%>