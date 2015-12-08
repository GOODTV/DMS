<!--#include file="../include/dbfunctionJ.asp"-->
<%
  SQL="Delete From LINKED2 Where Ser_No='"&request("ser_no")&"'"
  Set RS=Conn.Execute(SQL)
  session("errnumber")=1
  session("msg")=request("subject")&"類別刪除成功 ！"
  response.redirect "linked2_add.asp?linked_type="&request("Linked_Type")&"&subject="&request("subject")&"&linked_id="&request("linked_id")&""
%>