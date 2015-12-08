<!--#include file="../include/dbfunction.asp"-->
<% 
  sql="delete from prog_menu where ser_no = '"&request("ser_no")&"'"
  call UpdateSQL (SQL)
  Response.Redirect "menu_edit.asp?menu_id="&session("menu_id")
%>