<!--#include file="../include/dbfunction.asp"-->
<% 
   sql="delete from user_prog where ser_no = "&request("ser_no")&""
   call UpdateSQL (SQL)
   Response.Redirect "prog_edit.asp?prog_id="&session("prog_id")
%>