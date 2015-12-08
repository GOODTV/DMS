<!--#include file="../_function/dbfunction.asp"-->
<% 
sql="delete from user_prog where ser_no= '"&request("ser_no")&"' "
call UpdateSQL (SQL)
Server.Transfer "prog_edit.asp"
%>

