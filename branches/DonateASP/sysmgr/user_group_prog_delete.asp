<!--#include file="../include/dbfunction.asp"-->
<% 
'sql="update group_prog set user_right='"&request("user_right")&"' where ser_no= '"&request("ser_no")&"' "
sql="delete from group_prog where ser_no= '"&request("ser_no")&"'"
call UpdateSQL (SQL)

Response.Redirect "user_group_edit.asp?menu_id="&session("menuid")
%>