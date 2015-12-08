<!--#include file="../include/dbfunction.asp"-->
<% 
sql="select * from user_prog where ser_no='"&request("ser_no")&"'"
Call QuerySQL(SQL,RS1)
if Not RS1.EOF then
   dept_id=RS1("dept_id")
   user_id=RS1("user_id")
   menu_id=RS1("menu_id")
end if   
sql="delete from user_prog where ser_no= '"&request("ser_no")&"'"
call UpdateSQL (SQL)

Response.Redirect "userfile_edit.asp?uid="&dept_id+user_id&"&menuid="&menu_id
%>