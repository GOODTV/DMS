<!--#include file="../include/dbfunction.asp"-->
<% 
sql="select * from user_menu where ser_no='"&request("ser_no")&"'"
Call QuerySQL(SQL,RS1)
if Not RS1.EOF then
   dept_id=RS1("dept_id")
   user_id=RS1("user_id")
   menu_id=RS1("menu_id")
end if
sql="delete from user_prog where dept_id='"&dept_id&"' and user_id='"&user_id&"' and menu_id='"&menu_id&"'"_
+"   delete from user_menu where ser_no='"&request("ser_no")&"'"
Set RS=Conn.Execute(SQL)
Response.Redirect "userfile_edit.asp?uid="&dept_id+user_id
%>