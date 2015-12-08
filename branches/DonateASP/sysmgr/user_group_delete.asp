<!--#include file="../include/dbfunction.asp"-->
<% 
   sql="delete from groupfile where user_group='"&request("usergroup")&"'"_
+"      delete from group_menu where user_group = '"&request("usergroup")&"'"_
+"      delete from group_prog where user_group = '"&request("usergroup")&"'"
   call UpdateSQL (SQL)

Response.Redirect "user_group.asp"
%>