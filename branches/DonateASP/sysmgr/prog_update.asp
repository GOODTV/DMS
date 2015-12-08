<!--#include file="../_function/dbfunction.asp"-->
<% 
if request("action") = "update" then
   sql="update prog set prog_desc='"&request("prog_desc")&"',prog_url='"&request("prog_url")&"',"_
+"      prog_seq='"&request("prog_seq")&"',prog_type='"&request("prog_type")&"',security='"&request("security")&"'"_
+"      where prog_id = '"&session("prog_id")&"' "
end if
if request("action") = "delete" then
   sql="delete from prog where prog_id = '"&session("prog_id")&"'"_
+"      delete from prog_menu where progid='"&session("prog_id")&"'"_
+"      delete from user_prog where prog_id='"&session("prog_id")&"'"   
end if
call UpdateSQL (SQL)

Response.Redirect("prog.asp")
%>