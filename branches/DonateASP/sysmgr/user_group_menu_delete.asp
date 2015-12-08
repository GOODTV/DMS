<!--#include file="../include/dbfunction.asp"-->
<% 
   sql="select user_group,menu_id from group_menu where ser_no = '"&request("ser_no")&"'"
   Set RS1 = Server.CreateObject("ADODB.RecordSet")
   RS1.Open SQL,Conn,1,1
   if Not RS1.EOF then
      user_group=RS1("user_group")
      menu_id=RS1("menu_id")
   end if   
   sql="delete from group_menu where ser_no = '"&request("ser_no")&"'"_
+"      delete from group_prog where user_group='"&user_group&"' and menu_id='"&menu_id&"'"   
   Set RS=Conn.Execute(SQL)
   sql="select * from submenu where menu_id='"&menu_id&"'"
   Set RS2 = Server.CreateObject("ADODB.RecordSet")
   RS2.Open SQL,Conn,1,1
   While Not RS2.EOF
     sql="delete from group_prog where user_group='"&user_group&"' and menu_id='"&RS2("sub_menu")&"'"
     Set RS=Conn.Execute(SQL)
     RS2.MoveNext
   Wend
Response.Redirect "user_group_edit.asp?usergroup="&session("usergroup")
%>