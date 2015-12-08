<!--#include file="../_function/dbfunction.asp"-->
<% 
if request("action") = "update" then
   sql="update menu set menu_seq='"&request("menu_seq")&"',"_
+"      icon_url='"&request("icon_url")&"',security='"&request("security")&"'"_
+"      where 程式群組 = '"&request("prog_group")&"' "
   call UpdateSQL (SQL)
   Response.Redirect("menu.asp")
end if

if request("action") = "delete" then
   sql="delete from menu where 程式群組 = '"&request("prog_group")&"' "_
+"      delete from submenu where 程式群組 = '"&request("prog_group")&"'"_
+"      delete from prog_menu where menu_id = '"&request("prog_group")&"'"_
+"      delete from user_menu where 程式群組 = '"&request("prog_group")&"'"   
   call UpdateSQL (SQL)
   Response.Redirect("menu.asp")
end if

if request("action")="add" then
   sql="Declare @seq Char(2)"_
+"      select @seq=menu_seq from menu where 程式群組='"&request("sub_menu")&"'"_   
+"      insert into submenu values ('"&request("sub_menu")&"','"&request("prog_group")&"',@seq)"_
+"      update menu set submenu_no=submenu_no+1 where 程式群組='"&request("prog_group")&"' "   
   call UpdateSQL (SQL)
   Response.Redirect("menu_edit.asp")
end if   
   
%>