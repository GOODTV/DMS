<!--#include file="../_function/dbfunction.asp"-->
<%
For I = 1 to session("FieldsetCount")
  if Request("user_id"&I)<>"" then
     Update_data()
  end if       
Next
Response.Redirect "prog_edit.asp"

sub Update_data()
   sql="insert into user_prog values ('"&Request("user_id"&I)&"','"&request("程式群組")&"','"&request("prog_id")&"')"
   set RS=Server.CreateObject("ADODB.Recordset")
   RS.Open sql,conn,1,3
end sub
%>
