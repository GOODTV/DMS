<!--#include file="../include/dbfunction.asp"-->
<%
if request("action")="update" then
   SQL="update group_prog set user_right='"&request("user_right")&"' where ser_no='"&request("ser_no")&"'"
   Call ExecSQL(SQL)  
   response.redirect "user_group_edit.asp?menu_id="&request("menu_id")
end if

if request("Action")="delete" then
   sql="delete from group_prog where ser_no='"&request("ser_no")&"'"
   Call ExecSQL(SQL)  
   response.redirect "user_group_edit.asp?menu_id="&request("menu_id")
end if

sql="select group_prog.*,prog_desc from group_prog join prog on group_prog.prog_id=prog.prog_id where ser_no='"&request("ser_no")&"'"
Call QuerySQL(SQL,RS)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>修改程式授權</title>
</head>
<BODY class=gray>
<form name="form" action method="post" target="main">
<input type="hidden" name="action">
<input type="hidden" name="ser_no" value="<%=rs("ser_no")%>">
<div align="center">
  <center>
<table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25">
        <font color="#ffffff">修改程式授權</font></td>
  </tr>

  <tr>
    <td width="100%">
                    <table width="100%" border=0 cellspacing="0" cellpadding="2">
                      <tr> 
                        <td align=right width="20%" height=22>
                        <font color="#000080">帳號類別：</font></td>
                        <td valign=center width="80%" height=18> 
                          <input type="text" name="user_group" size="12" class="font9t" readonly value="<%=rs("user_group")%>"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22>
                        <font color="#000080">程式群組：</font></td>
                        <td valign=center width="80%" height=18> 
                          <input type="text" name="menu_id" size="12" class="font9t" readonly value="<%=rs("menu_id")%>"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22>
                        <font color="#000080">程式代號：</font></td>
                        <td valign=center width="80%" height=18> 
                          <input type="text" name="prog_id" size="12" class="font9t" readonly value="<%=rs("prog_id")%>"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22>
                        <font color="#000080">程式名稱：</font></td>
                        <td valign=center width="80%" height=18> 
                          <input type="text" name="prog_desc" size="30" class="font9t" readonly value="<%=rs("prog_desc")%>"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22>
                        <font color="#000080">程式授權：</font></td>
                        <td valign=center width="80%" height=18> 
                        <input type="checkbox" name="user_right" value="A" <%If rs("user_right")="A" Then%>Checked<%End If%>>新增 
                        <input type="checkbox" name="user_right" value="R" <%If rs("user_right")="R" Then%>Checked<%End If%>>查詢 
                        <input type="checkbox" name="user_right" value="U" <%If rs("user_right")="U" Then%>Checked<%End If%>>修改 
                        <input type="checkbox" name="user_right" value="D" <%If rs("user_right")="D" Then%>Checked<%End If%>>刪除
    			</td>
                      </tr>
                      <tr>
                        <td width="100%" height=15 colspan="2" align="center">
                         </td>
                      </tr>
                      <tr>
                        <td width="100%" height=15 colspan="2" align="center">
                         <%EditButton%></td>
                      </tr>
                   </table>
      </td>
  </tr>
</table>

  </center>
</div>

<%message%>           
</form>
</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<script language="VBScript"><!--
sub cancel_OnClick
  window.close
end sub

sub update_OnClick
  form.action.value="update"
  form.submit
  window.close
end sub

sub delete_OnClick
  DIM answer
   answer = window.confirm ( "您是否確定要刪除 ?")
   if answer = true then
      form.action.value="delete"
      form.submit
      window.close
   end if
end sub

--></script>