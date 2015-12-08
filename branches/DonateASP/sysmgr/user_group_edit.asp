<!--#include file="../include/dbfunction.asp"-->
<%
if request("action")="save" then
   sql="insert into group_menu values ('"&request("usergroup")&"','"&request("menu_id")&"')"_
+"      insert into group_prog"_
+"        select '"&request("usergroup")&"','"&request("menu_id")&"',progid,'"&request("user_right")&"' from prog_menu where menu_id='"&request("menu_id")&"'"   
   call ExecSQL (SQL)
   sql="select * from submenu where menu_id='"&request("menu_id")&"'"
   Set RS2 = Server.CreateObject("ADODB.RecordSet")
   RS2.Open SQL,Conn,1,1
   While Not RS2.EOF
     sql="insert into group_prog"_
+"          select '"&request("usergroup")&"',menu_id,progid,'"&all_prog&"' from prog_menu where menu_id='"&RS2("sub_menu")&"'"
     call ExecSQL (SQL)
     RS2.MoveNext
   Wend
end if   

if request("usergroup")<>"" then
   session("usergroup")=request("usergroup")
end if
if request("menu_id")<>"" then
   menu_id=request("menu_id")
   session("menuid")=request("menu_id")
else   
  SQL="select group_menu.menu_id from group_menu join menu on group_menu.menu_id=menu.menu_id where user_group='"&session("usergroup")&"' order by menu.menu_seq"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  if Not RS1.EOF then 
     menu_id=RS1("menu_id")
  end if   
  RS1.close  
end if

sql="select * from groupfile where user_group='"&session("usergroup")&"'"
call QuerySQL (SQL,RS)

Function DataGrids (SQL,HLink,LinkParam,LinkTarget)
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL,Conn,1,1
    FieldsCount = RS1.Fields.Count-1
    Dim i
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%' align='center'>"
    Response.Write "<tr>"
    For i = 1 To FieldsCount
	Response.Write "<td bgcolor='#808080'><font color='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i).Name & "</span></font></td>"
    Next
    Response.Write "<td bgcolor='#808080'><font color='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>處理</span></font></td>"
    Response.Write "</tr>"
    While Not RS1.EOF
	Response.Write "<tr bgcolor='FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
	For i = 1 To FieldsCount
	  if RS1(3)="Y" then
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i) & "</span></td>"
      else
          Response.Write "<td><font color='#FF0000'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i) & "</span></font></td>"
      end if    
	Next
    if RS1(3)="Y" then
       Response.Write "<td><a href='user_group_prog_delete.asp?user_right=N&ser_no="&RS1("ser_no")&"'><img src='../images/x5.gif' border=0 width='16' height='14' alt='不授權'></a></td>"    
    else
       Response.Write "<td><a href='user_group_prog_delete.asp?user_right=Y&ser_no="&RS1("ser_no")&"'><img src='../images/gnicok.gif' border=0 width='16' height='14' alt='授權'></a></td>"    
    end if
	RS1.MoveNext
	Response.Write "</tr>"
    Wend
    Response.Write "</table>"
    RS1.Close
End Function

Function GroupGrid (SQL,HLink,LinkParam,LinkTarget)
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL,Conn,1,1
    FieldsCount = RS1.Fields.Count-1
    Dim i
    Response.Write "<table id=grid border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For i = 1 To FieldsCount
	Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i).Name & "</span></font></td>"
    Next
    Response.Write "</tr>"
    While Not RS1.EOF
	Response.Write "<tr bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
	For i = 1 To FieldsCount
	    if i = 1 then
	      Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>" & RS1(i) & "</a></span></td>"
	    else
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i) & "</span></td>"
	    end if   
       Next    
       RS1.MoveNext
       Response.Write "</tr>"
    Wend
    Response.Write "</table>"
    RS1.Close
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
<link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</HEAD>
<BODY>
<p>
<form name="form" action="" method="post">
<input type="hidden" name="action">
<div align="center">
<table border="0" width="90%" cellspacing="1" cellpadding="0">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"  height="25">
        帳號權限設定【修改資料】</td>
  </tr>
  <tr>
    <td width="100%" align="right" height="25">
      <a href="user_group.asp">查詢帳號類別</a></td>
  </tr>
  <tr>
    <td width="100%">
                    <div align="center">
                      <center>
                    <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
                        <font color="#000080">
                        帳號類別：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <input class=font9 size=20 readonly name="usergroup" value="<%=RS("user_group")%>"> </td>
                      </tr>
                        </table>
                      </center>
                    </div>
<tr>
    <td width="100%" align="center" height="10">
      </td>
  </tr>
  <tr>
    <td width="100%" align="center" height="10">
      <div align="center">
        <center>
        <table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#6487DC" width="100%" id="AutoNumber1">
          <tr>
            <td width="50%" valign="top">
            <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2" height="49">
              <tr>
    			<td width="50%" align="center" bgcolor="#808080" height="20">
      			<font color="#FFFFFF" size="2">程式群組授權</font><span style="font-size: 11pt"><font color="#800000"> </font> </span>&nbsp;</td>
              </tr>
              <tr>
    		  <td width="50%" height="19">程式群組
     			<%SQL="select menu.menu_id from menu where menu.menu_id not in (select group_menu.menu_id from group_menu where user_group='"&session("usergroup")&"')"_
          +"       order by menu_seq"
            		FName="menu_id"
            		Listfield="menu_id"
            		menusize="1"
            		BoundColumn=""
            		call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
           		%>     
      <input type="button" value=" 設 定 " name="save" class="cbutton"></td>
              	</tr>
              	<tr>
    			<td width="50%" height="19">
    			程式授權<font color="#000080">：</font><input type="checkbox" name="user_right" checked value="A">新增 
                <input type="checkbox" name="user_right" checked value="R">查詢 
                <input type="checkbox" name="user_right" checked value="U">修改 
                <input type="checkbox" name="user_right" checked value="D">刪除
    			　</td>
              	</tr>
              <tr>
    <td width="50%" valign="top" height="12">
    <% 
      SQL="select ser_no,group_menu.menu_id,程式群組=group_menu.menu_id from group_menu join menu on group_menu.menu_id=menu.menu_id where user_group='"&session("usergroup")&"' order by menu.menu_seq"
      HLink="user_group_edit.asp?menu_id=" 
      LinkParam="menu_id"
      LinkTarget="_self"
      DataLink="user_group_menu_delete.asp?ser_no=" 
      DataParam="ser_no"
      DataTarget="_self"
      call DataLinkGrid (SQL,HLink,LinkParam,LinkTarget,DataLink,DataParam,DataTarget)
      SQL="select submenu.sub_menu,次群組=sub_menu,程式群組=submenu.menu_id from submenu join group_menu on submenu.menu_id=group_menu.menu_id"_
+"           join menu on submenu.menu_id=menu.menu_id where user_group='"&session("usergroup")&"' order by menu.menu_seq"       
      HLink="user_group_edit.asp?menu_id=" 
      LinkParam="sub_menu"
      LinkTarget="_self"
      call GroupGrid (SQL,HLink,LinkParam,LinkTarget)
%> </td>
              </tr>
            </table>
            </td>
            <td width="50%" valign="top">
            <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
              <tr>
    <td width="50%" align="center" bgcolor="#808080" height="20">
      <font color="#FFFFFF" size="2">程式授權</font></td>
              </tr>
              <tr>
    <td width="50%">
      <input type="text" name="menuid" size="12" readonly class="font9t" value="<%=menu_id%>"></td>
              </tr>
              <tr>
                <td width="100%">
    <% 
     SQL="select ser_no,程式代號=group_prog.prog_id,程式名稱=prog_desc,授權=user_right from group_prog inner join prog on group_prog.prog_id=prog.prog_id"_
+"        where user_group='"&session("usergroup")&"' and menu_id='"&menu_id&"' order by prog.prog_seq" 
     HLink="user_group_prog_edit.asp?ser_no=" 
     LinkParam="ser_no"
     LinkTarget="_self"
     LinkType="window"
     call LinkGrid (SQL,HLink,LinkParam,LinkTarget,LinkType) 
    %> 
                　</td>
              </tr>
            </table>
            </td>
          </tr>
        </table>
<%message%>           
        </center>
      </div>
      </td>
  </tr>
      </td>
  </tr>
</table>
</div>
</form>
</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--

sub save_OnClick
  <%call CheckString("menu_id","程式群組")%>
  form.action.value="save"
  form.submit   
end sub

--></script>