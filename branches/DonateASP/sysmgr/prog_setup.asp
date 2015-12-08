<!--#include file="../include/dbfunction.asp"-->
<%
if request("action")="save" then
   For I = 1 to session("FieldsetCount")
     if Request("user_id"&I)<>"" then
        Update_data()
     end if       
   Next
   Response.Redirect "prog_edit.asp?prog_id="&request("prog_id")
end if   

sub Update_data()
   sql="insert into user_prog values ('"&Request("user_id"&I)&"','"&request("menu_id")&"','"&request("prog_id")&"')"
   set RS=Server.CreateObject("ADODB.Recordset")
   RS.Open sql,conn,1,3
end sub

sql="select menu_id from prog_menu where progid='"&request("prog_id")&"'" 
Set RS1 = Server.CreateObject("ADODB.RecordSet") 
RS1.Open SQL,Conn,1,1 
if Not RS1.EOF then 
   menu_id=RS1("menu_id") 
else 
   menu_id=""    
end if 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>程式群組管理--資料輸入</title>
</head>
<BODY class=tool background="../images/logonbg.gif">
<p>
<form name="form" action="" method="post">
  <input type="hidden" name="action"> 
  <input type="hidden" name="AddButton" value="menu_add.asp">
<table border="0" width="80%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="font11" align="center">
        <font color="#800000">程式管理 </font><span style="font-family: 新細明體">
        <font size="2">【</font></span><span lang="zh-tw"><font size="2">帳號</font></span><span lang="zh-tw" style="font-family: 新細明體"><font size="2">設定</font></span><span style="font-family: 新細明體"><font size="2">】</font></span></td>
  </tr>
  <tr>
    <td width="100%">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td width="100%">
                    <div align="center">
                      <center>
                    <table width="80%" border=1 cellspacing="0" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
                        <font color="#000080">
                        <span lang="zh-tw">程式代號</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <input class=font9 size=20 readonly name="prog_id" value="<%=request("prog_id")%>">
          <input type="button" value=" 存檔 " name="save" class="cbutton">
                          <input type="button" value=" 取消 " name="cancel" class="cbutton"></td>
                      </tr>
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
                        程式名稱<font color="#000080">：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <input class=font9 size=40 readonly name="prog_desc" value="<%=request("prog_desc")%>"></td>
                      </tr>
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
                        程式群組<font color="#000080">：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                        <%SQL="select menu_id from menu order by menu_seq"
                          FName="menu_id"
                          Listfield="menu_id"
                          menusize="1"
                          BoundColumn=menu_id
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                         %>
                      　</td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="2">
                        <% SQL="select User_ID,帳號=User_ID,帳號名稱=user_name,帳號類別=user_group from userfile where dept_id='"&session("dept_id")&"' and user_id NOT IN (select user_id from user_prog where prog_id='"&request("prog_id")&"')"_
       +"                  order by user_id"      
                           FName="user_id"       
                           DataName="" 
                           DataNum=0    
                           call CheckBoxData (SQL,FName,DataName,DataNum,StrChecked)       
                         %>       
                        </td>
                        </table>
                      </center>
                    </div>
      </td>
  </tr>
</table>

<%message%>           
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
           
sub cancel_OnClick
  history.back
end sub

--></script>