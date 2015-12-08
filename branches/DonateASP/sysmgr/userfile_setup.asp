<!--#include file="../include/dbfunction.asp"-->
<%
if request("action")="save" then
   For I = 1 to session("FieldsetCount")
       if Request("prog_id"&I)<>"" then
          Update_data()
       end if       
   Next
   Response.Redirect "userfile_edit.asp?menuid="&request("menu_id")&"&uid="&request("uid")
end if   

sub Update_data()
   sql="insert into user_prog values ('"&request("userid")&"','"&request("menu_id")&"','"&Request("prog_id"&I)&"','"&request("user_right")&"','"&request("dept_id")&"')"
   set RS=Server.CreateObject("ADODB.Recordset")
   RS.Open sql,conn,1,3
end sub

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
  <input type="hidden" name="uid" value="<%=request("uid")%>">
<table border="0" width="80%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="font11" align="center">
        <font color="#800000">帳號管理 </font><span style="font-family: 新細明體">
        <font size="2">【<span lang="zh-tw">程式授權</span>】</font></span></td>
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
                        <span lang="zh-tw">帳號</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <input class=font9 size=20 readonly name="userid" value="<%=request("user_id")%>">
          <input type="button" value=" 存檔 " name="save" class="cbutton">
                          <input type="button" value=" 取消 " name="cancel" class="cbutton"></td>
                      </tr>
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
                        機構<font color="#000080">：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <%sql="select dept_id,comp_shortname from dept where dept_id='"&request("dept_id")&"'"
                          FName="dept_id"                                                                                                                                                                                                                                                              
                          Listfield="comp_shortname"                                                                                                                                                                                          
                          menusize="1"                                               
                          BoundColumn=request("dept_id")                                                                                                                                                                  
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                         %>
                          </td>
                      </tr>
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
                        程式群組<font color="#000080">：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <input class=font9 size=20 readonly name="menu_id" value="<%=request("menu_id")%>"></td>
                      </tr>
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
    			程式授權<font color="#000080">：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <input type="checkbox" name="user_right" checked value="A">新增 
                <input type="checkbox" name="user_right" checked value="R">查詢 
                <input type="checkbox" name="user_right" checked value="U">修改 
                <input type="checkbox" name="user_right" checked value="D">刪除
    			　</td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="2">
          <% SQL="select prog_id,程式代號=prog_id,程式名稱=prog_desc from prog where prog_id NOT IN (select prog_id from user_prog where dept_id+user_id='"&request("uid")&"'"_
       +"          and menu_id='"&request("menu_id")&"') and (security='' and '"&session("user_group")&"' <> '系統管理員' or '"&session("user_group")&"' = '系統管理員')"_
       +"          order by prog_seq"      
             FName="prog_id"       
             DataName="" 
             DataNum=0    
             call CheckBoxData (SQL,FName,DataName,DataNum,StrChecked)       
           %>                             </td>
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
  form.action.value="save"
  form.submit
end sub               
           
sub cancel_OnClick
  history.back
end sub

--></script>