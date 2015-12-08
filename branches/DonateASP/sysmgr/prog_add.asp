<!--#include file="../include/dbfunction.asp"-->
<%
if request("action")="save" then
   sql="select * from prog where prog_id='"&request("prog_id")&"'"
   Set RS1 = Server.CreateObject("ADODB.RecordSet")
   RS1.Open SQL,Conn,1,1
   if Not RS1.EOF then
      session("errnumber")=1
      session("msg")="資料重覆 ! "&request("prog_id")&" 程式資料已存在"
   else
      sql="insert into prog values ('"&request("prog_id")&"','"&request("prog_desc")&"','"&request("prog_url")&"','"&request("prog_seq")&"','"&request("prog_type")&"','"&request("security")&"')"_
+"         insert into prog_menu values ('"&request("menu_id")&"','"&request("prog_id")&"')"_
+"         insert into user_prog"_
+"           select user_id,menu_id,'"&request("prog_id")&"','A, R, U, D',dept_id from user_menu where menu_id='"&request("menu_id")&"'"
      call ExecSQL (SQL)
      Response.Redirect "prog.asp"
   end if
end if   
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>程式群組管理--資料輸入</title>
</head>
<BODY>
<table border="0" width="80%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"  height="25">
        程式管理【新增資料】</td>
  </tr>
  <tr>
    <td width="100%">
                  <form name="form" action="" method="post">
                  <input type="hidden" name="action"> 
                    <div align="center">
                      <center>
                    <table width="100%" border=1 cellspacing="1" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
                      <tr> 
                        <td noWrap align=right width="19%" bgcolor=#FFFFCC height=22>
                        <font color="#FF0000">*</font> <font color="#000080">
                        <span lang="zh-tw">程式代號</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <input class=font9 size=20 
                        name="prog_id" maxlength="40"></td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" height=22 bgcolor="#FFFFCC">
                        <font color="#FF0000">*</font>
                        <font color="#000080"><span lang="zh-tw">程式名稱</span>：</font></td>
                        <td width="80%" height=22 bgcolor="#FFFCE1"> 
                          <input class=font9 
                        size=40 name=prog_desc maxlength="40"></td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" bgcolor=#FFFFCC height=22 class="font9" style="background-color: #FFFFCC">
                        <font color="#000080"><span lang="zh-tw">連結URL</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <input class=font9 
                        size=40 name=prog_url maxlength="100"></td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" height=22 bgcolor="#FFFFCC" class="font9" style="background-color: #FFFFCC">
                        <font color="#000080"><span lang="zh-tw">排序</span>：</font></td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1"> 
                          <input type="text" name="prog_seq" size="4" class="font9" maxlength="3"></td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" height=22 class="td1">
                        <font color="#FF0000">*</font> 程式群組：</td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1"> 
          				      <%SQL="select menu_id,menu_seq='【'+menu_seq+'】'+menu_id from menu order by menu_seq"
          				        FName="menu_id"
            				      Listfield="menu_seq"
            				      menusize="1"
            				      BoundColumn=""
            				      call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
          					    %>
                          　</td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" height=22 bgcolor="#FFFFCC" class="font9" style="background-color: #FFFFCC">
                        <font color="#000080">程式類別：</font></td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1"> 
                          <select size="1" name="prog_type" class="font9A">
              <option selected value="ASP">ASP</option>
              <option value="QRY">QRY</option>
              <option value="RPT">RPT</option>
            </select></td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" bgcolor=#FFFFCC height=22 class="font9" style="background-color: #FFFFCC">
                        <font color="#000080"><span lang="zh-tw">安全控管</span>：</font></td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1">
                          <input type="text" name="security" size="4" class="font9" maxlength="1"></td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="2">
                         <%SaveButton%>　
                        </td>
                        </table>
                      </center>
                    </div>
                  </form>
      </td>
  </tr>
</table>

<%message%>           

</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--

sub save_OnClick
  <%call CheckString("prog_id","程式代號")%>
  <%call CheckString("prog_desc","程式名稱")%>
  <%call CheckString("menu_id","程式群組")%>
  <%call Checklen("prog_desc",40,"程式名稱")%>
  form.action.value="save"
  form.submit   
end sub

sub cancel_OnClick
  history.back
end sub

sub query_OnClick
  location.href="prog.asp"
end sub  

--></script>